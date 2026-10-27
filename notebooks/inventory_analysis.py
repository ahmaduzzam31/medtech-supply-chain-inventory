# %% [markdown]
# # Medtech Component Supply Chain & Inventory Optimisation
#
# **Company:** Corrib Medical Devices, a *fictional* Galway manufacturer of patient monitors (PM-200), wearable ECG
# patches (ECG-W1) and infusion pumps (IP-50). **All data is simulated** ([`src/generate_data.py`](../src/generate_data.py)).
#
# **Business question:** the operations director reports *"we have too much stock and we still run out of chips"*.
# Purchasing is also **€3.2M over budget**. The questions are:
# 1. Which components need the most control (ABC/XYZ)?
# 2. What should safety stock, reorder point and order quantity be for each part?
# 3. How much cash is tied up in excess stock, and which parts put production at risk?
# 4. Which suppliers are hurting us?
# 5. Why are we over budget: price or volume?
#
# The **main deliverable is the Excel model** ([`excel/Corrib_Medtech_Inventory_Model.xlsx`](../excel/)) with
# 26,000+ live formulas. This notebook recomputes the same logic independently in Python to **validate the Excel model**
# and to produce the charts.

# %%
import sys
from pathlib import Path

import duckdb
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats

ROOT = Path.cwd().parent if Path.cwd().name == "notebooks" else Path.cwd()
sys.path.insert(0, str(ROOT / "src"))
import style  # noqa: E402

style.apply()
FIG = ROOT / "reports/figures"
FIG.mkdir(parents=True, exist_ok=True)
pd.set_option("display.width", 150)

# Same assumptions as the Excel 'Assumptions' sheet
ORDER_COST, HOLD_RATE, DPM, MIN_COVER = 85, 0.22, 30.4, 1
SERVICE = {"A": 0.98, "B": 0.95, "C": 0.90}
A_THR, B_THR, X_THR, Y_THR = 0.80, 0.95, 0.25, 0.50

D = ROOT / "data"
items = pd.read_csv(D / "items.csv").set_index("sku")
suppliers = pd.read_csv(D / "suppliers.csv").set_index("supplier_id")
demand = pd.read_csv(D / "demand_monthly.csv")
pos = pd.read_csv(D / "purchase_orders.csv", parse_dates=["order_date", "promised_date", "received_date"])
stock = pd.read_csv(D / "inventory_snapshot.csv", parse_dates=["oldest_lot_expiry"]).set_index("sku")
budget = pd.read_csv(D / "budget_monthly.csv")
plan = pd.read_csv(D / "production_plan.csv")

print(f"{len(items)} components | {len(suppliers)} suppliers | {len(pos):,} PO lines | "
      f"{demand.month.nunique()} months of consumption")

# %% [markdown]
# ## 1. How demand is generated: production plan × bill of materials
# Component demand is driven by **how many devices are built** (with an August shutdown and fast growth on
# the wearable ECG patch), multiplied through the BOM, plus scrap and service spares.

# %%
pp = plan.pivot(index="month", columns="product", values="units_built")
fig, ax = plt.subplots(figsize=(9, 3.5))
for p, c in zip(["ECG-W1", "PM-200", "IP-50"], [style.BLUE, style.ORANGE, style.AQUA]):
    ax.plot(pp.index, pp[p], color=c, label=p)
ax.set_xticks(pp.index[::3]); ax.set_title("Units built per month by product")
ax.set_ylabel("Units"); ax.legend(ncol=3, loc="upper left")
style.save(fig, FIG / "01_production_plan.png")
pp.tail(3)

# %% [markdown]
# ## 2. ABC / XYZ classification
# * **ABC** ranks parts by last-12-month spend (Pareto): A = first 80% of spend, B = next 15%, C = the rest.
# * **XYZ** measures demand stability with the coefficient of variation (CV): X ≤ 0.25, Y ≤ 0.5, Z > 0.5.

# %%
d12 = demand[demand.month >= "2025-10"].pivot(index="sku", columns="month", values="qty_consumed")
pol = items[["description", "category", "supplier_id", "unit_cost_eur", "moq", "quoted_lead_time_days"]].copy()
pol["annual_usage"] = d12.sum(axis=1)
pol["avg_month"] = d12.mean(axis=1)
pol["sd_month"] = d12.std(axis=1)
pol["cv"] = pol.sd_month / pol.avg_month
pol["annual_spend"] = pol.annual_usage * pol.unit_cost_eur
pol = pol.sort_values("annual_spend", ascending=False)
pol["cum_spend_pct"] = pol.annual_spend.cumsum() / pol.annual_spend.sum()
pol["abc"] = np.where(pol.cum_spend_pct <= A_THR, "A", np.where(pol.cum_spend_pct <= B_THR, "B", "C"))
pol["xyz"] = np.where(pol.cv <= X_THR, "X", np.where(pol.cv <= Y_THR, "Y", "Z"))

matrix = pd.crosstab(pol.abc, pol.xyz, margins=True)
fig, ax = plt.subplots(figsize=(8, 3.8))
x = np.arange(1, len(pol) + 1) / len(pol) * 100
ax.plot(x, pol.cum_spend_pct * 100, color=style.BLUE)
for thr, cls in [(A_THR, "A"), (B_THR, "B")]:
    n = (pol.cum_spend_pct <= thr).sum()
    ax.axvline(n / len(pol) * 100, color=style.MUTED, lw=0.8, ls="--")
    ax.text(n / len(pol) * 100 + 1, 10, f"{cls} | {n / len(pol):.0%} of SKUs", color=style.INK_2, fontsize=9)
ax.set_title("Pareto: 25% of components make up 80% of spend")
ax.set_xlabel("% of components (ranked by spend)"); ax.set_ylabel("Cumulative % of spend")
style.save(fig, FIG / "02_abc_pareto.png")
matrix

# %% [markdown]
# **Takeaway:** **29 A-class parts (25% of SKUs) carry 80% of spend.** They are mostly microcontrollers,
# sensors and PCBs. They deserve weekly review and tight safety stock, while the 59 C-class parts (half the SKUs,
# 5% of spend) can run on simple min/max rules.

# %% [markdown]
# ## 3. Inventory policy: safety stock, reorder point, order quantity
# Safety stock uses **actual** lead times from purchase-order history, not supplier quotes, and accounts for both
# demand variability and lead-time variability:
#
# $$SS = z \sqrt{\tfrac{LT}{30.4}\,\sigma_d^2 + \bar d^2 \left(\tfrac{\sigma_{LT}}{30.4}\right)^2}
# \qquad ROP = \bar d \cdot \tfrac{LT}{30.4} + SS \qquad Q = \max(EOQ,\ MOQ,\ 1\text{ month})$$

# %%
rec = pos[pos.status == "Received"].copy()
rec["lt"] = (rec.received_date - rec.order_date).dt.days
pol["lt_mean"] = rec.groupby("sku").lt.mean()
pol["lt_sd"] = rec.groupby("sku").lt.std().fillna(0)
pol["z"] = pol.abc.map({k: stats.norm.ppf(v) for k, v in SERVICE.items()})
pol["safety_stock"] = np.ceil(pol.z * np.sqrt(pol.lt_mean / DPM * pol.sd_month ** 2
                                              + pol.avg_month ** 2 * (pol.lt_sd / DPM) ** 2))
pol["rop"] = np.ceil(pol.avg_month * pol.lt_mean / DPM + pol.safety_stock)
eoq = np.sqrt(2 * pol.annual_usage * ORDER_COST / (HOLD_RATE * pol.unit_cost_eur))
pol["order_qty"] = np.ceil(np.maximum.reduce([pol.moq, eoq, pol.avg_month * MIN_COVER]))
pol = pol.join(stock[["on_hand_qty", "on_order_qty", "oldest_lot_expiry"]])
pol["position"] = pol.on_hand_qty + pol.on_order_qty
pol["max_stock"] = pol.safety_stock + pol.order_qty
pol["status"] = np.select([pol.on_hand_qty < pol.safety_stock, pol.position <= pol.rop, pol.on_hand_qty > pol.max_stock],
                          ["Below safety stock", "Reorder now", "Excess"], "Healthy")
pol["stock_value"] = pol.on_hand_qty * pol.unit_cost_eur
pol["excess_value"] = (pol.on_hand_qty - pol.max_stock).clip(lower=0) * pol.unit_cost_eur
pol["shortfall_value"] = (pol.rop - pol.position).clip(lower=0) * pol.unit_cost_eur
pol["days_cover"] = pol.on_hand_qty / (pol.avg_month / DPM)

status = pol.groupby("status").agg(skus=("status", "size"), stock_value=("stock_value", "sum"),
                                   excess_value=("excess_value", "sum")).reindex(
    ["Below safety stock", "Reorder now", "Healthy", "Excess"])
status

# %%
fig, axes = plt.subplots(1, 2, figsize=(12, 3.8))
colors = [style.CRITICAL, style.WARNING, style.GOOD, style.MUTED]
ax = axes[0]
ax.barh(status.index[::-1], status.skus[::-1], color=colors[::-1], height=0.55)
for i, v in enumerate(status.skus[::-1]):
    ax.text(v + 0.8, i, f"{v}", va="center", color=style.INK_2)
ax.set_title("Components by stock status"); ax.set_xlabel("Number of SKUs")
ax.grid(axis="x"); ax.grid(axis="y", visible=False)
ax = axes[1]
ax.barh(status.index[::-1], status.stock_value[::-1] / 1e6, color=colors[::-1], height=0.55)
for i, v in enumerate(status.stock_value[::-1] / 1e6):
    ax.text(v + 0.03, i, f"€{v:.2f}M", va="center", color=style.INK_2)
ax.set_title("Stock value by status"); ax.set_xlabel("€ millions")
ax.grid(axis="x"); ax.grid(axis="y", visible=False)
fig.tight_layout()
style.save(fig, FIG / "03_stock_status.png")

# %% [markdown]
# ### Production at risk: parts that cannot be replenished in time
# A part is critical when its **days of cover are shorter than its replenishment lead time**. Stock will run out
# before a new order can arrive unless it is expedited.

# %%
critical = pol[(pol.status == "Below safety stock")].copy()
critical["cover_minus_lt_days"] = critical.days_cover - critical.lt_mean
critical["supplier"] = critical.supplier_id.map(suppliers.supplier_name)
critical[["description", "abc", "supplier", "days_cover", "lt_mean", "cover_minus_lt_days", "shortfall_value"]].round(0) \
    .sort_values("cover_minus_lt_days")

# %% [markdown]
# **Takeaway:** every part below safety stock is a **chip or battery** from a long-lead-time Asian or US supplier, and
# most are A-class. Several have only a few weeks of cover against lead times of 3 to 4 months, so a single late delivery
# stops a production line. These need **expediting and allocation agreements now**.

# %% [markdown]
# ## 4. Excess stock and working capital

# %%
excess = pol[pol.excess_value > 0].sort_values("excess_value", ascending=False)
by_cat = excess.groupby("category").excess_value.sum().sort_values(ascending=False)
carrying = pol.excess_value.sum() * HOLD_RATE
expiring = pol[pol.oldest_lot_expiry.notna() & (pol.oldest_lot_expiry <= pd.Timestamp("2026-09-30") + pd.DateOffset(months=6))]
print(f"Excess stock: €{pol.excess_value.sum():,.0f} ({pol.excess_value.sum() / pol.stock_value.sum():.0%} of inventory)")
print(f"Annual carrying cost of excess at {HOLD_RATE:.0%}: €{carrying:,.0f}")
print(f"Lots expiring within 6 months: {len(expiring)} SKUs, €{expiring.stock_value.sum():,.0f} of stock")
by_cat.round(0).to_frame("excess_value_eur")

# %% [markdown]
# **Takeaway:** **€1.06M (21%) of inventory is excess**, costing about **€234k a year** to carry. Much of it is cheap
# passive components bought in bulk "just in case". Each part is cheap, but together they hold a lot of cash. The same
# company is short of expensive chips: **the money is in the wrong parts.**

# %% [markdown]
# ## 5. Supplier performance

# %%
con = duckdb.connect()
for t in ["items", "suppliers", "purchase_orders"]:
    con.execute(f"CREATE VIEW {t} AS SELECT * FROM read_csv_auto('{D / (t + '.csv')}')")
sql = lambda f: con.sql((ROOT / "sql" / f).read_text()).df()
otif = sql("01_supplier_otif.sql")
fy_spend = pos[pos.order_date >= "2025-10-01"].assign(v=lambda d: d.qty_ordered * d.unit_price_eur) \
    .groupby("supplier_id").v.sum()
otif["fy_spend"] = otif.supplier_name.map(suppliers.reset_index().set_index("supplier_name").supplier_id).map(fy_spend)

fig, ax = plt.subplots(figsize=(9, 4.5))
ax.scatter(otif.avg_lead_time_days, otif.otif_pct, s=otif.fy_spend / 25_000, color=style.BLUE, alpha=0.75,
           edgecolor=style.SURFACE, linewidth=2)
for _, r in otif.iterrows():
    ax.annotate(r.supplier_name.split()[0], (r.avg_lead_time_days, r.otif_pct), xytext=(8, -3),
                textcoords="offset points", fontsize=8.5, color=style.INK_2)
ax.axhline(85, color=style.CRITICAL, lw=0.8, ls="--")
ax.text(140, 85.6, "85% OTIF minimum", color=style.CRITICAL, fontsize=8.5, ha="right")
ax.set_xlim(0, 150)
ax.set_title("Supplier risk: long lead time + low OTIF (bubble size = FY spend)")
ax.set_xlabel("Average actual lead time (days)"); ax.set_ylabel("OTIF %")
style.save(fig, FIG / "04_supplier_risk.png")
otif

# %%
sql("04_single_source_exposure.sql")

# %% [markdown]
# **Takeaway:** the two worst suppliers are also the most dangerous ones:
# * **Formosa Microtech** has 59% OTIF on 137-day lead times for microcontrollers.
# * **Shenzhen Brightcell** has 64% OTIF and is the **single source for all six batteries** (€3.0M a year), with the
#   highest reject rate (1.8%). A battery shortage stops all three product lines.
#
# Medtech context: qualifying a second battery supplier takes time (design verification, biocompatibility and regulatory
# change control), so it should start now.

# %% [markdown]
# ## 6. Budget variance: price or volume?

# %%
fy = pos[pos.order_date >= "2025-10-01"].join(items[["category", "unit_cost_eur"]], on="sku")
fy["actual"] = fy.qty_ordered * fy.unit_price_eur
fy["price_var"] = (fy.unit_price_eur - fy.unit_cost_eur) * fy.qty_ordered
bud = budget.groupby("category").budget_eur.sum()
var = pd.DataFrame({"budget": bud, "actual": fy.groupby("category").actual.sum(),
                    "price_variance": fy.groupby("category").price_var.sum()})
var["total_variance"] = var.actual - var.budget
var["volume_timing"] = var.total_variance - var.price_variance
var = var.sort_values("total_variance", ascending=False)
tot = var.sum()

steps = [("FY budget", tot.budget), ("Chip price\nincrease", var.loc["Microcontrollers & ICs", "price_variance"]),
         ("Other price\nincreases", tot.price_variance - var.loc["Microcontrollers & ICs", "price_variance"]),
         ("Volume /\ntiming", tot.volume_timing), ("FY actual", tot.actual)]
fig, ax = plt.subplots(figsize=(9, 4))
running = 0
for i, (lab, v) in enumerate(steps):
    if i in (0, len(steps) - 1):
        ax.bar(i, v / 1e6, color=style.BLUE, width=0.55); running = v
        ax.text(i, v / 1e6 + 0.3, f"€{v / 1e6:.1f}M", ha="center", color=style.INK)
    else:
        ax.bar(i, v / 1e6, bottom=running / 1e6, color=style.CRITICAL if v > 0 else style.GOOD, width=0.55)
        ax.text(i, (running + v) / 1e6 + 0.3, f"+€{v / 1e6:.2f}M", ha="center", color=style.INK)
        running += v
ax.set_xticks(range(len(steps)), [s[0] for s in steps])
ax.set_ylim(30, 39); ax.set_ylabel("€ millions (axis starts at €30M)")
ax.set_title("Purchasing variance bridge: most of the overspend is price, not volume")
style.save(fig, FIG / "05_variance_bridge.png")
(var / 1000).round(0).rename(columns=lambda c: c + " (€k)")

# %% [markdown]
# **Takeaway:** purchasing is **€3.2M (9.4%) over budget**. **€2.5M is price variance**, of which **€1.9M comes from
# an 18% increase on microcontrollers and ICs** since July 2025. Volume and timing add only €0.7M. The budget used
# pre-shortage standard costs, so the fix is commercial (long-term agreements, second sources, updated standard
# costs), not a planning failure.

# %% [markdown]
# ## 7. Validation: does the Excel model match Python?
# [`src/reconcile.py`](../src/reconcile.py) recalculates the workbook in LibreOffice and compares every SKU with this
# notebook's numbers. Summary of the last run:

# %%
recon_file = ROOT / "reports/reconciliation.csv"
pd.read_csv(recon_file) if recon_file.exists() else "Run src/reconcile.py to produce reports/reconciliation.csv"

# %% [markdown]
# ## 8. Recommendations
#
# | # | Action | Impact | Owner |
# |---|---|---|---|
# | 1 | **Expedite the 9 parts below safety stock**, which are all chips and batteries | Protects production on all three product lines | Planning / Purchasing |
# | 2 | **Stop buying excess C-class parts.** Switch to min/max with the calculated reorder points and let €1.06M of excess run down | ≈€234k a year carrying cost, about €1M cash released | Planning |
# | 3 | **Qualify a second battery supplier** (start design verification and regulatory change control now) | Removes a single point of failure on €3.0M of spend | Engineering + Quality + Purchasing |
# | 4 | **Put Formosa Microtech and Shenzhen Brightcell on supplier improvement plans** with monthly OTIF reviews | OTIF 59% and 64% vs 85% minimum | Supplier quality |
# | 5 | **Negotiate long-term agreements on A-class ICs** and update standard costs for the next budget | Addresses €1.9M chip price variance | Purchasing / Finance |
# | 6 | **Plan with actual lead times, not quotes,** and review A-class parts weekly | More accurate reorder points | Planning |
