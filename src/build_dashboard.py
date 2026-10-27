"""Build the interactive inventory dashboard (docs/index.html, served by GitHub Pages).

Reads the policy results from the recalculated Excel model, so the web dashboard and the Excel file always agree.
"""
import shutil
import subprocess
import tempfile
from pathlib import Path

import duckdb
import pandas as pd
import plotly.graph_objects as go

from dashboard_kit import BLUE, CRITICAL, GOOD, INK_2, MUTED, ORANGE, WARNING, build_page

ROOT = Path(__file__).resolve().parents[1]
D = ROOT / "data"
MODEL = ROOT / "excel/Corrib_Medtech_Inventory_Model.xlsx"

with tempfile.TemporaryDirectory() as tmp:
    shutil.copy(MODEL, Path(tmp) / "m.xlsx")
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", f"{tmp}/out", f"{tmp}/m.xlsx"],
                   check=True, capture_output=True, timeout=600)
    xl = Path(tmp) / "out/m.xlsx"
    pol = pd.read_excel(xl, sheet_name="Inventory_Policy", header=3)
    score = pd.read_excel(xl, sheet_name="Supplier_Scorecard", header=3).dropna(subset=["Supplier"])
    bva = pd.read_excel(xl, sheet_name="Budget_vs_Actual", header=None)

con = duckdb.connect()
for t in ["items", "suppliers", "purchase_orders"]:
    con.execute(f"CREATE VIEW {t} AS SELECT * FROM read_csv_auto('{D / (t + '.csv')}')")
price = con.sql((ROOT / "sql/03_price_inflation.sql").read_text()).df()

# Stock status
order = ["Below safety stock", "Reorder now", "Healthy", "Excess"]
colors = {"Below safety stock": CRITICAL, "Reorder now": WARNING, "Healthy": GOOD, "Excess": MUTED}
st = pol.groupby("Stock status").agg(skus=("SKU", "size"), value=("Stock value (€)", "sum")).reindex(order)
f1 = go.Figure(go.Bar(x=st.value / 1e6, y=st.index, orientation="h", marker_color=[colors[s] for s in order],
                      customdata=st.skus, hovertemplate="%{y}<br>€%{x:.2f}M of stock<br>%{customdata} SKUs<extra></extra>"))
f1.update_layout(xaxis_title="Stock value (€M)", yaxis=dict(autorange="reversed"), bargap=0.35)

# Below safety stock list as a bar of days cover vs lead time
crit = pol[pol["Stock status"] == "Below safety stock"].sort_values("Days of cover")
f2 = go.Figure()
f2.add_bar(y=crit.Description, x=crit["Avg actual lead time (days)"], orientation="h", name="Lead time (days)",
           marker_color="#cde2fb", hovertemplate="%{y}<br>Lead time %{x:.0f} days<extra></extra>")
f2.add_bar(y=crit.Description, x=crit["Days of cover"], orientation="h", name="Days of cover", marker_color=CRITICAL,
           customdata=crit[["SKU", "Supplier", "ABC"]],
           hovertemplate="%{customdata[0]} (%{customdata[2]}-class, %{customdata[1]})<br>%{x:.0f} days of cover<extra></extra>")
f2.update_layout(barmode="overlay", xaxis_title="Days", yaxis=dict(autorange="reversed"), bargap=0.3)

# Supplier OTIF
sc = score.sort_values("OTIF %")
f3 = go.Figure(go.Bar(x=sc["OTIF %"] * 100, y=sc.Supplier, orientation="h",
                      marker_color=[CRITICAL if v < 0.85 else BLUE for v in sc["OTIF %"]],
                      customdata=sc[["Avg actual lead time (days)", "Reject rate", "Rating", "Single-source SKUs"]],
                      hovertemplate="%{y}<br>OTIF %{x:.1f}%<br>Lead time %{customdata[0]:.0f} days"
                      "<br>Rejects %{customdata[1]:.2%}<br>%{customdata[2]}"
                      "<br>Single-source SKUs: %{customdata[3]}<extra></extra>"))
f3.add_vline(x=85, line=dict(color=INK_2, width=1, dash="dash"), annotation_text="85% minimum", annotation_position="top")
f3.update_layout(xaxis=dict(title="OTIF %", range=[50, 100]), bargap=0.3)

# Price index
f4 = go.Figure()
other = (price[price.category != "Microcontrollers & ICs"].groupby("quarter", as_index=False)["price_index"].mean()
         .assign(category="All other categories (average)"))
price = pd.concat([price, other])
for cat, color in [("Microcontrollers & ICs", ORANGE), ("All other categories (average)", BLUE)]:
    d = price[price.category == cat].sort_values("quarter")
    f4.add_scatter(x=d.quarter, y=d.price_index, name=cat, mode="lines+markers", line=dict(color=color, width=2.5),
                   marker=dict(size=8), hovertemplate=cat + "<br>%{x}: index %{y:.1f}<extra></extra>")
f4.update_layout(hovermode="x unified", yaxis_title="Price index (standard cost = 100)")

# Variance bridge
i0 = bva.index[bva[0] == "FY variance analysis: price vs volume/timing"][0] + 2
br = bva.iloc[i0:i0 + 10, :8]
br.columns = ["category", "budget", "actual", "variance", "var_pct", "price", "volume", "driver"]
tot = br[br.category == "Total"].iloc[0]
chip = br[br.category == "Microcontrollers & ICs"].iloc[0].price
f5 = go.Figure(go.Waterfall(
    x=["FY budget", "Chip price increase", "Other price increases", "Volume / timing", "FY actual"],
    measure=["absolute", "relative", "relative", "relative", "total"],
    y=[tot.budget / 1e6, chip / 1e6, (tot.price - chip) / 1e6, tot.volume / 1e6, 0],
    text=[f"€{tot.budget / 1e6:.1f}M", f"+€{chip / 1e6:.2f}M", f"+€{(tot.price - chip) / 1e6:.2f}M",
          f"+€{tot.volume / 1e6:.2f}M", f"€{tot.actual / 1e6:.1f}M"], textposition="outside",
    increasing=dict(marker_color=CRITICAL), decreasing=dict(marker_color=GOOD), totals=dict(marker_color=BLUE),
    connector=dict(line=dict(color=MUTED, width=1)), hovertemplate="%{x}<br>€%{y:.2f}M<extra></extra>"))
f5.update_layout(yaxis=dict(title="€ millions (axis starts at €30M)", range=[30, 38.5]))

inv, exc = pol["Stock value (€)"].sum(), pol["Excess value (€)"].sum()
w_otif = (score["OTIF %"] * score["POs received"]).sum() / score["POs received"].sum()
build_page(
    title="Medtech Inventory & Supplier Dashboard",
    intro="Corrib Medical Devices, a <b>fictional</b> Galway maker of patient monitors, wearable ECG patches and infusion "
          "pumps. 116 components, 12 suppliers, 2,024 purchase orders. <b>All data is simulated.</b> "
          "Numbers come from the Excel model in the repository.",
    kpis=[("Inventory value", f"€{inv / 1e6:.2f}M", "snapshot 30 Sep 2026"),
          ("Excess stock", f"€{exc / 1e6:.2f}M", f"{exc / inv:.0%} of inventory"),
          ("Parts below safety stock", f"{(pol['Stock status'] == 'Below safety stock').sum()}", "all chips or batteries"),
          ("Supplier OTIF", f"{w_otif:.1%}", "weighted by PO count"),
          ("Purchasing vs budget", f"+€{tot.variance / 1e6:.2f}M", f"+{tot.var_pct:.1%}, mostly price")],
    cards=[
        {"title": "Stock value by status", "fig": f1, "sub": "€1.06M sits above maximum stock while critical chips run short."},
        {"title": "Parts below safety stock: cover vs lead time", "fig": f2, "height": 380,
         "sub": "Red = days of stock left. Light bar = lead time to replenish. Red shorter than light means a stockout without expediting."},
        {"title": "Supplier OTIF", "fig": f3, "height": 400,
         "sub": "Formosa (MCUs) and Shenzhen Brightcell (sole battery source) are far below 85%."},
        {"title": "Purchase price index by quarter", "fig": f4, "height": 400,
         "sub": "Microcontroller prices rose 18% from Q3 2025. Other categories rose 3% in 2026."},
        {"title": "Purchasing variance bridge (FY Oct-25 to Sep-26)", "fig": f5, "wide": True, "height": 360,
         "sub": "Of the €3.2M overspend, €2.5M is price. Volume and timing explain only €0.7M."},
    ],
    insights=[
        "<b>The money is in the wrong parts.</b> €1.06M of excess is mostly cheap passives and sensors, while 9 critical chips and batteries are below safety stock.",
        "<b>Production risk:</b> several A-class ICs have 1 to 4 weeks of cover against 3 to 5 month lead times. Expedite now.",
        "<b>Single point of failure:</b> one supplier at 64% OTIF provides every battery (€3.0M per year). Qualify a second source.",
        "<b>The overspend is price, not planning:</b> €1.9M comes from an 18% chip price increase. Negotiate long-term agreements and update standard costs.",
    ],
    footer="Simulated data for a fictional company. The model logic is in the Excel workbook and reconciled with Python (0 mismatches on 116 SKUs).",
    path=ROOT / "docs/index.html",
)
print("Dashboard written to docs/index.html")
