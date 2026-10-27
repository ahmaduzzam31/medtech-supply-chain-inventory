"""Generate the simulated dataset for Corrib Medical Devices (a fictional Galway medtech manufacturer).

Everything is synthetic but built to behave like a real component supply chain:
production plan -> bill of materials -> component demand, supplier-specific lead times and
reliability, price increases on semiconductors, batteries with shelf life, and an inventory
snapshot that mixes overstocked and understocked parts.

Outputs (data/):
    suppliers.csv, items.csv, bom.csv, production_plan.csv,
    demand_monthly.csv, purchase_orders.csv, inventory_snapshot.csv, budget_monthly.csv
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data"
MONTHS = pd.period_range("2024-10", "2026-09", freq="M")    # 24 months of history
SNAPSHOT_DATE = pd.Timestamp("2026-09-30")
rng = np.random.default_rng(SEED)

# --------------------------------------------------------------------------------------
# Suppliers: (id, name, country, lead-time multiplier, on-time probability, in-full prob, reject rate)
# --------------------------------------------------------------------------------------
SUPPLIERS = [
    ("S01", "Shannon Electronics Ltd", "Ireland", 0.8, 0.96, 0.98, 0.002),
    ("S02", "Rhein Sensorik GmbH", "Germany", 1.0, 0.93, 0.97, 0.004),
    ("S03", "Formosa Microtech Co.", "Taiwan", 1.6, 0.66, 0.90, 0.006),
    ("S04", "Penang Precision Sdn Bhd", "Malaysia", 1.3, 0.82, 0.95, 0.010),
    ("S05", "Shenzhen Brightcell Power", "China", 1.5, 0.74, 0.92, 0.018),
    ("S06", "Midlands Connect UK", "United Kingdom", 0.9, 0.91, 0.96, 0.005),
    ("S07", "Lakeview Circuits Inc.", "USA", 1.1, 0.88, 0.97, 0.008),
    ("S08", "Bavaria Display AG", "Germany", 1.2, 0.85, 0.96, 0.012),
    ("S09", "Atlantic Polymers Ltd", "Ireland", 0.7, 0.95, 0.99, 0.003),
    ("S10", "Nordic Passive Components", "Sweden", 0.9, 0.94, 0.98, 0.001),
    ("S11", "Kyushu Semiconductor KK", "Japan", 1.4, 0.90, 0.98, 0.002),
    ("S12", "Dublin Sterile Packaging", "Ireland", 0.6, 0.97, 0.99, 0.001),
]

PRODUCTS = {
    "PM-200": "Patient monitor PM-200",
    "ECG-W1": "Wearable ECG patch W1",
    "IP-50": "Infusion pump IP-50",
}

# Category: (supplier pool, base lead time days, unit cost range EUR, shelf life months or 0)
CATEGORIES = {
    "Microcontrollers & ICs": (["S03", "S11", "S07"], 84, (2.5, 38.0), 0),
    "Sensors":                (["S02", "S04", "S11"], 56, (1.8, 45.0), 0),
    "Printed circuit boards": (["S07", "S04", "S01"], 35, (4.0, 60.0), 0),
    "Batteries":              (["S05"],             70, (3.0, 28.0), 24),
    "Passive components":     (["S10", "S04", "S01"], 28, (0.01, 0.40), 0),
    "Connectors & cables":    (["S06", "S01"],      21, (0.30, 12.0), 0),
    "Displays":               (["S08"],             63, (18.0, 95.0), 0),
    "Enclosures & plastics":  (["S09"],             28, (0.8, 22.0), 0),
    "Sterile packaging":      (["S12"],             14, (0.15, 3.5), 36),
}
ITEMS_PER_CAT = {"Microcontrollers & ICs": 16, "Sensors": 14, "Printed circuit boards": 10, "Batteries": 6,
                 "Passive components": 30, "Connectors & cables": 14, "Displays": 5,
                 "Enclosures & plastics": 12, "Sterile packaging": 9}
DESCRIPTORS = {
    "Microcontrollers & ICs": ["ARM Cortex-M4 MCU", "ARM Cortex-M0+ MCU", "Bluetooth LE SoC", "ECG analog front end",
                               "Power management IC", "Motor driver IC", "16-bit ADC", "Op-amp quad",
                               "Flash memory 64Mb", "Isolated gate driver", "LDO regulator 3.3V", "Buck converter IC",
                               "Real-time clock IC", "Wi-Fi module", "Secure element IC", "USB-C PD controller"],
    "Sensors": ["SpO2 optical sensor", "Temperature sensor", "Pressure transducer", "Air-in-line ultrasonic sensor",
                "Occlusion force sensor", "3-axis accelerometer", "Flow sensor", "Hall effect sensor",
                "NIBP pressure sensor", "Ambient light sensor", "Bioimpedance sensor", "Humidity sensor",
                "Drop detector", "Gyroscope MEMS"],
    "Printed circuit boards": ["PM-200 main board", "PM-200 power board", "ECG-W1 flex PCB", "IP-50 control board",
                               "IP-50 motor board", "Display interface board", "Sensor daughter board",
                               "Charging dock PCB", "Comms module PCB", "LED indicator PCB"],
    "Batteries": ["Li-ion 3.7V 2000mAh pack", "Li-ion 3.7V 5000mAh pack", "Coin cell CR2032",
                  "LiPo 3.7V 180mAh pouch", "Backup NiMH 4.8V pack", "Li-ion 7.4V 3400mAh pack"],
    "Passive components": [f"{t} {v}" for t, v in [
        ("Resistor 0402", "10k"), ("Resistor 0402", "100R"), ("Resistor 0603", "4k7"), ("Resistor 0603", "1M"),
        ("Capacitor 0402", "100nF"), ("Capacitor 0402", "10pF"), ("Capacitor 0603", "1uF"), ("Capacitor 0805", "10uF"),
        ("Capacitor 1206", "47uF"), ("Inductor 0805", "2.2uH"), ("Inductor 1210", "10uH"), ("Ferrite bead 0603", "600R"),
        ("TVS diode", "SMBJ5.0"), ("Schottky diode", "SS34"), ("Zener diode", "3.3V"), ("LED 0603", "green"),
        ("LED 0603", "red"), ("Crystal", "32.768kHz"), ("Crystal", "16MHz"), ("MOSFET", "N-ch 30V"),
        ("MOSFET", "P-ch 20V"), ("Fuse 0603", "500mA"), ("Thermistor NTC", "10k"), ("Resistor 0402", "1k"),
        ("Capacitor 0402", "1nF"), ("Resistor 0603", "22R"), ("Capacitor 0603", "4.7uF"), ("Varistor", "14V"),
        ("BJT", "NPN SOT-23"), ("Resistor array", "4x10k")]],
    "Connectors & cables": ["USB-C receptacle", "ECG lead wire set", "SpO2 extension cable", "Board-to-board connector",
                            "FPC connector 24-pin", "DC barrel jack", "Mains power cord EU", "Mains power cord UK",
                            "Wire harness PM-200", "Wire harness IP-50", "Snap electrode connector", "Pogo pin connector",
                            "RJ45 jack", "Battery connector JST"],
    "Displays": ["10.1in TFT touch display", "7in TFT display", "2.4in OLED display", "Touch panel controller", "LED backlight unit"],
    "Enclosures & plastics": ["PM-200 front housing", "PM-200 rear housing", "IP-50 housing", "IP-50 door assembly",
                              "ECG-W1 top shell", "ECG-W1 bottom shell", "Silicone keypad", "Light pipe",
                              "Mounting bracket", "Gasket seal set", "Battery door", "Rubber feet set"],
    "Sterile packaging": ["Tyvek pouch small", "Tyvek pouch large", "Blister tray ECG-W1", "Shipper carton PM-200",
                          "Shipper carton IP-50", "IFU leaflet", "UDI label roll", "Desiccant pack", "Hydrogel electrode pad"],
}


def build_suppliers() -> pd.DataFrame:
    return pd.DataFrame(SUPPLIERS, columns=["supplier_id", "supplier_name", "country", "lt_multiplier",
                                            "p_on_time", "p_in_full", "reject_rate"])


def build_items(sup: pd.DataFrame) -> pd.DataFrame:
    rows, n = [], 1
    for cat, (pool, base_lt, (lo, hi), shelf) in CATEGORIES.items():
        for desc in DESCRIPTORS[cat][: ITEMS_PER_CAT[cat]]:
            supplier = pool[0] if len(pool) == 1 else rng.choice(pool, p=None)
            mult = sup.set_index("supplier_id").loc[supplier, "lt_multiplier"]
            cost = float(np.exp(rng.uniform(np.log(lo), np.log(hi))))
            rows.append({
                "sku": f"CMD-{n:04d}", "description": desc, "category": cat, "supplier_id": supplier,
                "unit_cost_eur": round(cost, 4 if cost < 1 else 2),
                "quoted_lead_time_days": int(round(base_lt * mult / 7) * 7) or 7,
                "moq": int(rng.choice([1, 50, 100, 250, 1000, 5000]) if cost < 1 else rng.choice([1, 10, 25, 50, 100])),
                "shelf_life_months": shelf,
                "single_source": len(pool) == 1,
            })
            n += 1
    return pd.DataFrame(rows)


def build_bom(items: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, it in items.iterrows():
        named = [p for p in PRODUCTS if p in it["description"] or p.split("-")[0] in it["description"].split()]
        if named:
            users = named                                   # product-specific part (e.g. "PM-200 main board")
        elif it["category"] in ("Displays", "Printed circuit boards", "Enclosures & plastics"):
            users = rng.choice(["PM-200", "IP-50"], size=1)
        else:
            users = rng.choice(list(PRODUCTS), size=rng.integers(1, 4), replace=False)
        if it["category"] == "Passive components":
            qty_range = (2, 40)
        elif "pad" in it["description"]:
            qty_range = (3, 10)
        else:
            qty_range = (1, 2)
        for p in users:
            rows.append({"product": p, "sku": it["sku"], "qty_per_unit": int(rng.integers(*qty_range, endpoint=True))})
    return pd.DataFrame(rows)


def build_production_plan() -> pd.DataFrame:
    base = {"PM-200": 420, "ECG-W1": 2600, "IP-50": 310}
    growth = {"PM-200": 0.006, "ECG-W1": 0.025, "IP-50": 0.004}     # wearables growing fastest
    rows = []
    for i, m in enumerate(MONTHS):
        season = 1 + 0.08 * np.sin(2 * np.pi * (m.month - 3) / 12) - (0.25 if m.month == 8 else 0)  # August shutdown
        for p in PRODUCTS:
            units = base[p] * (1 + growth[p]) ** i * season * rng.normal(1, 0.06)
            rows.append({"month": str(m), "product": p, "units_built": int(round(units))})
    return pd.DataFrame(rows)


def build_demand(items, bom, plan) -> pd.DataFrame:
    req = bom.merge(plan, on="product")
    req["qty"] = req["qty_per_unit"] * req["units_built"]
    dem = req.groupby(["sku", "month"])["qty"].sum().unstack().reindex(columns=[str(m) for m in MONTHS])
    # Scrap, service spares and engineering builds add noise; a few items are lumpy (spares-driven)
    shuffled = rng.permutation(dem.index)
    lumpy, variable = set(shuffled[:12]), set(shuffled[12:40])
    for sku in dem.index:
        noise = rng.normal(1.03, 0.05, len(MONTHS))
        if sku in variable:                                  # engineering builds, field-service spares
            noise = rng.normal(1.03, rng.uniform(0.22, 0.40), len(MONTHS)).clip(0.2)
        if sku in lumpy:
            noise = noise * rng.choice([0.2, 0.5, 1.0, 1.0, 2.2, 3.0], len(MONTHS))
        dem.loc[sku] = np.maximum(0, np.round(dem.loc[sku].values * noise))
    return dem.reset_index().melt(id_vars="sku", var_name="month", value_name="qty_consumed").astype({"qty_consumed": int})


def build_purchase_orders(items, sup, demand) -> pd.DataFrame:
    s = sup.set_index("supplier_id")
    avg = demand.groupby("sku")["qty_consumed"].mean()
    rows, po = [], 100000
    for _, it in items.iterrows():
        sp = s.loc[it["supplier_id"]]
        monthly = max(avg[it["sku"]], 1)
        orders_per_year = int(np.clip(round(monthly * 12 * it["unit_cost_eur"] / 6000), 2, 12))
        dates = pd.date_range("2024-10-05", "2026-09-20", periods=orders_per_year * 2) + \
            pd.to_timedelta(rng.integers(-4, 5, orders_per_year * 2), unit="D")
        for d in dates:
            qty = max(int(it["moq"]), int(round(monthly * 12 / orders_per_year * rng.normal(1, 0.1))))
            promised = d + pd.Timedelta(days=int(it["quoted_lead_time_days"]))
            late = 0 if rng.random() < sp["p_on_time"] else int(rng.gamma(2.0, 6 * sp["lt_multiplier"]))
            early = -int(rng.integers(0, 4)) if late == 0 else 0
            received = promised + pd.Timedelta(days=late + early)
            in_full = rng.random() < sp["p_in_full"]
            qty_recv = qty if in_full else int(qty * rng.uniform(0.6, 0.95))
            rejected = int(rng.binomial(qty_recv, sp["reject_rate"])) if qty_recv > 0 else 0
            # Semiconductor price increases from 2025-07 (allocation / shortage pricing)
            price = it["unit_cost_eur"]
            if it["category"] == "Microcontrollers & ICs" and d >= pd.Timestamp("2025-07-01"):
                price *= 1.18
            elif d >= pd.Timestamp("2026-01-01"):
                price *= 1.03
            open_po = received > SNAPSHOT_DATE
            rows.append({
                "po_number": f"PO{po}", "sku": it["sku"], "supplier_id": it["supplier_id"],
                "order_date": d.date(), "promised_date": promised.date(),
                "received_date": None if open_po else received.date(),
                "qty_ordered": qty, "qty_received": 0 if open_po else qty_recv,
                "qty_rejected": 0 if open_po else rejected, "unit_price_eur": round(price, 4),
                "status": "Open" if open_po else "Received",
            })
            po += 1
    return pd.DataFrame(rows).sort_values("order_date").reset_index(drop=True)


def build_inventory(items, demand, pos) -> pd.DataFrame:
    """Stock snapshot expressed relative to a sensible policy, with planner behaviour baked in.

    Cheap parts are over-bought, allocation-constrained chips and batteries run short, a share of
    items are over-forecast, and the rest sit in a healthy range.
    """
    last12 = demand[demand.month >= "2025-10"].pivot(index="sku", columns="month", values="qty_consumed")
    mean, sd = last12.mean(axis=1), last12.std(axis=1)
    rec = pos[pos.status == "Received"].copy()
    rec["lt"] = (pd.to_datetime(rec.received_date) - pd.to_datetime(rec.order_date)).dt.days
    lt_mean, lt_sd = rec.groupby("sku")["lt"].mean(), rec.groupby("sku")["lt"].std().fillna(0)
    on_order = pos[pos.status == "Open"].groupby("sku")["qty_ordered"].sum()
    rows = []
    for _, it in items.iterrows():
        k = it["sku"]
        m = max(mean[k], 1)
        ss = 1.75 * np.sqrt(lt_mean[k] / 30.4 * sd[k] ** 2 + m ** 2 * (lt_sd[k] / 30.4) ** 2)
        q = max(m, it["moq"])
        u = rng.random()
        if it["unit_cost_eur"] < 0.5 and u < 0.7:
            on_hand = ss + q * rng.uniform(2.5, 7)           # cheap parts bought in bulk "just in case"
        elif it["category"] in ("Microcontrollers & ICs", "Batteries") and u < 0.40:
            on_hand = ss * rng.uniform(0.25, 0.9)            # allocation-constrained parts running short
        elif u < 0.18:
            on_hand = ss + q * rng.uniform(2.0, 5.0)         # over-forecast / slow-moving
        else:
            on_hand = ss + q * rng.uniform(0.05, 0.95)       # healthy range
        expiry = None
        if it["shelf_life_months"]:
            expiry = (SNAPSHOT_DATE + pd.DateOffset(months=int(rng.integers(2, it["shelf_life_months"])))).date()
        rows.append({"sku": k, "on_hand_qty": int(round(on_hand)), "on_order_qty": int(on_order.get(k, 0)),
                     "oldest_lot_expiry": expiry, "snapshot_date": SNAPSHOT_DATE.date()})
    return pd.DataFrame(rows)


def build_budget(items, pos, plan) -> pd.DataFrame:
    """FY budget (Oct-25 to Sep-26) set before the year: planned purchase spend at STANDARD cost.

    Annual volume per category is planned at standard cost, then phased by month in line with the
    production plan. Price increases during the year therefore show up as variance.
    """
    p = pos.merge(items[["sku", "category", "unit_cost_eur"]], on="sku")
    p = p[pd.to_datetime(p.order_date) >= "2025-10-01"]
    annual = (p.qty_ordered * p.unit_cost_eur).groupby(p.category).sum() * rng.normal(1.0, 0.03, p.category.nunique())
    phasing = plan[plan.month >= "2025-10"].groupby("month")["units_built"].sum()
    phasing = phasing / phasing.sum()
    rows = [{"month": m, "category": c, "budget_eur": round(annual[c] * w, -2)}
            for m, w in phasing.items() for c in annual.index]
    return pd.DataFrame(rows)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    sup = build_suppliers()
    items = build_items(sup)
    bom = build_bom(items)
    plan = build_production_plan()
    demand = build_demand(items, bom, plan)
    pos = build_purchase_orders(items, sup, demand)
    inv = build_inventory(items, demand, pos)
    budget = build_budget(items, pos, plan)

    sup[["supplier_id", "supplier_name", "country"]].to_csv(OUT / "suppliers.csv", index=False)
    items.to_csv(OUT / "items.csv", index=False)
    bom.to_csv(OUT / "bom.csv", index=False)
    plan.to_csv(OUT / "production_plan.csv", index=False)
    demand.to_csv(OUT / "demand_monthly.csv", index=False)
    pos.to_csv(OUT / "purchase_orders.csv", index=False)
    inv.to_csv(OUT / "inventory_snapshot.csv", index=False)
    budget.to_csv(OUT / "budget_monthly.csv", index=False)
    print(f"items {len(items)} | suppliers {len(sup)} | POs {len(pos)} (open {(pos.status == 'Open').sum()}) | "
          f"demand rows {len(demand)}")


if __name__ == "__main__":
    main()
