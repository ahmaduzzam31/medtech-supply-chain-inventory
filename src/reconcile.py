"""Recalculate the Excel model in LibreOffice and reconcile it SKU by SKU with an independent Python calculation.

Writes reports/reconciliation.csv. Requires LibreOffice (`soffice`) on the PATH.
"""
import shutil
import subprocess
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "excel/Corrib_Medtech_Inventory_Model.xlsx"
D = ROOT / "data"


def recalculated_copy(tmp: Path) -> Path:
    src = tmp / "model.xlsx"
    shutil.copy(MODEL, src)
    out = tmp / "out"
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(out), str(src)],
                   check=True, capture_output=True, timeout=600)
    return out / "model.xlsx"


def python_policy() -> pd.DataFrame:
    items = pd.read_csv(D / "items.csv").set_index("sku")
    demand = pd.read_csv(D / "demand_monthly.csv")
    pos = pd.read_csv(D / "purchase_orders.csv", parse_dates=["order_date", "received_date"])
    stock = pd.read_csv(D / "inventory_snapshot.csv").set_index("sku")
    d12 = demand[demand.month >= "2025-10"].pivot(index="sku", columns="month", values="qty_consumed")
    p = items.copy()
    p["avg"], p["sd"] = d12.mean(axis=1), d12.std(axis=1)
    p["spend"] = d12.sum(axis=1) * p.unit_cost_eur
    p = p.sort_values("spend", ascending=False)
    cum = p.spend.cumsum() / p.spend.sum()
    p["abc"] = np.where(cum <= 0.80, "A", np.where(cum <= 0.95, "B", "C"))
    rec = pos[pos.status == "Received"].assign(lt=lambda d: (d.received_date - d.order_date).dt.days)
    p["lt_mean"], p["lt_sd"] = rec.groupby("sku")["lt"].mean(), rec.groupby("sku")["lt"].std().fillna(0)
    z = p.abc.map({"A": stats.norm.ppf(0.98), "B": stats.norm.ppf(0.95), "C": stats.norm.ppf(0.90)})
    p["ss"] = np.ceil(z * np.sqrt(p.lt_mean / 30.4 * p.sd ** 2 + p.avg ** 2 * (p.lt_sd / 30.4) ** 2))
    p["rop"] = np.ceil(p.avg * p.lt_mean / 30.4 + p.ss)
    p["value"] = stock.on_hand_qty * p.unit_cost_eur
    return p


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        xl = pd.read_excel(recalculated_copy(Path(tmp)), sheet_name="Inventory_Policy", header=3).set_index("SKU")
    py = python_policy().reindex(xl.index)
    checks = {
        "ABC class": (xl["ABC"] != py["abc"]).sum(),
        "Safety stock (units)": (xl["Safety stock (units)"] - py["ss"]).abs().gt(0.5).sum(),
        "Reorder point (units)": (xl["Reorder point (units)"] - py["rop"]).abs().gt(0.5).sum(),
        "Stock value (€)": (xl["Stock value (€)"] - py["value"]).abs().gt(0.01).sum(),
    }
    out = pd.DataFrame({"check": list(checks), "skus_compared": len(xl), "mismatches": list(checks.values())})
    out.to_csv(ROOT / "reports/reconciliation.csv", index=False)
    print(out.to_string(index=False))


if __name__ == "__main__":
    main()
