"""Build the Excel inventory model: excel/Corrib_Medtech_Inventory_Model.xlsx

Every analysis cell is a live Excel formula that recalculates when inputs change. Python only
loads the raw data and writes formulas; the calculations themselves live in Excel.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUT = ROOT / "excel/Corrib_Medtech_Inventory_Model.xlsx"

FONT = "Arial"
NAVY, BLUE_INPUT, GREEN_LINK = "1F3864", "0000FF", "008000"
F_BASE = Font(name=FONT, size=10)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SUB = Font(name=FONT, size=10, italic=True, color="52514E")
F_INPUT = Font(name=FONT, size=10, color=BLUE_INPUT)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_KPI = Font(name=FONT, size=18, bold=True, color=NAVY)
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_SECTION = PatternFill("solid", fgColor="DCE6F2")
FILL_KPI = PatternFill("solid", fgColor="F2F5FA")
THIN = Side(style="thin", color="C3C2B7")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RED_FILL, AMBER_FILL, GREEN_FILL, GREY_FILL = (PatternFill("solid", fgColor=c) for c in ("F8D7D7", "FCE9C5", "D9F2D9", "EEEEEE"))

EUR = '€#,##0;(€#,##0);"-"'
EUR2 = '€#,##0.00;(€#,##0.00);"-"'
EUR4 = '€#,##0.0000'
QTY = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = "dd-mmm-yyyy"

items = pd.read_csv(DATA / "items.csv")
suppliers = pd.read_csv(DATA / "suppliers.csv")
demand = pd.read_csv(DATA / "demand_monthly.csv")
pos = pd.read_csv(DATA / "purchase_orders.csv", parse_dates=["order_date", "promised_date", "received_date"])
stock = pd.read_csv(DATA / "inventory_snapshot.csv", parse_dates=["oldest_lot_expiry", "snapshot_date"])
budget = pd.read_csv(DATA / "budget_monthly.csv")

N_ITEMS, N_PO = len(items), len(pos)
I_END, P_END = N_ITEMS + 1, N_PO + 1            # last data row on data sheets (header in row 1)
POL_FIRST, POL_LAST = 5, 4 + N_ITEMS            # Inventory_Policy data rows
MONTHS = sorted(demand.month.unique())
FY_MONTHS = MONTHS[-12:]
CATS = list(dict.fromkeys(items.category))

wb = Workbook()


def header_row(ws, row, headers, widths=None):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font, cell.fill, cell.border = F_HEAD, FILL_HEAD, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = F_TITLE, F_SUB
    ws.sheet_view.showGridLines = False


def style_range(ws, rng, fmt=None, font=F_BASE):
    for row in ws[rng]:
        for cell in row:
            cell.font = font
            cell.border = BOX
            if fmt:
                cell.number_format = fmt


# ======================================================================================
# README
# ======================================================================================
ws = wb.active
ws.title = "README"
title(ws, "Corrib Medical Devices: Component Inventory & Supplier Model",
      "Simulated data for a fictional Galway medtech manufacturer (patient monitors, wearable ECG patches, infusion pumps)")
lines = [
    ("Purpose", "Set inventory policy (safety stock, reorder point, order quantity) for 116 electronic and packaging "
                "components, find excess and shortage risk, score suppliers, and explain budget variance."),
    ("How to use", "Change any yellow cell on the 'Assumptions' sheet (service levels, holding cost, thresholds). "
                   "Every analysis sheet recalculates. Use the SKU selector on the 'Dashboard' to inspect any part."),
    ("Sheets", "Dashboard: KPIs, charts, SKU lookup  |  Inventory_Policy: ABC/XYZ, safety stock, ROP, EOQ, status  |  "
               "Supplier_Scorecard: OTIF, rejects, lead times  |  Budget_vs_Actual: monthly variance with price/volume split  |  "
               "Assumptions  |  Data sheets: Items, Demand, PO_Lines, Stock_Snapshot, Suppliers"),
    ("Colour legend", "Blue text = hard-coded input  |  Black text = formula  |  Yellow fill = key assumption you can change"),
    ("Method", "ABC by 12-month spend (Pareto). XYZ by coefficient of variation of monthly demand. "
               "Safety stock = z x sqrt(LT x sd_demand^2 + demand^2 x sd_LT^2) using actual supplier lead times. "
               "ROP = demand during lead time + safety stock. Order qty = largest of EOQ (sqrt(2 x annual demand x order cost / holding cost per unit)), supplier MOQ and minimum order cover. Excess = stock above safety stock + one order quantity."),
    ("Data", "All data is simulated (see src/generate_data.py) and the company is fictional. Snapshot date 30-Sep-2026; "
             "24 months of consumption; 2,026 purchase order lines."),
]
for i, (k, v) in enumerate(lines, start=4):
    ws.cell(row=i, column=1, value=k).font = F_BOLD
    c = ws.cell(row=i, column=2, value=v)
    c.font, c.alignment = F_BASE, Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[i].height = 45
ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 16, 120
ws["B7"].font = Font(name=FONT, size=10, color=BLUE_INPUT)

# ======================================================================================
# Assumptions
# ======================================================================================
wa = wb.create_sheet("Assumptions")
title(wa, "Assumptions", "Yellow cells are inputs. Every analysis formula references these cells.")
assump = [
    ("Snapshot date", pd.Timestamp("2026-09-30").to_pydatetime(), DATE, "Date of the stock count."),
    ("Ordering cost per PO (€)", 85, EUR, "Buyer time, receiving, inspection and invoice processing per order."),
    ("Annual holding cost rate", 0.22, PCT, "Cost of capital + storage + insurance + obsolescence, % of unit cost per year."),
    ("ABC: A-class up to cumulative spend", 0.80, PCT, "Items making up the first 80% of spend are A."),
    ("ABC: B-class up to cumulative spend", 0.95, PCT, "Next 15% are B, remainder C."),
    ("XYZ: X-class max coefficient of variation", 0.25, "0.00", "Stable demand."),
    ("XYZ: Y-class max coefficient of variation", 0.50, "0.00", "Variable demand; above this is Z (erratic)."),
    ("Days per month", 30.4, "0.0", "Converts lead time in days to months of demand."),
    ("Expiry warning window (months)", 6, "0", "Flag lots expiring within this window."),
    ("Budget year start", pd.Timestamp("2025-10-01").to_pydatetime(), DATE, "FY runs Oct to Sep."),
    ("Supplier OTIF: Preferred at or above", 0.95, PCT, "Rating thresholds for the scorecard."),
    ("Supplier OTIF: Approved at or above", 0.85, PCT, "Below this = improvement plan."),
    ("Minimum order cover (months)", 1, "0.0", "Suppliers ship monthly at most, so order quantity is never below 1 month of demand."),
]
header_row(wa, 3, ["Assumption", "Value", "Notes"], [42, 16, 70])
for i, (label, val, fmt, note) in enumerate(assump, start=4):
    wa.cell(row=i, column=1, value=label).font = F_BASE
    c = wa.cell(row=i, column=2, value=val)
    c.font, c.fill, c.number_format, c.border = F_INPUT, FILL_INPUT, fmt, BOX
    wa.cell(row=i, column=3, value=note).font = F_SUB
A = {k: f"Assumptions!$B${i}" for i, (k, *_r) in enumerate(assump, start=4)}
SNAP, ORDER_COST, HOLD, A_THR, B_THR, X_THR, Y_THR, DPM, EXP_WIN, FY_START, OTIF_P, OTIF_A, MIN_COVER = A.values()

# Service level table (E-H)
wa["E3"], wa["F3"], wa["G3"] = "ABC class", "Target service level", "z-score"
for c in ("E3", "F3", "G3"):
    wa[c].font, wa[c].fill, wa[c].border = F_HEAD, FILL_HEAD, BOX
for i, (cls, sl) in enumerate([("A", 0.98), ("B", 0.95), ("C", 0.90)], start=4):
    wa.cell(row=i, column=5, value=cls).font = F_BOLD
    c = wa.cell(row=i, column=6, value=sl)
    c.font, c.fill, c.number_format, c.border = F_INPUT, FILL_INPUT, PCT, BOX
    z = wa.cell(row=i, column=7, value=f"=NORMSINV(F{i})")
    z.number_format, z.border, z.font = "0.00", BOX, F_BASE
wa["E8"] = "Higher service for A items: they are critical to production and expensive to run out of."
wa["E8"].font = F_SUB
for col, w in zip("EFG", (12, 20, 10)):
    wa.column_dimensions[col].width = w

# ======================================================================================
# Data sheets (values only, blue = inputs)
# ======================================================================================
def write_df(name, df, formats, widths, note):
    w = wb.create_sheet(name)
    header_row(w, 1, list(df.columns), widths)
    for r, row in enumerate(df.itertuples(index=False), start=2):
        for c, v in enumerate(row, 1):
            if pd.isna(v) if not isinstance(v, (list, tuple)) else False:
                v = None
            elif isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            cell = w.cell(row=r, column=c, value=v)
            cell.font = F_INPUT
            if formats.get(c):
                cell.number_format = formats[c]
    w.freeze_panes = "B2"
    w.auto_filter.ref = f"A1:{get_column_letter(df.shape[1])}{len(df) + 1}"
    w.cell(row=1, column=1).comment = Comment(note, "Model")
    return w


write_df("Items", items, {5: EUR4, 6: "0", 7: QTY}, [11, 30, 22, 10, 12, 12, 9, 10, 10],
         "Item master: one row per component SKU. unit_cost_eur = standard cost.")
write_df("Suppliers", suppliers, {}, [11, 30, 16], "Approved supplier list.")
dem_wide = demand.pivot(index="sku", columns="month", values="qty_consumed").reset_index()
write_df("Demand", dem_wide, {c: QTY for c in range(2, 26)}, [11] + [9] * 24,
         "Monthly consumption (units) per SKU, from production backflush. Columns N:Y = last 12 months.")
write_df("Stock_Snapshot", stock, {2: QTY, 3: QTY, 4: DATE, 5: DATE}, [11, 12, 12, 16, 14],
         "Physical stock count at the snapshot date. on_order_qty = open purchase orders.")

wp = write_df("PO_Lines", pos, {4: DATE, 5: DATE, 6: DATE, 7: QTY, 8: QTY, 9: QTY, 10: EUR4},
              [11, 11, 9, 12, 12, 12, 11, 11, 10, 11, 10], "Purchase order lines, Oct-2024 to Sep-2026.")
po_calc = [
    ("category", "=INDEX(Items!$C$2:$C${ie},MATCH(B{r},Items!$A$2:$A${ie},0))", None, 22),
    ("std_cost_eur", "=INDEX(Items!$E$2:$E${ie},MATCH(B{r},Items!$A$2:$A${ie},0))", EUR4, 11),
    ("order_month", '=TEXT(D{r},"yyyy-mm")', None, 10),
    ("po_value_eur", "=G{r}*J{r}", EUR, 12),
    ("price_variance_eur", "=(J{r}-M{r})*G{r}", EUR, 12),
    ("received_flag", '=IF(K{r}="Received",1,0)', "0", 9),
    ("lead_time_days", "=IF(Q{r}=1,F{r}-D{r},0)", "0", 9),
    ("on_time", "=IF(Q{r}=1,IF(F{r}<=E{r},1,0),0)", "0", 8),
    ("in_full", "=IF(Q{r}=1,IF(H{r}>=G{r},1,0),0)", "0", 8),
    ("otif", "=S{r}*T{r}", "0", 7),
    ("days_late", "=IF(Q{r}=1,MAX(0,F{r}-E{r}),0)", "0", 8),
]
for j, (h, f, fmt, wdt) in enumerate(po_calc, start=12):
    hc = wp.cell(row=1, column=j, value=h)
    hc.font, hc.fill, hc.border = F_HEAD, PatternFill("solid", fgColor="2A78D6"), BOX
    hc.alignment = Alignment(horizontal="center", wrap_text=True)
    wp.column_dimensions[get_column_letter(j)].width = wdt
    for r in range(2, P_END + 1):
        c = wp.cell(row=r, column=j, value=f.format(r=r, ie=I_END))
        c.font = F_BASE
        if fmt:
            c.number_format = fmt
wp.auto_filter.ref = f"A1:V{P_END}"
wp["L1"].comment = Comment("Columns L:V (lighter header) are formulas.", "Model")

PO = {k: f"PO_Lines!${c}$2:${c}${P_END}" for k, c in
      dict(sku="B", sup="C", order_date="D", qty_ord="G", qty_recv="H", qty_rej="I", cat="L", month="N",
           value="O", pvar="P", recv="Q", lt="R", ontime="S", infull="T", otif="U", late="V").items()}

# ======================================================================================
# Inventory_Policy
# ======================================================================================
wi = wb.create_sheet("Inventory_Policy", 1)
title(wi, "Inventory Policy by SKU",
      "ABC/XYZ classification, safety stock, reorder point, EOQ and stock status. All cells are formulas driven by the Assumptions sheet.")
cols = [
    ("SKU", None, 11), ("Description", None, 28), ("Category", None, 20), ("Supplier", None, 9),
    ("Unit cost (€)", EUR4, 10), ("Annual usage (units)", QTY, 11), ("Avg monthly demand", QTY, 10),
    ("Std dev monthly", QTY, 10), ("CV", "0.00", 7), ("Annual spend (€)", EUR, 12), ("Cumulative spend %", PCT, 10),
    ("ABC", None, 6), ("XYZ", None, 6), ("ABC-XYZ", None, 8), ("Service level", PCT, 9), ("z", "0.00", 6),
    ("Avg actual lead time (days)", "0.0", 10), ("Lead time std dev (days)", "0.0", 10), ("Safety stock (units)", QTY, 11),
    ("Reorder point (units)", QTY, 11), ("Order qty (EOQ, MOQ, min cover)", QTY, 12), ("On hand (units)", QTY, 11),
    ("On order (units)", QTY, 11), ("Inventory position", QTY, 11), ("Days of cover", "0", 8),
    ("Max stock (SS + order qty)", QTY, 11), ("Stock status", None, 17), ("Stock value (€)", EUR, 12),
    ("Excess value (€)", EUR, 12), ("Suggested order qty", QTY, 11), ("Shortfall vs ROP (€)", EUR, 12),
    ("Oldest lot expiry", DATE, 12), ("Expiry flag", None, 13),
]
header_row(wi, 4, [c[0] for c in cols], [c[2] for c in cols])
wi.row_dimensions[4].height = 42
ie, mrow = I_END, "MATCH($A{r},Items!$A$2:$A${ie},0)"
f = {
    "B": "=INDEX(Items!$B$2:$B${ie}," + mrow + ")",
    "C": "=INDEX(Items!$C$2:$C${ie}," + mrow + ")",
    "D": "=INDEX(Items!$D$2:$D${ie}," + mrow + ")",
    "E": "=INDEX(Items!$E$2:$E${ie}," + mrow + ")",
    "F": "=SUM(INDEX(Demand!$N$2:$Y${ie},MATCH($A{r},Demand!$A$2:$A${ie},0),0))",
    "G": "=F{r}/12",
    "H": "=STDEV(INDEX(Demand!$N$2:$Y${ie},MATCH($A{r},Demand!$A$2:$A${ie},0),0))",
    "I": "=IF(G{r}>0,H{r}/G{r},0)",
    "J": "=F{r}*E{r}",
    "K": "=SUMIF($J${p0}:$J${p1},\">=\"&J{r})/SUM($J${p0}:$J${p1})",
    "L": f"=IF(K{{r}}<={A_THR},\"A\",IF(K{{r}}<={B_THR},\"B\",\"C\"))",
    "M": f"=IF(I{{r}}<={X_THR},\"X\",IF(I{{r}}<={Y_THR},\"Y\",\"Z\"))",
    "N": "=L{r}&M{r}",
    "O": "=INDEX(Assumptions!$F$4:$F$6,MATCH(L{r},Assumptions!$E$4:$E$6,0))",
    "P": "=NORMSINV(O{r})",
    "Q": f"=IFERROR(AVERAGEIFS({PO['lt']},{PO['sku']},$A{{r}},{PO['recv']},1),INDEX(Items!$F$2:$F${{ie}},{mrow}))",
    "R": f"=IFERROR(SQRT(SUMPRODUCT(({PO['sku']}=$A{{r}})*{PO['recv']}*({PO['lt']}-Q{{r}})^2)"
         f"/(COUNTIFS({PO['sku']},$A{{r}},{PO['recv']},1)-1)),0)",
    "S": f"=ROUNDUP(P{{r}}*SQRT((Q{{r}}/{DPM})*H{{r}}^2+G{{r}}^2*(R{{r}}/{DPM})^2),0)",
    "T": f"=ROUNDUP(G{{r}}*Q{{r}}/{DPM}+S{{r}},0)",
    "U": f"=ROUNDUP(MAX(INDEX(Items!$G$2:$G${{ie}},{mrow}),SQRT(2*F{{r}}*{ORDER_COST}/({HOLD}*E{{r}})),G{{r}}*{MIN_COVER}),0)",
    "V": "=INDEX(Stock_Snapshot!$B$2:$B${ie},MATCH($A{r},Stock_Snapshot!$A$2:$A${ie},0))",
    "W": "=INDEX(Stock_Snapshot!$C$2:$C${ie},MATCH($A{r},Stock_Snapshot!$A$2:$A${ie},0))",
    "X": "=V{r}+W{r}",
    "Y": f"=IF(G{{r}}>0,V{{r}}/(G{{r}}/{DPM}),999)",
    "Z": "=S{r}+U{r}",
    "AA": '=IF(V{r}<S{r},"Below safety stock",IF(X{r}<=T{r},"Reorder now",IF(V{r}>Z{r},"Excess","Healthy")))',
    "AB": "=V{r}*E{r}",
    "AC": "=MAX(0,V{r}-Z{r})*E{r}",
    "AD": "=IF(X{r}<=T{r},ROUNDUP(T{r}+U{r}-X{r},0),0)",
    "AE": "=MAX(0,T{r}-X{r})*E{r}",
    "AF": '=IF(INDEX(Stock_Snapshot!$D$2:$D${ie},MATCH($A{r},Stock_Snapshot!$A$2:$A${ie},0))="","",'
          'INDEX(Stock_Snapshot!$D$2:$D${ie},MATCH($A{r},Stock_Snapshot!$A$2:$A${ie},0)))',
    "AG": f'=IF(AF{{r}}="","",IF(AF{{r}}<=EDATE({SNAP},{EXP_WIN}),"Expires < "&{EXP_WIN}&" m","OK"))',
}
for i, sku in enumerate(items.sku):
    r = POL_FIRST + i
    c = wi.cell(row=r, column=1, value=sku)
    c.font, c.border = F_BASE, BOX
    for col_idx in range(2, len(cols) + 1):
        letter = get_column_letter(col_idx)
        cell = wi.cell(row=r, column=col_idx, value=f[letter].format(r=r, ie=ie, p0=POL_FIRST, p1=POL_LAST))
        cell.font, cell.border = F_BASE, BOX
        if cols[col_idx - 1][1]:
            cell.number_format = cols[col_idx - 1][1]
wi.freeze_panes = "B5"
wi.auto_filter.ref = f"A4:AG{POL_LAST}"
rng = f"AA{POL_FIRST}:AA{POL_LAST}"
for text, fill in [("Below safety stock", RED_FILL), ("Reorder now", AMBER_FILL), ("Excess", GREY_FILL), ("Healthy", GREEN_FILL)]:
    wi.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill))
wi.conditional_formatting.add(f"AG{POL_FIRST}:AG{POL_LAST}", FormulaRule(formula=[f'LEFT(AG{POL_FIRST},7)="Expires"'], fill=RED_FILL))
wi.conditional_formatting.add(f"L{POL_FIRST}:L{POL_LAST}", CellIsRule(operator="equal", formula=['"A"'], font=Font(name=FONT, bold=True, color="C00000")))
POL = lambda col: f"Inventory_Policy!${col}${POL_FIRST}:${col}${POL_LAST}"

# ======================================================================================
# Supplier_Scorecard
# ======================================================================================
wsu = wb.create_sheet("Supplier_Scorecard", 2)
title(wsu, "Supplier Scorecard", "On-time, in-full (OTIF), quality and lead-time performance per supplier, all POs received Oct-2024 to Sep-2026.")
sc_cols = [("Supplier ID", None, 10), ("Supplier", None, 28), ("Country", None, 15), ("SKUs supplied", "0", 9),
           ("Single-source SKUs", "0", 10), ("POs received", "0", 10), ("FY spend (€)", EUR, 13), ("On-time %", PCT, 10),
           ("In-full %", PCT, 10), ("OTIF %", PCT, 10), ("Avg days late (late POs)", "0.0", 11), ("Reject rate", '0.00%', 10),
           ("Avg actual lead time (days)", "0", 11), ("Rating", None, 17)]
header_row(wsu, 4, [c[0] for c in sc_cols], [c[2] for c in sc_cols])
wsu.row_dimensions[4].height = 42
S_FIRST, S_LAST = 5, 4 + len(suppliers)
for i, s in enumerate(suppliers.itertuples(index=False)):
    r = S_FIRST + i
    vals = [s.supplier_id, s.supplier_name, s.country,
            f"=COUNTIF(Items!$D$2:$D${ie},A{r})",
            f"=COUNTIFS(Items!$D$2:$D${ie},A{r},Items!$I$2:$I${ie},TRUE)",
            f"=COUNTIFS({PO['sup']},A{r},{PO['recv']},1)",
            f"=SUMIFS({PO['value']},{PO['sup']},A{r},{PO['order_date']},\">=\"&{FY_START})",
            f"=IFERROR(SUMIFS({PO['ontime']},{PO['sup']},A{r})/F{r},0)",
            f"=IFERROR(SUMIFS({PO['infull']},{PO['sup']},A{r})/F{r},0)",
            f"=IFERROR(SUMIFS({PO['otif']},{PO['sup']},A{r})/F{r},0)",
            f"=IFERROR(AVERAGEIFS({PO['late']},{PO['sup']},A{r},{PO['late']},\">0\"),0)",
            f"=IFERROR(SUMIFS({PO['qty_rej']},{PO['sup']},A{r})/SUMIFS({PO['qty_recv']},{PO['sup']},A{r}),0)",
            f"=IFERROR(AVERAGEIFS({PO['lt']},{PO['sup']},A{r},{PO['recv']},1),0)",
            f'=IF(J{r}>={OTIF_P},"Preferred",IF(J{r}>={OTIF_A},"Approved","Improvement plan"))']
    for j, v in enumerate(vals, 1):
        c = wsu.cell(row=r, column=j, value=v)
        c.font, c.border = F_BASE, BOX
        if sc_cols[j - 1][1]:
            c.number_format = sc_cols[j - 1][1]
tr = S_LAST + 1
wsu.cell(row=tr, column=1, value="Total / weighted").font = F_BOLD
for col, fml in {"D": f"=SUM(D{S_FIRST}:D{S_LAST})", "E": f"=SUM(E{S_FIRST}:E{S_LAST})",
                 "F": f"=SUM(F{S_FIRST}:F{S_LAST})", "G": f"=SUM(G{S_FIRST}:G{S_LAST})",
                 "H": f"=SUMPRODUCT(H{S_FIRST}:H{S_LAST},$F${S_FIRST}:$F${S_LAST})/$F${tr}",
                 "I": f"=SUMPRODUCT(I{S_FIRST}:I{S_LAST},$F${S_FIRST}:$F${S_LAST})/$F${tr}",
                 "J": f"=SUMPRODUCT(J{S_FIRST}:J{S_LAST},$F${S_FIRST}:$F${S_LAST})/$F${tr}"}.items():
    c = wsu[f"{col}{tr}"]
    c.value, c.font, c.border = fml, F_BOLD, BOX
    c.number_format = {"G": EUR}.get(col, PCT if col in "HIJ" else "0")
wsu.conditional_formatting.add(f"J{S_FIRST}:J{S_LAST}", ColorScaleRule(start_type="num", start_value=0.65, start_color="F8D7D7",
                                                                      mid_type="num", mid_value=0.85, mid_color="FCE9C5",
                                                                      end_type="num", end_value=1, end_color="D9F2D9"))
for text, fill in [("Improvement plan", RED_FILL), ("Approved", AMBER_FILL), ("Preferred", GREEN_FILL)]:
    wsu.conditional_formatting.add(f"N{S_FIRST}:N{S_LAST}", CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill))
wsu.freeze_panes = "C5"
wsu[f"A{tr + 2}"] = "OTIF = share of received POs delivered on or before the promised date AND in full. Rating thresholds are on the Assumptions sheet."
wsu[f"A{tr + 2}"].font = F_SUB

# ======================================================================================
# Budget_vs_Actual
# ======================================================================================
wbv = wb.create_sheet("Budget_vs_Actual", 3)
title(wbv, "Purchasing Budget vs Actual: FY Oct-2025 to Sep-2026",
      "Budget = planned purchases at standard cost. Actual = PO value at actual price (formulas on PO_Lines).")
wbv.column_dimensions["A"].width = 24
for j in range(2, 16):
    wbv.column_dimensions[get_column_letter(j)].width = 11


def block(top, label, cell_fn, fmt=EUR):
    wbv.cell(row=top, column=1, value=label).font = F_BOLD
    wbv.cell(row=top, column=1).fill = FILL_SECTION
    header_row(wbv, top + 1, ["Category"] + FY_MONTHS + ["FY total"])
    for i, cat in enumerate(CATS):
        r = top + 2 + i
        wbv.cell(row=r, column=1, value=cat).font = F_BASE
        for j, m in enumerate(FY_MONTHS, start=2):
            c = wbv.cell(row=r, column=j, value=cell_fn(r, j, cat, m))
            c.number_format, c.border = fmt, BOX
            c.font = F_INPUT if label.startswith("Budget") else F_BASE
        t = wbv.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})")
        t.number_format, t.border, t.font = fmt, BOX, F_BOLD
    tot = top + 2 + len(CATS)
    wbv.cell(row=tot, column=1, value="Total").font = F_BOLD
    for j in range(2, 15):
        L = get_column_letter(j)
        c = wbv.cell(row=tot, column=j, value=f"=SUM({L}{top + 2}:{L}{tot - 1})")
        c.number_format, c.font, c.border = fmt, F_BOLD, BOX
    return top + 2, tot


bud = budget.set_index(["category", "month"])["budget_eur"]
B0, BT = block(4, "Budget (€)", lambda r, j, cat, m: float(bud[(cat, m)]))
A0, AT = block(BT + 3, "Actual (€)", lambda r, j, cat, m:
               f"=SUMIFS({PO['value']},{PO['cat']},$A{r},{PO['month']},{get_column_letter(j)}${BT + 4})")
V0, VT = block(AT + 3, "Variance: actual minus budget (€)", lambda r, j, cat, m:
               f"={get_column_letter(j)}{A0 + (r - (AT + 5))}-{get_column_letter(j)}{B0 + (r - (AT + 5))}")
for rr in range(V0, VT + 1):
    for j in range(2, 15):
        c = wbv.cell(row=rr, column=j)
        c.font = F_BOLD if rr == VT or j == 14 else F_BASE
wbv.conditional_formatting.add(f"B{V0}:N{VT}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, color="C00000")))
wbv.conditional_formatting.add(f"B{V0}:N{VT}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color="008000")))

# Variance bridge: price vs volume/timing
S0 = VT + 3
wbv.cell(row=S0, column=1, value="FY variance analysis: price vs volume/timing").font = F_BOLD
wbv.cell(row=S0, column=1).fill = FILL_SECTION
header_row(wbv, S0 + 1, ["Category", "FY budget (€)", "FY actual (€)", "Variance (€)", "Variance %",
                         "Price variance (€)", "Volume / timing (€)", "Main driver"])
for i, cat in enumerate(CATS + ["Total"]):
    r = S0 + 2 + i
    src = i if cat != "Total" else len(CATS)
    vals = [cat, f"=N{B0 + src}", f"=N{A0 + src}", f"=C{r}-B{r}", f"=IFERROR(D{r}/B{r},0)",
            (f"=SUMIFS({PO['pvar']},{PO['cat']},A{r},{PO['order_date']},\">=\"&{FY_START})" if cat != "Total"
             else f"=SUM(F{S0 + 2}:F{r - 1})"),
            f"=D{r}-F{r}",
            f'=IF(ABS(D{r})<0.02*B{r},"On budget",IF(ABS(F{r})>ABS(G{r}),"Price","Volume / timing"))']
    for j, v in enumerate(vals, 1):
        c = wbv.cell(row=r, column=j, value=v)
        c.font, c.border = (F_BOLD if cat == "Total" else F_BASE), BOX
        c.number_format = {2: EUR, 3: EUR, 4: EUR, 5: PCT, 6: EUR, 7: EUR}.get(j, "General")
BRIDGE_FIRST, BRIDGE_TOTAL = S0 + 2, S0 + 2 + len(CATS)
wbv.conditional_formatting.add(f"D{BRIDGE_FIRST}:G{BRIDGE_TOTAL}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, color="C00000")))
wbv.freeze_panes = "B4"

# ======================================================================================
# Dashboard
# ======================================================================================
wd = wb.create_sheet("Dashboard", 0)
title(wd, "Inventory & Supplier Dashboard: Corrib Medical Devices",
      "Snapshot 30-Sep-2026. All figures are live formulas; change the Assumptions sheet to run scenarios.")
for col, w in zip("ABCDEFGHIJKLMN", (24, 24, 14, 14, 3, 24, 24, 14, 14, 3, 22, 14, 14, 14)):
    wd.column_dimensions[col].width = w

kpis = [
    ("Inventory value", f"=SUM({POL('AB')})", EUR),
    ("Excess stock value", f"=SUM({POL('AC')})", EUR),
    ("SKUs below safety stock", f'=COUNTIF({POL("AA")},"Below safety stock")', "0"),
    ("SKUs to reorder now", f'=COUNTIF({POL("AA")},"Reorder now")', "0"),
    ("Supplier OTIF (weighted)", f"=Supplier_Scorecard!J{tr}", PCT),
    ("FY purchase variance", f"=Budget_vs_Actual!D{BRIDGE_TOTAL}", EUR),
    ("of which price variance", f"=Budget_vs_Actual!F{BRIDGE_TOTAL}", EUR),
]
for k, (label, fml, fmt) in enumerate(kpis):
    col = 1 + (k % 4) * 3 if k < 4 else 1 + (k - 4) * 3
    row = 4 if k < 4 else 7
    lab = wd.cell(row=row, column=col, value=label)
    lab.font, lab.fill = Font(name=FONT, size=9, color="52514E"), FILL_KPI
    val = wd.cell(row=row + 1, column=col, value=fml)
    val.font, val.fill, val.number_format = F_KPI, FILL_KPI, fmt
    wd.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    wd.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
wd.row_dimensions[5].height = wd.row_dimensions[8].height = 28

# ABC summary table
r0 = 11
wd.cell(row=r0 - 1, column=1, value="ABC classification").font = F_BOLD
header_row(wd, r0, ["Class", "SKUs", "% of SKUs", "% of spend"])
for i, cls in enumerate("ABC"):
    r = r0 + 1 + i
    wd.cell(row=r, column=1, value=cls).font = F_BOLD
    wd.cell(row=r, column=2, value=f'=COUNTIF({POL("L")},A{r})').number_format = "0"
    wd.cell(row=r, column=3, value=f"=B{r}/SUM($B${r0 + 1}:$B${r0 + 3})").number_format = PCT
    wd.cell(row=r, column=4, value=f'=SUMIF({POL("L")},A{r},{POL("J")})/SUM({POL("J")})').number_format = PCT
style_range(wd, f"B{r0 + 1}:D{r0 + 3}")
for i in range(3):
    wd.cell(row=r0 + 1 + i, column=3).number_format = PCT
    wd.cell(row=r0 + 1 + i, column=4).number_format = PCT

# Stock status table
wd.cell(row=r0 - 1, column=6, value="Stock status").font = F_BOLD
for j, h in enumerate(["Status", "SKUs", "Stock value (€)", "Excess (€)"], start=6):
    c = wd.cell(row=r0, column=j, value=h)
    c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BOX
for i, st in enumerate(["Below safety stock", "Reorder now", "Healthy", "Excess"]):
    r = r0 + 1 + i
    wd.cell(row=r, column=6, value=st).font = F_BASE
    wd.cell(row=r, column=7, value=f'=COUNTIF({POL("AA")},F{r})').number_format = "0"
    wd.cell(row=r, column=8, value=f'=SUMIF({POL("AA")},F{r},{POL("AB")})').number_format = EUR
    wd.cell(row=r, column=9, value=f'=SUMIF({POL("AA")},F{r},{POL("AC")})').number_format = EUR
    for j in range(6, 10):
        wd.cell(row=r, column=j).border = BOX
for text, fill in [("Below safety stock", RED_FILL), ("Reorder now", AMBER_FILL), ("Excess", GREY_FILL), ("Healthy", GREEN_FILL)]:
    wd.conditional_formatting.add(f"F{r0 + 1}:F{r0 + 4}", CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill))

# ABC-XYZ matrix
wd.cell(row=r0 - 1, column=11, value="ABC-XYZ matrix (SKU count)").font = F_BOLD
for j, h in enumerate(["ABC \\ XYZ", "X stable", "Y variable", "Z erratic"], start=11):
    c = wd.cell(row=r0, column=j, value=h)
    c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BOX
for i, cls in enumerate("ABC"):
    r = r0 + 1 + i
    wd.cell(row=r, column=11, value=cls).font = F_BOLD
    for j, x in enumerate("XYZ", start=12):
        c = wd.cell(row=r, column=j, value=f'=COUNTIF({POL("N")},"{cls}{x}")')
        c.border, c.number_format = BOX, "0"
wd.conditional_formatting.add(f"L{r0 + 1}:N{r0 + 3}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="6DA7EC"))

# Top 10 excess and top 10 shortfall lists (LARGE + INDEX/MATCH, no dynamic arrays)
t0 = r0 + 7
for col0, label, src in [(1, "Top 10 excess stock (€)", "AC"), (6, "Top 10 shortfall vs reorder point (€)", "AE")]:
    wd.cell(row=t0 - 1, column=col0, value=label).font = F_BOLD
    for j, h in enumerate(["SKU", "Description", "Supplier", "Value (€)"]):
        c = wd.cell(row=t0, column=col0 + j, value=h)
        c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BOX
    for k in range(1, 11):
        r = t0 + k
        Lv = get_column_letter(col0 + 3)
        val = wd.cell(row=r, column=col0 + 3, value=f"=LARGE({POL(src)},{k})")
        val.number_format = EUR
        match = f"MATCH({Lv}{r},{POL(src)},0)"
        wd.cell(row=r, column=col0, value=f'=IF({Lv}{r}=0,"",INDEX({POL("A")},{match}))')
        wd.cell(row=r, column=col0 + 1, value=f'=IF({Lv}{r}=0,"",INDEX({POL("B")},{match}))')
        wd.cell(row=r, column=col0 + 2, value=f'=IF({Lv}{r}=0,"",INDEX({POL("D")},{match}))')
        for j in range(4):
            wd.cell(row=r, column=col0 + j).border = BOX
            wd.cell(row=r, column=col0 + j).font = F_BASE

# SKU lookup
k0 = t0
wd.cell(row=k0 - 1, column=11, value="SKU lookup: pick a SKU").font = F_BOLD
sel = wd.cell(row=k0, column=12, value=items.sku.iloc[0])
sel.font, sel.fill, sel.border = F_INPUT, FILL_INPUT, BOX
wd.cell(row=k0, column=11, value="Selected SKU").font = F_BOLD
dv = DataValidation(type="list", formula1=f"={POL('A')}", allow_blank=False)
wd.add_data_validation(dv)
dv.add(sel)
lookup = [("Description", "B", None), ("Category", "C", None), ("Supplier", "D", None), ("ABC-XYZ", "N", None),
          ("Avg monthly demand", "G", QTY), ("Actual lead time (days)", "Q", "0.0"), ("Safety stock", "S", QTY),
          ("Reorder point", "T", QTY), ("Order quantity", "U", QTY), ("On hand", "V", QTY), ("On order", "W", QTY),
          ("Days of cover", "Y", "0"), ("Status", "AA", None), ("Suggested order qty", "AD", QTY)]
for i, (lab, col, fmt) in enumerate(lookup, start=1):
    r = k0 + i
    wd.cell(row=r, column=11, value=lab).font = F_BASE
    c = wd.cell(row=r, column=12, value=f"=INDEX({POL(col)},MATCH($L${k0},{POL('A')},0))")
    c.font, c.border = F_BOLD, BOX
    if fmt:
        c.number_format = fmt
    wd.merge_cells(start_row=r, start_column=12, end_row=r, end_column=13)

# Supplier OTIF chart data (sorted copy via formulas pointing at scorecard) and charts
CHART_ROW = t0 + 17
c0 = CHART_ROW + 21
wd.cell(row=c0 - 1, column=1, value="Chart data (feeds the charts above)").font = F_SUB
header_row(wd, c0, ["Supplier", "OTIF %"])
for i in range(len(suppliers)):
    r = c0 + 1 + i
    wd.cell(row=r, column=1, value=f"=Supplier_Scorecard!B{S_FIRST + i}")
    wd.cell(row=r, column=2, value=f"=Supplier_Scorecard!J{S_FIRST + i}").number_format = PCT
for j, h in enumerate(["Month", "Budget (€)", "Actual (€)"], start=6):
    c = wd.cell(row=c0, column=j, value=h)
    c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BOX
for i, m in enumerate(FY_MONTHS):
    r = c0 + 1 + i
    col = get_column_letter(2 + i)
    wd.cell(row=r, column=6, value=m)
    wd.cell(row=r, column=7, value=f"=Budget_vs_Actual!{col}{BT}").number_format = EUR
    wd.cell(row=r, column=8, value=f"=Budget_vs_Actual!{col}{AT}").number_format = EUR

ch = BarChart()
ch.type, ch.title, ch.style = "bar", "Supplier OTIF %", 2
ch.add_data(Reference(wd, min_col=2, min_row=c0, max_row=c0 + len(suppliers)), titles_from_data=True)
ch.set_categories(Reference(wd, min_col=1, min_row=c0 + 1, max_row=c0 + len(suppliers)))
ch.y_axis.number_format, ch.y_axis.scaling.min, ch.y_axis.scaling.max = "0%", 0.5, 1.0
ch.legend, ch.height, ch.width = None, 9, 16
ch.series[0].graphicalProperties.solidFill = "2A78D6"
wd.add_chart(ch, f"A{CHART_ROW}")

lc = LineChart()
lc.title, lc.style, lc.height, lc.width = "Monthly purchasing: budget vs actual (€)", 2, 9, 16
lc.add_data(Reference(wd, min_col=7, max_col=8, min_row=c0, max_row=c0 + 12), titles_from_data=True)
lc.set_categories(Reference(wd, min_col=6, min_row=c0 + 1, max_row=c0 + 12))
lc.series[0].graphicalProperties.line.solidFill = "898781"
lc.series[1].graphicalProperties.line.solidFill = "EB6834"
lc.y_axis.number_format = '€#,##0,"k"'
for srs in lc.series:
    srs.smooth = False
    srs.graphicalProperties.line.width = 28000
lc.legend.position = "b"
wd.add_chart(lc, f"F{CHART_ROW}")

bc = BarChart()
bc.title, bc.style, bc.height, bc.width = "Share of SKUs vs share of spend by ABC class", 2, 9, 14
bc.add_data(Reference(wd, min_col=3, max_col=4, min_row=r0, max_row=r0 + 3), titles_from_data=True)
bc.set_categories(Reference(wd, min_col=1, min_row=r0 + 1, max_row=r0 + 3))
bc.series[0].graphicalProperties.solidFill = "86B6EF"
bc.series[1].graphicalProperties.solidFill = "1F3864"
bc.y_axis.number_format = "0%"
bc.dataLabels = DataLabelList()
bc.dataLabels.showVal = True
bc.dataLabels.showSerName = bc.dataLabels.showCatName = bc.dataLabels.showLegendKey = False
bc.legend.position = "b"
wd.add_chart(bc, f"K{CHART_ROW}")

for sheet in wb.worksheets:
    sheet.sheet_properties.tabColor = {"Dashboard": "1F3864", "Inventory_Policy": "2A78D6", "Supplier_Scorecard": "2A78D6",
                                       "Budget_vs_Actual": "2A78D6", "Assumptions": "FFC000", "README": "898781"}.get(sheet.title, "C3C2B7")

wb.calculation.fullCalcOnLoad = True    # Excel computes every formula on open
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print(f"Workbook written: {OUT}")
