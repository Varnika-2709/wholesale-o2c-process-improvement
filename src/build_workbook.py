import pandas as pd, numpy as np, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
DATES = ["order_date", "entered_date", "allocated_date", "routing_requested_date", "ship_date",
         "carton_confirm_date", "asn_sent_date", "invoice_date"]
cur = pd.read_csv(f"{ROOT}/data/order_log_current.csv", parse_dates=DATES, keep_default_na=False, na_values=[""])
fut = pd.read_csv(f"{ROOT}/data/order_log_future.csv", parse_dates=DATES, keep_default_na=False, na_values=[""])
dtc = pd.read_csv(f"{ROOT}/data/dtc_weekly_oversell.csv", parse_dates=["week_start"])
for df in (cur, fut):
    df["chargeback_type"] = df["chargeback_type"].fillna("")
OUT = os.path.join(ROOT, "analysis", "PaddleCo_O2C_Analysis.xlsx")

F = "Arial"
H_FILL = PatternFill("solid", fgColor="1F3A5F"); H_FONT = Font(name=F, bold=True, color="FFFFFF", size=10)
BASE = Font(name=F, size=10); BOLD = Font(name=F, size=10, bold=True)
BLUE = Font(name=F, size=10, color="0000FF"); GREEN = Font(name=F, size=10, color="008000")
TITLE = Font(name=F, size=14, bold=True, color="1F3A5F"); SUB = Font(name=F, size=10, italic=True, color="595959")
YEL = PatternFill("solid", fgColor="FFFF00"); LIGHT = PatternFill("solid", fgColor="EEF3F8")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()

def header(ws, row, cols, start_col=1):
    for j, c in enumerate(cols):
        cell = ws.cell(row=row, column=start_col + j, value=c)
        cell.font = H_FONT; cell.fill = H_FILL; cell.alignment = Alignment(wrap_text=True, vertical="center"); cell.border = BOX

def title(ws, t, s=None):
    ws["A1"] = t; ws["A1"].font = TITLE
    if s: ws["A2"] = s; ws["A2"].font = SUB

def widths(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(i)].width = x

# ---------- README ----------
ws = wb.active; ws.title = "README"
title(ws, "PaddleCo Wholesale Order-to-Cash Analysis", "Portfolio case study for a Business Process Analyst role. All data is synthetic.")
rows = [
 ("What this is", "A simulated order log of 2,000 wholesale orders (Apr–Sep 2026) for a fictional paddle-sports brand, with seven process problems planted in it, plus a modeled future state after the proposed fixes."),
 ("Question it answers", "Why do wholesale orders miss the 3-day ship target, what does that cost, and which fixes matter most?"),
 ("How to use", "Change blue cells on the Assumptions sheet (yellow = key levers). Every result on Analysis, KPI_Scorecard and ATS_Allocation_Sim is a formula and recalculates."),
 ("Color legend", "Blue text = hardcoded input · Black = formula · Green = formula pulling from another sheet · Yellow fill = key assumption to adjust"),
 ("", ""),
 ("Sheet", "Purpose"),
 ("Assumptions", "Labor rate, minutes per manual entry, rework time, KPI targets"),
 ("Order_Log_Current", "2,000 orders as the process runs today (like a NetSuite saved-search export). Cycle time and invoice lag are formulas."),
 ("Order_Log_Future", "The same 2,000 orders re-simulated with the fixes in place. Modeled, not observed."),
 ("Analysis", "Where time goes by stage, EDI vs email POs, entry errors, credit holds, chargeback Pareto, cost of inefficiency"),
 ("KPI_Scorecard", "Baseline vs modeled future vs target for 8 KPIs"),
 ("DTC_Oversell", "Weekly DTC orders cancelled because the ATS spreadsheet was stale; launch weeks flagged"),
 ("ATS_Allocation_Sim", "Launch-week ATS calculation and FIFO allocation with reserved inventory. Edit inputs to see allocation change."),
 ("Adoption_Tracker", "Post-launch training and adoption tracker for the new order-entry process (example rows)"),
 ("", ""),
 ("Planted problems", "1 Email PDF POs keyed by hand · 2 No allocation priority rule · 3 Single credit-hold approver · 4 Weekly ATS spreadsheet · 5 Late routing requests · 6 ASN sent before carton confirmation · 7 Late invoicing of split shipments"),
 ("Data caveat", "Synthetic data generated with seeded random delays (Python). Treat figures as illustrations of method, not industry benchmarks."),
]
for i, (a, b) in enumerate(rows, 4):
    ws.cell(row=i, column=1, value=a).font = BOLD; c = ws.cell(row=i, column=2, value=b); c.font = BASE; c.alignment = WRAP
for r in (9,):
    for col in (1, 2): ws.cell(row=r, column=col).font = H_FONT; ws.cell(row=r, column=col).fill = H_FILL
widths(ws, [22, 110])

# ---------- Assumptions ----------
A = wb.create_sheet("Assumptions")
title(A, "Assumptions and targets", "Blue = input you can change. Yellow = key lever.")
header(A, 4, ["Assumption", "Value", "Unit", "Source / note"])
assum = [
 ("Loaded hourly cost, Customer Service", 35, "$/hour", "Assumption: wage + benefits for a CS rep; replace with your own figure", True),
 ("Minutes to key one email PDF order (today)", 12, "minutes", "Assumption from time study; validate by shadowing CS", True),
 ("Minutes per email order with PO capture tool", 3, "minutes", "Assumption: review-and-confirm only", True),
 ("Rework hours per order-entry error", 0.75, "hours", "Assumption: find, correct, re-confirm with customer", False),
 ("Ship target, all wholesale orders", 3, "days", "Business target from the problem statement", False),
 ("Target: entry error rate (email POs)", 0.02, "%", "Project target", False),
 ("Target: credit hold release time", 4, "hours", "Project target", False),
 ("Target: OTIF", 0.95, "%", "Project target", False),
 ("Target: chargebacks as % of wholesale revenue", 0.005, "%", "Project target", False),
 ("Target: invoice lag", 0.5, "days", "Project target", False),
 ("Target: DTC oversell rate", 0.005, "%", "Project target", False),
 ("Target: share of orders shipped within target", 0.85, "%", "Project target", False),
]
for i, (n, v, u, s, key) in enumerate(assum, 5):
    A.cell(row=i, column=1, value=n).font = BASE
    c = A.cell(row=i, column=2, value=v); c.font = BLUE
    if u == "%": c.number_format = "0.0%"
    if key: c.fill = YEL
    A.cell(row=i, column=3, value=u).font = BASE
    A.cell(row=i, column=4, value=s).font = SUB
widths(A, [46, 12, 10, 70])
# named refs
AS = {k: f"Assumptions!$B${i}" for i, k in enumerate(
    ["rate", "min_now", "min_fut", "rework", "ship_t", "t_err", "t_hold", "t_otif", "t_cb", "t_inv", "t_dtc", "t_ship"], 5)}

# ---------- Order logs ----------
cols = ["order_id", "customer", "channel", "po_method", "order_date", "entered_date", "credit_hold_flag",
        "credit_hold_hours", "short_stock_flag", "allocated_date", "routing_requested_date", "routing_late_flag",
        "ship_date", "carton_confirm_date", "asn_sent_date", "asn_before_confirm_flag", "split_shipment_flag",
        "invoice_date", "lines", "units", "order_value", "entry_error_flag", "chargeback_type", "chargeback_amount",
        "d_entry", "d_credit", "d_alloc", "d_route", "d_ship"]
# formula columns appended: cycle_days, invoice_lag_days, shipped_within_target, otif_flag
def write_log(name, df, sub):
    ws = wb.create_sheet(name)
    ws["A1"] = name.replace("_", " "); ws["A1"].font = TITLE; ws["A2"] = sub; ws["A2"].font = SUB
    allc = cols + ["cycle_days", "invoice_lag_days", "within_target", "otif_flag"]
    header(ws, 3, allc)
    L = {c: get_column_letter(i + 1) for i, c in enumerate(allc)}
    datecols = {"order_date", "entered_date", "allocated_date", "routing_requested_date", "ship_date",
                "carton_confirm_date", "asn_sent_date", "invoice_date"}
    for ri, rec in enumerate(df[cols].itertuples(index=False), 4):
        for ci, (c, v) in enumerate(zip(cols, rec), 1):
            if isinstance(v, float) and np.isnan(v): v = None
            if c in datecols:
                v = None if pd.isna(v) else pd.Timestamp(v).to_pydatetime()
            cell = ws.cell(row=ri, column=ci, value=v)
            if c in datecols and v is not None: cell.number_format = "yyyy-mm-dd hh:mm"
            if c in ("order_value", "chargeback_amount"): cell.number_format = "$#,##0.00;($#,##0.00);-"
            if c.startswith("d_"): cell.number_format = "0.00"
        n = len(cols)
        ws.cell(row=ri, column=n + 1, value=f"={L['ship_date']}{ri}-{L['order_date']}{ri}").number_format = "0.00"
        ws.cell(row=ri, column=n + 2, value=f"=INT({L['invoice_date']}{ri})-INT({L['ship_date']}{ri})").number_format = "0"
        ws.cell(row=ri, column=n + 3, value=f'=IF({L["cycle_days"]}{ri}<={AS["ship_t"]},"Y","N")')
        ws.cell(row=ri, column=n + 4, value=v_otif[name][ri - 4])
    ws.freeze_panes = "B4"; ws.auto_filter.ref = f"A3:{get_column_letter(len(allc))}{len(df) + 3}"
    for i, c in enumerate(allc, 1):
        ws.column_dimensions[get_column_letter(i)].width = 18 if ("date" in c or c == "customer" or c == "chargeback_type") else 13
    return L, len(df) + 3

v_otif = {"Order_Log_Current": list(cur.otif_flag), "Order_Log_Future": list(fut.otif_flag)}
Lc, lastc = write_log("Order_Log_Current", cur, "Current state, Apr–Sep 2026, synthetic. Columns d_* = days spent in each stage. otif_flag = shipped complete inside the customer window (EDI 5 days, specialty 7).")
Lf, lastf = write_log("Order_Log_Future", fut, "MODELED future state: same orders, fixes applied (PO capture, credit workflow, allocation rules, routing alerts, ASN gating, auto-invoice).")
# fix: the per-row otif note for future sheet
for ws_name in ("Order_Log_Current", "Order_Log_Future"):
    for c in wb[ws_name][3]: c.font = H_FONT

def rng(L, last, col, sheet="Order_Log_Current"):
    return f"{sheet}!${L[col]}$4:${L[col]}${last}"

C = lambda col: rng(Lc, lastc, col, "Order_Log_Current")
FU = lambda col: rng(Lf, lastf, col, "Order_Log_Future")

# ---------- Analysis ----------
an = wb.create_sheet("Analysis", 2)
title(an, "Analysis: where the time and money go", "All figures are formulas over Order_Log_Current (and Order_Log_Future for comparison).")
r = 4
an.cell(row=r, column=1, value="1. Average days per stage (order received → shipped)").font = BOLD; r += 1
header(an, r, ["Stage", "Current (days)", "Modeled future (days)", "Change (days)", "Planted problem"]); r += 1
stages = [("Order entry incl. rework", "d_entry", "PDF POs keyed by hand; errors re-worked"),
          ("Credit hold", "d_credit", "Single approver, email checked twice a day"),
          ("Allocation", "d_alloc", "No written priority rule when stock is short"),
          ("Routing wait (EDI)", "d_route", "Routing requested late, pickups missed"),
          ("Pick, pack, ship (3PL)", "d_ship", "Not targeted in this project")]
s1 = r
for name, col, note in stages:
    an.cell(row=r, column=1, value=name).font = BASE
    an.cell(row=r, column=2, value=f"=AVERAGE({C(col)})").font = GREEN
    an.cell(row=r, column=3, value=f"=AVERAGE({FU(col)})").font = GREEN
    an.cell(row=r, column=4, value=f"=C{r}-B{r}").font = BASE
    an.cell(row=r, column=5, value=note).font = SUB
    for cc in (2, 3, 4): an.cell(row=r, column=cc).number_format = "0.00;-0.00"
    r += 1
an.cell(row=r, column=1, value="Total (≈ average cycle time)").font = BOLD
for cc, L_ in ((2, "B"), (3, "C"), (4, "D")):
    an.cell(row=r, column=cc, value=f"=SUM({L_}{s1}:{L_}{r-1})").font = BOLD; an.cell(row=r, column=cc).number_format = "0.00;-0.00"
s1_end = r - 1; r += 2

ch = BarChart(); ch.type = "bar"; ch.style = 10; ch.title = "Days per stage: current vs modeled future"
ch.add_data(Reference(an, min_col=2, max_col=3, min_row=s1 - 1, max_row=s1_end), titles_from_data=True)
ch.set_categories(Reference(an, min_col=1, min_row=s1, max_row=s1_end)); ch.height = 7; ch.width = 16
an.add_chart(ch, "I4")

an.cell(row=r, column=1, value="2. Segment comparison: EDI vs email PDF orders").font = BOLD; r += 1
header(an, r, ["Segment", "Orders", "Share of orders", "Avg cycle (days)", "% shipped within target", "Entry error rate", "Avg entry days"]); r += 1
s2 = r
for seg in ("EDI 850", "Email PDF"):
    an.cell(row=r, column=1, value=seg).font = BASE
    an.cell(row=r, column=2, value=f'=COUNTIFS({C("po_method")},A{r})').font = GREEN
    an.cell(row=r, column=3, value=f"=B{r}/SUM($B${s2}:$B${s2+1})").number_format = "0.0%"
    an.cell(row=r, column=4, value=f'=AVERAGEIFS({C("cycle_days")},{C("po_method")},A{r})').number_format = "0.00"
    an.cell(row=r, column=5, value=f'=COUNTIFS({C("po_method")},A{r},{C("within_target")},"Y")/B{r}').number_format = "0.0%"
    an.cell(row=r, column=6, value=f'=COUNTIFS({C("po_method")},A{r},{C("entry_error_flag")},"Y")/B{r}').number_format = "0.0%"
    an.cell(row=r, column=7, value=f'=AVERAGEIFS({C("d_entry")},{C("po_method")},A{r})').number_format = "0.00"
    r += 1
r += 1
an.cell(row=r, column=1, value="3. Credit holds and short-stock allocation").font = BOLD; r += 1
header(an, r, ["Measure", "Value", "Note"]); r += 1
items = [
 ("Specialty orders placed on credit hold", f'=COUNTIFS({C("credit_hold_flag")},"Y")/COUNTIFS({C("po_method")},"Email PDF")', "0.0%", "Holds apply to specialty accounts only"),
 ("Avg hours to release a hold", f'=AVERAGEIFS({C("credit_hold_hours")},{C("credit_hold_flag")},"Y")', "0.0", "Target on Assumptions"),
 ("Holds taking longer than 24 hours", f'=COUNTIFS({C("credit_hold_hours")},">24")/COUNTIFS({C("credit_hold_flag")},"Y")', "0.0%", "Two email checks a day push most past a full day"),
 ("Avg allocation days, short-stock orders", f'=AVERAGEIFS({C("d_alloc")},{C("short_stock_flag")},"Y")', "0.00", "Waiting on Sales via Slack"),
 ("Avg allocation days, in-stock orders", f'=AVERAGEIFS({C("d_alloc")},{C("short_stock_flag")},"N")', "0.00", ""),
 ("Avg cycle days, held orders", f'=AVERAGEIFS({C("cycle_days")},{C("credit_hold_flag")},"Y")', "0.00", ""),
 ("Avg cycle days, orders not held", f'=AVERAGEIFS({C("cycle_days")},{C("credit_hold_flag")},"N")', "0.00", ""),
]
for n, f, fmt, note in items:
    an.cell(row=r, column=1, value=n).font = BASE
    c = an.cell(row=r, column=2, value=f); c.number_format = fmt; c.font = GREEN
    an.cell(row=r, column=3, value=note).font = SUB; r += 1
r += 1
an.cell(row=r, column=1, value="4. Chargeback Pareto (current state, 6 months)").font = BOLD; r += 1
header(an, r, ["Chargeback type", "Count", "Amount ($)", "% of total", "Cumulative %", "Root cause"]); r += 1
cbt = cur[cur.chargeback_amount > 0].groupby("chargeback_type").chargeback_amount.sum().sort_values(ascending=False).index.tolist()
causes = {"Late routing / missed pickup": "Routing requested after the retailer's window", "ASN accuracy": "ASN sent before 3PL confirmed cartons",
          "Carton labeling": "Label spec not checked pre-ship", "Late shipment (OTIF)": "Upstream delays push ship date past window"}
s4 = r
for t in cbt:
    an.cell(row=r, column=1, value=t).font = BASE
    an.cell(row=r, column=2, value=f'=COUNTIFS({C("chargeback_type")},A{r})').font = GREEN
    c = an.cell(row=r, column=3, value=f'=SUMIFS({C("chargeback_amount")},{C("chargeback_type")},A{r})'); c.font = GREEN; c.number_format = "$#,##0"
    an.cell(row=r, column=4, value=f"=C{r}/SUM($C${s4}:$C${s4+len(cbt)-1})").number_format = "0.0%"
    an.cell(row=r, column=5, value=f"=SUM($C${s4}:C{r})/SUM($C${s4}:$C${s4+len(cbt)-1})").number_format = "0.0%"
    an.cell(row=r, column=6, value=causes.get(t, "")).font = SUB; r += 1
an.cell(row=r, column=1, value="Total").font = BOLD
c = an.cell(row=r, column=3, value=f"=SUM(C{s4}:C{r-1})"); c.font = BOLD; c.number_format = "$#,##0"
s4_end = r - 1; r += 2
pc = BarChart(); pc.title = "Chargebacks by type ($)"; pc.style = 10
pc.add_data(Reference(an, min_col=3, min_row=s4 - 1, max_row=s4_end), titles_from_data=True)
pc.set_categories(Reference(an, min_col=1, min_row=s4, max_row=s4_end)); pc.legend = None
lc = LineChart(); lc.add_data(Reference(an, min_col=5, min_row=s4 - 1, max_row=s4_end), titles_from_data=True)
lc.y_axis.axId = 200; lc.y_axis.title = "Cumulative %"; lc.y_axis.crosses = "max"; lc.y_axis.number_format = "0%"
pc += lc; pc.height = 7; pc.width = 16
an.add_chart(pc, "I22")

an.cell(row=r, column=1, value="5. Cost of inefficiency (6 months)").font = BOLD; r += 1
header(an, r, ["Cost driver", "Current", "Modeled future", "Saving", "How calculated"]); r += 1
s5 = r
cost = [
 ("Manual order-entry hours", f'=COUNTIFS({C("po_method")},"Email PDF")*{AS["min_now"]}/60',
  f'=COUNTIFS({FU("po_method")},"Email PDF")*{AS["min_fut"]}/60', "0", "Email orders × minutes per order ÷ 60"),
 ("Rework hours from entry errors", f'=COUNTIFS({C("entry_error_flag")},"Y")*{AS["rework"]}',
  f'=COUNTIFS({FU("entry_error_flag")},"Y")*{AS["rework"]}', "0", "Errors × rework hours"),
 ("Labor cost of the above ($)", f"=(B{r}+B{r+1})*{AS['rate']}", f"=(C{r}+C{r+1})*{AS['rate']}", "$#,##0", "Hours × loaded rate"),
 ("Compliance chargebacks ($)", f'=SUM({C("chargeback_amount")})', f'=SUM({FU("chargeback_amount")})', "$#,##0", "Sum of deductions"),
 ("Total quantified cost ($)", f"=B{r+2}+B{r+3}", f"=C{r+2}+C{r+3}", "$#,##0", "Labor + chargebacks"),
 ("Annualized ($)", f"=B{r+4}*2", f"=C{r+4}*2", "$#,##0", "6 months × 2"),
]
for n, f1, f2, fmt, how in cost:
    an.cell(row=r, column=1, value=n).font = BOLD if "Total" in n or "Annual" in n else BASE
    for cc, f in ((2, f1), (3, f2)):
        c = an.cell(row=r, column=cc, value=f); c.number_format = fmt; c.font = GREEN
    c = an.cell(row=r, column=4, value=f"=B{r}-C{r}"); c.number_format = fmt
    an.cell(row=r, column=5, value=how).font = SUB; r += 1
an.cell(row=r + 1, column=1, value="Not quantified here: working capital tied up by late invoices, sales lost to DTC oversells, management time spent on Slack escalations.").font = SUB
widths(an, [42, 16, 18, 16, 22, 16, 16])

# ---------- KPI Scorecard ----------
k = wb.create_sheet("KPI_Scorecard", 3)
title(k, "KPI scorecard: baseline vs modeled future vs target", "Baseline = Order_Log_Current. Future = Order_Log_Future (modeled). Status compares future to target.")
header(k, 4, ["KPI", "Definition", "Baseline", "Modeled future", "Target", "Better if", "Status"])
kp = [
 ("Order-to-ship cycle time (days)", "Avg days from order received to shipped", f'=AVERAGE({C("cycle_days")})', f'=AVERAGE({FU("cycle_days")})', f'={AS["ship_t"]}', "lower", "0.00"),
 ("Shipped within target", "% of orders shipped within the target days", f'=COUNTIFS({C("within_target")},"Y")/COUNTA({C("order_id")})', f'=COUNTIFS({FU("within_target")},"Y")/COUNTA({FU("order_id")})', f'={AS["t_ship"]}', "higher", "0.0%"),
 ("Order entry error rate", "Email orders needing correction ÷ email orders", f'=COUNTIFS({C("po_method")},"Email PDF",{C("entry_error_flag")},"Y")/COUNTIFS({C("po_method")},"Email PDF")', f'=COUNTIFS({FU("po_method")},"Email PDF",{FU("entry_error_flag")},"Y")/COUNTIFS({FU("po_method")},"Email PDF")', f'={AS["t_err"]}', "lower", "0.0%"),
 ("Credit hold release time (hours)", "Avg hours a held order waits for release", f'=AVERAGEIFS({C("credit_hold_hours")},{C("credit_hold_flag")},"Y")', f'=AVERAGEIFS({FU("credit_hold_hours")},{FU("credit_hold_flag")},"Y")', f'={AS["t_hold"]}', "lower", "0.0"),
 ("On-time-in-full (OTIF)", "Shipped complete inside the customer window", f'=COUNTIFS({C("otif_flag")},"Y")/COUNTA({C("order_id")})', f'=COUNTIFS({FU("otif_flag")},"Y")/COUNTA({FU("order_id")})', f'={AS["t_otif"]}', "higher", "0.0%"),
 ("Chargebacks, % of wholesale revenue", "Deductions ÷ order value", f'=SUM({C("chargeback_amount")})/SUM({C("order_value")})', f'=SUM({FU("chargeback_amount")})/SUM({FU("order_value")})', f'={AS["t_cb"]}', "lower", "0.00%"),
 ("Invoice lag (days)", "Avg days from ship to invoice", f'=AVERAGE({C("invoice_lag_days")})', f'=AVERAGE({FU("invoice_lag_days")})', f'={AS["t_inv"]}', "lower", "0.00"),
 ("DTC oversell rate", "Orders cancelled for no stock ÷ DTC orders", "=SUM(DTC_Oversell!$D$5:$D$30)/SUM(DTC_Oversell!$C$5:$C$30)", "=SUM(DTC_Oversell!$E$5:$E$30)/SUM(DTC_Oversell!$C$5:$C$30)", f'={AS["t_dtc"]}', "lower", "0.00%"),
]
for i, (n, d, b, f, t, better, fmt) in enumerate(kp, 5):
    k.cell(row=i, column=1, value=n).font = BOLD; k.cell(row=i, column=2, value=d).font = SUB
    for cc, v in ((3, b), (4, f), (5, t)):
        c = k.cell(row=i, column=cc, value=v); c.number_format = fmt; c.font = GREEN
    k.cell(row=i, column=6, value=better).font = BASE
    k.cell(row=i, column=7, value=f'=IF(F{i}="lower",IF(D{i}<=E{i},"Meets target","Gap"),IF(D{i}>=E{i},"Meets target","Gap"))')
k.conditional_formatting.add("G5:G12", CellIsRule(operator="equal", formula=['"Meets target"'], font=Font(name=F, color="008000", bold=True)))
k.conditional_formatting.add("G5:G12", CellIsRule(operator="equal", formula=['"Gap"'], font=Font(name=F, color="C00000", bold=True)))
k["A14"] = "Read the future column as a model of what the fixes could achieve, to be replaced by real post-launch actuals in weeks 1, 2 and 4."; k["A14"].font = SUB
widths(k, [36, 44, 12, 15, 10, 9, 15])

# ---------- DTC oversell ----------
d = wb.create_sheet("DTC_Oversell")
title(d, "DTC oversells from a stale ATS feed", "Weekly. ATS updated once a week from a spreadsheet; launch weeks drain stock between updates.")
header(d, 4, ["Week starting", "Launch week", "DTC orders", "Cancelled: no stock (current)", "Cancelled: modeled with 15-min ATS feed", "Oversell rate (current)"])
for i, rec in enumerate(dtc.itertuples(index=False), 5):
    d.cell(row=i, column=1, value=pd.Timestamp(rec.week_start).to_pydatetime()).number_format = "yyyy-mm-dd"
    d.cell(row=i, column=2, value=rec.launch_week)
    for cc, v in ((3, rec.dtc_orders), (4, rec.oversold_cancellations), (5, rec.modeled_cancellations_future)):
        d.cell(row=i, column=cc, value=int(v)).font = BLUE
    d.cell(row=i, column=6, value=f"=D{i}/C{i}").number_format = "0.0%"
d["A32"] = "Totals"; d["A32"].font = BOLD
for col in "CDE": d[f"{col}32"] = f"=SUM({col}5:{col}30)"; d[f"{col}32"].font = BOLD
d["F32"] = "=D32/C32"; d["F32"].number_format = "0.0%"
d["A33"] = "Launch weeks"; d["A33"].font = BOLD
d["F33"] = '=SUMIFS(D5:D30,B5:B30,"Y")/SUMIFS(C5:C30,B5:B30,"Y")'; d["F33"].number_format = "0.0%"
d["A34"] = "Other weeks"; d["A34"].font = BOLD
d["F34"] = '=SUMIFS(D5:D30,B5:B30,"N")/SUMIFS(C5:C30,B5:B30,"N")'; d["F34"].number_format = "0.0%"
lch = LineChart(); lch.title = "Weekly oversell rate"; lch.style = 10
lch.add_data(Reference(d, min_col=6, min_row=4, max_row=30), titles_from_data=True)
lch.set_categories(Reference(d, min_col=1, min_row=5, max_row=30)); lch.y_axis.number_format = "0%"; lch.height = 7; lch.width = 16
d.add_chart(lch, "H4")
widths(d, [14, 12, 12, 18, 22, 16])

# ---------- ATS & allocation simulator ----------
s = wb.create_sheet("ATS_Allocation_Sim")
title(s, "Launch SKU: ATS calculation and FIFO allocation with reservations", "Edit blue inputs. Reserved units go only to the named account; everything else is allocated first-in-first-out by order date.")
header(s, 4, ["Input (SKU: PRO-16 launch paddle)", "Units", "Note"])
inputs = [("On hand at 3PL", 1200, "From WMS snapshot"), ("Reserved for BigBox Sports A launch", 400, "Launch commitment, expires on release date"),
          ("Safety stock", 100, "Buffer for damages and count errors"), ("Inbound arriving within 14 days", 600, "From open PO receipts"),
          ("Count inbound in ATS? (1 = yes, 0 = no)", 0, "Policy choice; 0 is the safer default for DTC")]
for i, (n, v, note) in enumerate(inputs, 5):
    s.cell(row=i, column=1, value=n).font = BASE; c = s.cell(row=i, column=2, value=v); c.font = BLUE
    if i in (6, 9): c.fill = YEL
    s.cell(row=i, column=3, value=note).font = SUB
s["A11"] = "Available-to-sell (ATS)"; s["A11"].font = BOLD
s["B11"] = "=MAX(0,B5-B6-B7+B8*B9)"; s["B11"].font = BOLD
s["C11"] = "ATS = On hand − Reserved − Safety stock + Inbound × flag (allocations below draw it down)"; s["C11"].font = SUB
header(s, 13, ["Order", "Customer", "Order date", "Qty ordered", "Uses reservation?", "Allocated", "Short", "Fill rate", "Remaining ATS after", "Remaining reservation after"])
queue = [("SO-20001", "Specialty Shop 012", "2026-10-01", 60, "N"), ("SO-20002", "BigBox Sports A", "2026-10-01", 350, "Y"),
         ("SO-20003", "DTC website (daily batch)", "2026-10-02", 220, "N"), ("SO-20004", "Specialty Shop 087", "2026-10-02", 48, "N"),
         ("SO-20005", "BigBox Outdoor B", "2026-10-03", 300, "N"), ("SO-20006", "BigBox Sports A", "2026-10-03", 120, "Y"),
         ("SO-20007", "Specialty Shop 140", "2026-10-04", 36, "N"), ("SO-20008", "DTC website (daily batch)", "2026-10-04", 260, "N"),
         ("SO-20009", "Specialty Shop 031", "2026-10-05", 72, "N"), ("SO-20010", "Marketplace", "2026-10-05", 150, "N")]
dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=False); s.add_data_validation(dv)
for i, (o, cu, dt, q, res) in enumerate(queue, 14):
    s.cell(row=i, column=1, value=o).font = BLUE; s.cell(row=i, column=2, value=cu).font = BLUE
    c = s.cell(row=i, column=3, value=pd.Timestamp(dt).to_pydatetime()); c.number_format = "yyyy-mm-dd"; c.font = BLUE
    s.cell(row=i, column=4, value=q).font = BLUE
    c = s.cell(row=i, column=5, value=res); c.font = BLUE; dv.add(c)
    prev_ats = "$B$11" if i == 14 else f"I{i-1}"; prev_res = "$B$6" if i == 14 else f"J{i-1}"
    s.cell(row=i, column=6, value=f'=IF(E{i}="Y",MIN(D{i},{prev_res}+{prev_ats}),MIN(D{i},{prev_ats}))')
    s.cell(row=i, column=7, value=f"=D{i}-F{i}")
    s.cell(row=i, column=8, value=f"=F{i}/D{i}").number_format = "0%"
    s.cell(row=i, column=9, value=f'=IF(E{i}="Y",{prev_ats}-MAX(0,F{i}-{prev_res}),{prev_ats}-F{i})')
    s.cell(row=i, column=10, value=f'=IF(E{i}="Y",MAX(0,{prev_res}-F{i}),{prev_res})')
s["A25"] = "Totals"; s["A25"].font = BOLD
for col in "DFG": s[f"{col}25"] = f"=SUM({col}14:{col}23)"; s[f"{col}25"].font = BOLD
s["H25"] = "=F25/D25"; s["H25"].number_format = "0%"; s["H25"].font = BOLD
notes = ["How to read it:",
         "• Orders are already sorted oldest first, so filling top-down is FIFO by order date.",
         "• A reserved order draws on its reservation first, then on general ATS if the reservation runs out.",
         "• Try setting B9 to 1 (count inbound) or B6 to 0 (no reservation) and watch who gets shorted.",
         "• Exercise: add a priority column (key account / cancel date) and sort by it before order date to model the prioritization decision tree."]
for i, n in enumerate(notes, 27): s.cell(row=i, column=1, value=n).font = BOLD if i == 27 else BASE
widths(s, [40, 26, 12, 12, 16, 11, 9, 9, 18, 22])

# ---------- Adoption tracker ----------
t = wb.create_sheet("Adoption_Tracker")
title(t, "Adoption tracker: new email-PO capture and validation process", "Example rows show the format; replace with real team members after go-live.")
header(t, 4, ["Team member", "Role", "Trained (date)", "First order in new flow", "Audit sample (orders)", "Audit passes", "Pass rate", "Blockers raised", "Status"])
ex = [("Rep A", "CS rep", "2026-10-05", "2026-10-06", 10, 10, "None", "Adopted"), ("Rep B", "CS rep", "2026-10-05", "2026-10-07", 10, 9, "Unsure on ship-to overrides", "Adopted"),
      ("Rep C", "CS rep", "2026-10-05", "2026-10-09", 10, 7, "Still keys some PDFs by hand", "Needs coaching"), ("Rep D", "CS rep", "2026-10-07", "", 0, 0, "On leave launch week", "Not started"),
      ("Lead E", "CS lead", "2026-10-02", "2026-10-05", 10, 10, "None", "Adopted"), ("Analyst F", "Finance AR", "2026-10-02", "2026-10-05", 8, 8, "Workflow email volume", "Adopted"),
      ("Rep G", "Sales ops", "2026-10-06", "2026-10-08", 6, 5, "Wants mobile approvals", "Adopted")]
stv = DataValidation(type="list", formula1='"Not started,Needs coaching,Adopted"'); t.add_data_validation(stv)
for i, (n, ro, tr, fo, a, p, b, st) in enumerate(ex, 5):
    vals = [n, ro, pd.Timestamp(tr).to_pydatetime(), pd.Timestamp(fo).to_pydatetime() if fo else None, a, p]
    for cc, v in enumerate(vals, 1):
        c = t.cell(row=i, column=cc, value=v); c.font = BLUE
        if cc in (3, 4): c.number_format = "yyyy-mm-dd"
    t.cell(row=i, column=7, value=f'=IF(E{i}=0,"",F{i}/E{i})').number_format = "0%"
    t.cell(row=i, column=8, value=b).font = BLUE
    c = t.cell(row=i, column=9, value=st); c.font = BLUE; stv.add(c)
t["A14"] = "Adoption rate"; t["A14"].font = BOLD
t["B14"] = '=COUNTIF(I5:I11,"Adopted")/COUNTA(A5:A11)'; t["B14"].number_format = "0%"; t["B14"].font = BOLD
t["A15"] = "Audit pass rate"; t["A15"].font = BOLD
t["B15"] = "=SUM(F5:F11)/SUM(E5:E11)"; t["B15"].number_format = "0%"; t["B15"].font = BOLD
t["A16"] = "Target"; t["B16"] = 0.9; t["B16"].number_format = "0%"; t["B16"].font = BLUE
widths(t, [16, 12, 14, 18, 16, 12, 10, 30, 15])

# fonts default for any unstyled cells
for wsx in wb.worksheets:
    for row in wsx.iter_rows():
        for c in row:
            if c.value is not None and c.font.name != F:
                c.font = Font(name=F, size=10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)
wb.move_sheet("Assumptions", offset=-(wb.sheetnames.index("Assumptions") - 1))
wb.save(OUT); print("saved", OUT, wb.sheetnames)
