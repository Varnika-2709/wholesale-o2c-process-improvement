# Wholesale Order-to-Cash Process Improvement

A business process analysis case study of the wholesale order-to-cash (O2C) process at
**PaddleCo**, a fictional paddle-sports brand that grew faster than its order process. It runs
the full cycle a process analyst would: discovery, current-state mapping, data analysis,
root cause, future-state design, SOPs, and a change management plan.

**The company and all data are synthetic.** 2,000 wholesale orders were simulated with seven
process problems planted in them, so the analysis has something real to find. The method is
the point; the figures are illustrations, not industry benchmarks.

## The questions

| # | Question | Where it's answered |
|---|---|---|
| 1 | Where do the 4.8 days between order receipt and shipment actually go? | `Analysis` sheet §1, [current-state swimlane](docs/current_state_swimlane.pdf) |
| 2 | Which order types and customers drive the delays? | `Analysis` sheet §2–3 |
| 3 | What is driving retailer chargebacks, and what do they cost? | `Analysis` sheet §4–5, chargeback Pareto |
| 4 | Which fixes give the most impact for the least effort? | Impact vs effort 2x2, automation backlog |
| 5 | What would the process look like after the fixes, and how will adoption be measured? | Future-state flow, `KPI_Scorecard`, change toolkit |

## Selected findings

- **Only 1.5 of the 4.8 days is value-added work** (pick, pack and ship at the 3PL). The rest
  is waiting: order entry and rework (1.6 days), allocation (1.05), routing (0.5), credit (0.18).
  Only **30%** of orders ship within the 3-day target.
- **Emailed PDF POs are 70% of orders** and wait 2.3 days for entry including rework; 7.8% need
  correction, against 0.7% of EDI orders.
- **Averages hide the tails.** Short-stock orders wait 3.4 days at allocation because priority is
  decided case by case on Slack, and credit holds average 26 hours, with 51% taking more than a day.
- **Two timing problems cause 95% of the $187K in chargebacks** over six months: routing
  requested after the retailer's window ($124K, 66%) and ASNs sent before the 3PL confirmed
  cartons ($54K, 29%). Neither is a people problem; both are about *when* a step fires.
- **The common root cause:** decisions that should be rules (price checks, allocation priority,
  credit release, ASN timing) were being made in email and Slack.
- **Modeled future state:** re-simulating the same orders with the fixes applied cuts cycle time
  to **2.3 days**, raises on-time shipping to **89%**, and cuts chargebacks **82%**, about
  **$325K a year** in chargebacks and labor avoided. Three KPIs land just short of target in the
  model (entry errors, OTIF, DTC oversells), which is the gap SOP audits and reinforcement close.

## What's in the process library

`docs/process_library.pdf` (27 pages):

| Section | Contents |
|---|---|
| Overview | Executive summary, findings → fixes → modeled results |
| 1 Discovery | SIPOC, stakeholder map (power vs interest), interview guide, 90-minute workshop agenda, synthesized findings |
| 2 Current state | Cross-functional swimlane, step inventory with value-added classification, RACI, fishbone and three 5 Whys chains, impact vs effort 2x2 |
| 3 Future state | Future-state flow, order prioritization decision tree, inventory allocation rules (reservations, tiers, FIFO), ATS formula and feed design, automation backlog |
| 4 SOPs | Library index plus six SOPs in one template: order entry, credit release, prioritization, routing, ASN compliance, invoicing |
| 5 Change toolkit | Two-wave rollout plan (ADKAR), role-by-role impact, launch email, FAQ, training agenda, quick reference guide, resistance plan, adoption measures |

## The workbook

`analysis/PaddleCo_O2C_Analysis.xlsx`: every result is a live formula over the order logs,
so changing an input on the `Assumptions` sheet recalculates the analysis.

| Sheet | Purpose |
|---|---|
| `Assumptions` | Labor rate, minutes per manual entry, rework time, KPI targets |
| `Analysis` | Days per stage, EDI vs email comparison, credit holds, chargeback Pareto, cost of inefficiency |
| `KPI_Scorecard` | Baseline vs modeled future vs target for eight KPIs |
| `Order_Log_Current` / `Order_Log_Future` | 2,000 orders each, shaped like a NetSuite saved-search export |
| `DTC_Oversell` | Weekly DTC cancellations caused by a stale weekly ATS spreadsheet |
| `ATS_Allocation_Sim` | Launch-SKU ATS calculation and FIFO allocation with reserved inventory |
| `Adoption_Tracker` | Post-launch training and audit tracker (example rows) |

## Repository contents

```
analysis/  PaddleCo_O2C_Analysis.xlsx   analysis workbook (formulas, charts, scorecard, simulator)
data/      order_log_current.csv        2,000 simulated orders, current process
           order_log_future.csv         the same orders with the fixes applied (modeled)
           dtc_weekly_oversell.csv      26 weeks of DTC orders and stock-out cancellations
           kpi_summary.json             headline KPIs for both scenarios
docs/      process_library.pdf          maps, analysis, specs, SOPs and change toolkit
           case_study_deck.pdf          11-slide case study presentation
           current_state_swimlane.pdf   current-state process map, standalone full-size
src/       generate_data.py             simulates the order logs (seeded, reproducible)
           build_workbook.py            builds the Excel workbook from the CSVs
```

## Reproducing it

```bash
pip install -r requirements.txt
python src/generate_data.py    # writes data/*.csv
python src/build_workbook.py   # writes analysis/PaddleCo_O2C_Analysis.xlsx
```

The generator is seeded, so it reproduces the same orders. Each order's attributes come from
one random stream and its process delays from a per-order stream, which keeps the current and
future scenarios on identical orders so the comparison isolates the effect of the fixes.
openpyxl writes formulas without cached values, so open the workbook in Excel or LibreOffice
(or recalculate) to see results.

## Concepts covered

O2C, P2P and NPI process scope · SIPOC · swimlane mapping · value-added analysis · RACI ·
fishbone and 5 Whys · Pareto analysis · impact vs effort prioritization · EDI 850 / 855 / 856 /
810 · ASN and retailer compliance chargebacks · routing guides · available-to-sell (ATS) and
inventory feeds · FIFO allocation with reservations · NetSuite saved searches, SuiteFlow and
bill-as-shipped · SOP design · ADKAR change management · adoption tracking

`Process mapping · Excel · Python · pandas · openpyxl · NetSuite concepts · change management`
