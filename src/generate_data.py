"""Synthetic PaddleCo wholesale O2C dataset with planted process problems.
Simulates current state and a modeled future state (same orders, fixes applied)."""
import numpy as np, pandas as pd, json, os
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
os.makedirs(OUT, exist_ok=True)
rng = np.random.default_rng(42)
N = 2000
start = pd.Timestamp("2026-04-01"); end = pd.Timestamp("2026-09-30")

bigbox = ["BigBox Sports A", "BigBox Outdoor B", "BigBox Club C"]
specialty = [f"Specialty Shop {i:03d}" for i in range(1, 151)]

def simulate(future=False, seed=42):
    ra = np.random.default_rng(seed)
    rows = []
    days = (end - start).days
    for i in range(N):
        r = np.random.default_rng([seed, i])
        edi = ra.random() < 0.30
        cust = ra.choice(bigbox) if edi else ra.choice(specialty)
        po_method = "EDI 850" if edi else "Email PDF"
        order_date = start + pd.Timedelta(days=int(ra.integers(0, days - 14)))
        lines = int(ra.integers(8, 30)) if edi else int(ra.integers(2, 12))
        units = int(lines * (ra.integers(60, 160) if edi else ra.integers(6, 24)))
        value = round(units * ra.uniform(11, 17), 2)
        # 1 Order entry
        if edi:
            entry_d = 0.0
            err = r.random() < 0.01
        else:
            entry_d = 0.0 if future else float(r.choice([0, 1, 2, 3, 4], p=[.08, .2, .3, .27, .15]))
            if future:  # PO capture + validation rules
                entry_d = float(r.choice([0, 1], p=[.8, .2]))
            err = r.random() < (0.02 if future else 0.09)
        rework_d = float(r.integers(1, 3)) if err else 0.0
        # 2 Credit
        hold = (not edi) and r.random() < 0.22
        hold_hrs = 0.0
        if hold:
            hold_hrs = float(np.clip(r.lognormal(np.log(3.0), 0.5), 1, 12)) if future \
                else float(np.clip(r.lognormal(np.log(24), 0.45), 4, 96))
        credit_d = hold_hrs / 24
        # 3 Allocation (short stock)
        short = r.random() < 0.25
        alloc_d = (float(r.choice([0, 1], p=[.85, .15])) if future else float(r.integers(2, 6))) if short \
            else float(r.choice([0, 1], p=[.7, .3]))
        # 4 Routing (EDI only)
        route_late = edi and r.random() < (0.06 if future else 0.30)
        route_d = (3.0 + float(r.integers(0, 2))) if route_late else (1.0 if edi else 0.0)
        # 5 Pick/pack/ship at 3PL
        ship_d = float(r.integers(1, 3))
        order_ts = order_date
        entered = order_ts + pd.Timedelta(days=entry_d + rework_d)
        released = entered + pd.Timedelta(days=credit_d)
        allocated = released + pd.Timedelta(days=alloc_d)
        routing_req = allocated + pd.Timedelta(days=0.5) if edi else pd.NaT
        shipped = allocated + pd.Timedelta(days=route_d + ship_d)
        # 6 ASN (EDI only)
        asn_early = edi and (not future) and r.random() < 0.18
        carton_confirm = shipped if edi else pd.NaT
        asn_sent = (shipped - pd.Timedelta(hours=6) if asn_early else shipped + pd.Timedelta(hours=2)) if edi else pd.NaT
        # 7 Invoice
        split = r.random() < 0.20
        inv_lag = (0.0 if r.random() < .7 else 1.0) if future else \
            (float(r.integers(3, 7)) if split else float(r.choice([0, 1, 2], p=[.3, .45, .25])))
        invoiced = shipped + pd.Timedelta(days=inv_lag)
        # Chargebacks
        cb_type, cb_amt = "", 0.0
        if edi:
            if asn_early and r.random() < 0.65:
                cb_type, cb_amt = "ASN accuracy", round(value * 0.03, 2)
            elif route_late:
                cb_type, cb_amt = "Late routing / missed pickup", round(value * 0.03, 2)
            elif r.random() < (0.02 if future else 0.05):
                cb_type, cb_amt = "Carton labeling", round(max(150, value * 0.01), 2)
            elif (shipped - order_ts).days > 7 and r.random() < 0.5:
                cb_type, cb_amt = "Late shipment (OTIF)", round(value * 0.03, 2)
        cycle = (shipped - order_ts).total_seconds() / 86400
        window = 5 if edi else 7
        otif = (cycle <= window) and not err and r.random() > 0.03
        rows.append(dict(
            order_id=f"SO-{10000 + i}", customer=cust,
            channel="Wholesale-EDI" if edi else "Wholesale-Specialty", po_method=po_method,
            order_date=order_ts, entered_date=entered, credit_hold_flag="Y" if hold else "N",
            credit_hold_hours=round(hold_hrs, 1), short_stock_flag="Y" if short else "N",
            allocated_date=allocated, routing_requested_date=routing_req,
            routing_late_flag="Y" if route_late else "N", ship_date=shipped,
            carton_confirm_date=carton_confirm, asn_sent_date=asn_sent,
            asn_before_confirm_flag="Y" if asn_early else "N", split_shipment_flag="Y" if split else "N",
            invoice_date=invoiced, lines=lines, units=units, order_value=value,
            entry_error_flag="Y" if err else "N", chargeback_type=cb_type, chargeback_amount=cb_amt,
            cycle_days=round(cycle, 2), invoice_lag_days=inv_lag, otif_flag="Y" if otif else "N",
            # stage durations
            d_entry=entry_d + rework_d, d_credit=credit_d, d_alloc=alloc_d, d_route=route_d, d_ship=ship_d))
    return pd.DataFrame(rows)

cur = simulate(False); fut = simulate(True)

RATE = 35.0  # loaded $/hr for CS
def kpis(df, future=False):
    email = df[df.po_method == "Email PDF"]
    holds = df[df.credit_hold_flag == "Y"]
    rev = df.order_value.sum()
    entry_min = 3 if future else 12  # minutes per email order (capture tool vs manual keying)
    hours = len(email) * entry_min / 60 + (df.entry_error_flag == "Y").sum() * 0.75
    return {
        "orders": len(df), "revenue": round(rev, 0),
        "cycle_days": round(df.cycle_days.mean(), 2),
        "cycle_edi": round(df[df.po_method == "EDI 850"].cycle_days.mean(), 2),
        "cycle_email": round(email.cycle_days.mean(), 2),
        "pct_ship_3d": round((df.cycle_days <= 3).mean() * 100, 1),
        "entry_error_rate": round((email.entry_error_flag == "Y").mean() * 100, 1),
        "credit_hold_hrs": round(holds.credit_hold_hours.mean(), 1),
        "otif": round((df.otif_flag == "Y").mean() * 100, 1),
        "chargebacks": round(df.chargeback_amount.sum(), 0),
        "chargeback_pct_rev": round(df.chargeback_amount.sum() / rev * 100, 2),
        "invoice_lag": round(df.invoice_lag_days.mean(), 2),
        "manual_hours": round(hours, 0), "manual_cost": round(hours * RATE, 0),
        "stage": {k: round(df[k].mean(), 2) for k in ["d_entry", "d_credit", "d_alloc", "d_route", "d_ship"]},
        "cb_by_type": df[df.chargeback_amount > 0].groupby("chargeback_type").chargeback_amount.sum().round(0).sort_values(ascending=False).to_dict(),
        "cb_count": df[df.chargeback_amount > 0].groupby("chargeback_type").size().to_dict(),
    }

# DTC oversell: weekly
weeks = pd.date_range("2026-04-06", periods=26, freq="W-MON")
launch_weeks = set(list(range(8, 11)) + list(range(20, 23)))
dtc = []
r2 = np.random.default_rng(7)
for w, wk in enumerate(weeks):
    orders = int(r2.normal(1500, 120)) + (900 if w in launch_weeks else 0)
    rate = (0.085 if w in launch_weeks else 0.012) + r2.normal(0, 0.003)
    rate_f = (0.008 if w in launch_weeks else 0.003) + abs(r2.normal(0, 0.001))
    dtc.append(dict(week_start=wk.date(), launch_week="Y" if w in launch_weeks else "N", dtc_orders=orders,
                    oversold_cancellations=int(orders * max(rate, 0.002)),
                    modeled_cancellations_future=int(orders * rate_f)))
dtc = pd.DataFrame(dtc)

K = {"current": kpis(cur), "future": kpis(fut, True),
     "dtc_oversell_pct": round(dtc.oversold_cancellations.sum() / dtc.dtc_orders.sum() * 100, 2),
     "dtc_oversell_launch_pct": round(dtc[dtc.launch_week == "Y"].oversold_cancellations.sum() / dtc[dtc.launch_week == "Y"].dtc_orders.sum() * 100, 2),
     "dtc_oversell_future_pct": round(dtc.modeled_cancellations_future.sum() / dtc.dtc_orders.sum() * 100, 2),
     "dtc_cancellations": int(dtc.oversold_cancellations.sum()),
     "email_share": round((cur.po_method == "Email PDF").mean() * 100, 1),
     "hold_share_specialty": round((cur[cur.po_method == "Email PDF"].credit_hold_flag == "Y").mean() * 100, 1)}
cur.to_csv(f"{OUT}/order_log_current.csv", index=False)
fut.to_csv(f"{OUT}/order_log_future.csv", index=False)
dtc.to_csv(f"{OUT}/dtc_weekly_oversell.csv", index=False)
json.dump(K, open(f"{OUT}/kpi_summary.json", "w"), indent=1, default=str)
print(json.dumps(K, indent=1, default=str))
