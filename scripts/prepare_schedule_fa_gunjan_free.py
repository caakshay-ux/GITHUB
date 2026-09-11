#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Gunjan Narulkar — IBKR U22995548 (Free account).

Single source: since-inception Activity Statement (03-Dec-2025 – 19-Aug-2026).
GOOG 179 shares received via Internal transfer from U16931511 on 03-Dec-2025.
A3 Col C = Symbol only.
"""

from __future__ import annotations

import csv
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from prepare_schedule_fa_ray import (  # noqa: E402
    build_lots_as_of,
    extract_dividends,
    extract_interest,
    extract_open_positions,
    extract_orders,
    extract_stock_splits,
    extract_transfers,
    extract_withholding,
    merge_orders,
    merge_splits,
    normalize_symbol,
    parse_dt,
    parse_ibkr,
    to_usd,
    transfers_to_buy_orders,
    main,
)

SINCE = ROOT / "input_Since_Inception_Gunjan_Free.csv"
# Reuse same file as inception + fiscal; synthesize CY2025 annual
INCEPTION = SINCE
FISCAL = SINCE
ANNUAL = ROOT / "input_Annual_Statement_Gunjan_Free.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U22995548_AY2026-27.xlsx"


def _in_cy2025(d: date) -> bool:
    return date(2025, 1, 1) <= d <= date(2025, 12, 31)


def synthesize_annual(src: Path, annual_out: Path) -> None:
    _, s = parse_ibkr(src)
    open_pos = extract_open_positions(s)
    orders = merge_orders(
        extract_orders(s),
        transfers_to_buy_orders(extract_transfers(s), [open_pos]),
    )
    for o in orders:
        o["symbol"] = normalize_symbol(o["symbol"])
    splits = merge_splits(extract_stock_splits(s))
    lots_ye = build_lots_as_of(orders, splits, date(2025, 12, 31))

    rows: list[list] = []

    def add(sec, kind, *vals):
        rows.append([sec, kind, *vals])

    add("Statement", "Header", "Field Name", "Field Value")
    add("Statement", "Data", "Title", "Activity Statement")
    add("Statement", "Data", "Period", "January 1, 2025 - December 31, 2025")
    add("Statement", "Data", "WhenGenerated", "Synthesized from Free since-inception statement (opened 03-Dec-2025)")
    add("Account Information", "Header", "Field Name", "Field Value")
    for _, r in s["Account Information"]:
        add("Account Information", "Data", *r)

    add(
        "Open Positions",
        "Header",
        "DataDiscriminator",
        "Asset Category",
        "Currency",
        "Symbol",
        "Quantity",
        "Mult",
        "Cost Price",
        "Cost Basis",
        "Close Price",
        "Value",
        "Unrealized P/L",
        "Code",
    )
    stock_usd = 0.0
    mtm_rows = []
    # YE close proxy: use open position close from since-inception statement (later date) only as qty/cost;
    # for YE2025 price use cost proxy (no 31-Dec mark in file) — Notes caveat
    open_now = extract_open_positions(s)
    for sym, lots in sorted(lots_ye.items()):
        qty = sum(L.qty for L in lots)
        if qty <= 1e-8:
            continue
        cost = sum(L.cost_local for L in lots)
        ccy = lots[0].currency or "USD"
        px = (cost / qty) if qty else 0.0
        value = px * qty
        stock_usd += to_usd(value, ccy)
        add(
            "Open Positions",
            "Data",
            "Summary",
            "Stocks",
            ccy,
            sym,
            round(qty, 6),
            1,
            round(cost / qty, 9) if qty else 0,
            round(cost, 6),
            round(px, 6),
            round(value, 6),
            round(value - cost, 6),
            "",
        )
        mtm_rows.append((sym, qty, px))

    ending = round(stock_usd, 2)
    add("Net Asset Value", "Header", "Asset Class", "Prior Total", "Current Long", "Current Short", "Current Total", "Change")
    add("Net Asset Value", "Data", "Stock", 0, ending, 0, ending, ending)
    add("Net Asset Value", "Data", "Total", 0, ending, 0, ending, ending)
    add("Change in NAV", "Header", "Field Name", "Field Value")
    add("Change in NAV", "Data", "Starting Value", 0)
    add("Change in NAV", "Data", "Ending Value", ending)
    add(
        "Mark-to-Market Performance Summary",
        "Header",
        "Asset Category",
        "Symbol",
        "Prior Quantity",
        "Current Quantity",
        "Prior Price",
        "Current Price",
        "Mark-to-Market P/L Position",
        "Mark-to-Market P/L Transaction",
        "Mark-to-Market P/L Commissions",
        "Mark-to-Market P/L Other",
        "Mark-to-Market P/L Total",
        "Code",
    )
    for sym, qty, px in mtm_rows:
        add("Mark-to-Market Performance Summary", "Data", "Stocks", sym, 0, qty, "--", px, 0, 0, 0, 0, 0, "")

    add(
        "Transfers",
        "Header",
        "Asset Category",
        "Currency",
        "Symbol",
        "Date",
        "Type",
        "Direction",
        "Xfer Company",
        "Xfer Account",
        "Qty",
        "Xfer Price",
        "Market Value",
        "Realized P/L",
        "Cash Amount",
        "Code",
    )
    for kind, r in s.get("Transfers", []):
        if kind != "Data" or r[0] == "Total":
            continue
        # Normalize date to ISO for downstream
        r2 = list(r)
        try:
            r2[3] = parse_dt(r2[3]).isoformat()
        except Exception:
            pass
        add("Transfers", "Data", *r2)

    add(
        "Trades",
        "Header",
        "DataDiscriminator",
        "Asset Category",
        "Currency",
        "Symbol",
        "Date/Time",
        "Quantity",
        "T. Price",
        "C. Price",
        "Proceeds",
        "Comm/Fee",
        "Basis",
        "Realized P/L",
        "MTM P/L",
        "Code",
    )
    add("Corporate Actions", "Header", "Asset Category", "Currency", "Report Date", "Date/Time", "Description", "Quantity", "Proceeds", "Value", "Realized P/L", "Code")

    add("Dividends", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted([x for x in extract_dividends(s) if _in_cy2025(x["date"])], key=lambda x: x["date"]):
        add("Dividends", "Data", d["currency"], d["date"].isoformat(), d["description"], d["amount"])

    add("Withholding Tax", "Header", "Currency", "Date", "Description", "Amount", "Code")
    for w in sorted([x for x in extract_withholding(s) if _in_cy2025(x["date"])], key=lambda x: x["date"]):
        add("Withholding Tax", "Data", w["currency"], w["date"].isoformat(), w["description"], w["amount"], "")

    add("Deposits & Withdrawals", "Header", "Currency", "Settle Date", "Description", "Amount")
    add("Interest", "Header", "Currency", "Date", "Description", "Amount")

    add(
        "Financial Instrument Information",
        "Header",
        "Asset Category",
        "Symbol",
        "Description",
        "Conid",
        "Security ID",
        "Underlying",
        "Listing Exch",
        "Multiplier",
        "Type",
        "Code",
    )
    for kind, r in s.get("Financial Instrument Information", []):
        if kind == "Data":
            add("Financial Instrument Information", "Data", *r)

    annual_out.parent.mkdir(parents=True, exist_ok=True)
    with open(annual_out, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    print(f"Synthesized annual → {annual_out} ({len(rows)} rows, YE symbols={len(mtm_rows)}, NAV~{ending})")


NOTES = [
    ("Assessee", "Gunjan Narulkar — IBKR U22995548 (Free) — Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED"),
    ("Related account", "GOOG received via Internal transfer In from U16931511 on 03-Dec-2025 (see Paid workbook)"),
    ("Holdings", "Single security: GOOG × 179 throughout after transfer; no trades"),
    ("Capital gains", "Nil — no disposals"),
    (
        "YE2025 closing",
        "No Calendar-Year Annual statement; YE closing proxied at cost. Prefer IBKR mark / PortfolioAnalyst for filing.",
    ),
    ("Cost basis", "IBKR Open Positions cost basis USD 43,146.02 used for transfer-in lot"),
]


def run():
    synthesize_annual(SINCE, ANNUAL)
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-12-03 (Internal transfer In of GOOG from U16931511)",
        assessee_label="Gunjan Narulkar (U22995548 Free)",
        notes_extra=NOTES,
    )


if __name__ == "__main__":
    run()
