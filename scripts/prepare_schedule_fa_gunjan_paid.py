#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Gunjan Narulkar — IBKR U16931511 (Paid account).

Inputs: FY2024-25 (31-Dec-2024–31-Mar-2025) + FY2025-26.
CY2025 Annual synthesized. Opening GOOG 216 shares seeded (held before statement start).
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
    parse_ibkr,
    to_usd,
    transfers_to_buy_orders,
    main,
)

INCEPTION = ROOT / "input_Inception_FY2024-25_Gunjan_Paid.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Gunjan_Paid.csv"
ANNUAL = ROOT / "input_Annual_Statement_Gunjan_Paid.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U16931511_AY2026-27.xlsx"

# GOOG opening (MTM prior qty 216 at 31-Dec-2024). Cost backsolved from Jan-2025
# sell bases + Mar-2025 open cost after FOP in (see Notes).
GOOG_SEED_QTY = 216.0
GOOG_SEED_COST = 19300.32


def _in_cy2025(d: date) -> bool:
    return date(2025, 1, 1) <= d <= date(2025, 12, 31)


def goog_seed_order():
    return {
        "currency": "USD",
        "symbol": "GOOG",
        "datetime": "2024-12-31, 00:00:00",
        "date": date(2024, 12, 31),
        "qty": GOOG_SEED_QTY,
        "price": GOOG_SEED_COST / GOOG_SEED_QTY,
        "proceeds": -GOOG_SEED_COST,
        "comm_usd": 0.0,
        "basis": GOOG_SEED_COST,
        "realized": 0.0,
        "code": "O",
    }


def synthesize_annual(inception: Path, fiscal: Path, annual_out: Path) -> None:
    _, sI = parse_ibkr(inception)
    _, sF = parse_ibkr(fiscal)

    open_inc = extract_open_positions(sI)
    open_fy = extract_open_positions(sF)
    orders = merge_orders(
        extract_orders(sI),
        extract_orders(sF),
        transfers_to_buy_orders(
            extract_transfers(sI) + extract_transfers(sF),
            [open_inc, open_fy],
        ),
        [goog_seed_order()],
    )
    for o in orders:
        o["symbol"] = normalize_symbol(o["symbol"])

    splits = merge_splits(extract_stock_splits(sI), extract_stock_splits(sF))
    lots_ye = build_lots_as_of(orders, splits, date(2025, 12, 31))

    close_px = {}
    for sym, info in open_inc.items():
        if info.get("close_price") is not None:
            close_px[normalize_symbol(sym)] = (info["close_price"], info.get("currency") or "USD")

    rows: list[list] = []

    def add(sec, kind, *vals):
        rows.append([sec, kind, *vals])

    add("Statement", "Header", "Field Name", "Field Value")
    add("Statement", "Data", "Title", "Activity Statement")
    add("Statement", "Data", "Period", "January 1, 2025 - December 31, 2025")
    add(
        "Statement",
        "Data",
        "WhenGenerated",
        "Synthesized for Schedule FA from FY statements + GOOG opening seed; YE marks proxied",
    )
    add("Account Information", "Header", "Field Name", "Field Value")
    for _, r in sI["Account Information"]:
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
    for sym, lots in sorted(lots_ye.items()):
        qty = sum(L.qty for L in lots)
        if qty <= 1e-8:
            continue
        cost = sum(L.cost_local for L in lots)
        ccy = lots[0].currency or "USD"
        px = (cost / qty) if qty else 0.0
        if sym in close_px and close_px[sym][1] == ccy:
            px = close_px[sym][0]
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
    for src in (sI, sF):
        for kind, r in src.get("Transfers", []):
            if kind != "Data" or r[0] == "Total":
                continue
            add("Transfers", "Data", *r)

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
    for src in (sI, sF):
        for kind, r in src.get("Trades", []):
            if kind != "Data" or r[0] != "Order" or r[1] != "Stocks":
                continue
            d = datetime.strptime(r[4][:10], "%Y-%m-%d").date()
            if _in_cy2025(d):
                r2 = list(r)
                r2[3] = normalize_symbol(r2[3])
                add("Trades", "Data", *r2)

    add(
        "Corporate Actions",
        "Header",
        "Asset Category",
        "Currency",
        "Report Date",
        "Date/Time",
        "Description",
        "Quantity",
        "Proceeds",
        "Value",
        "Realized P/L",
        "Code",
    )
    for src in (sI, sF):
        for kind, r in src.get("Corporate Actions", []):
            if kind != "Data" or not r or r[0] in ("Total", "Total in USD"):
                continue
            d = datetime.strptime(r[2][:10], "%Y-%m-%d").date()
            if _in_cy2025(d):
                add("Corporate Actions", "Data", *r)

    add("Dividends", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted(
        [x for x in extract_dividends(sI) + extract_dividends(sF) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Dividends", "Data", d["currency"], d["date"].isoformat(), d["description"], d["amount"])

    add("Withholding Tax", "Header", "Currency", "Date", "Description", "Amount", "Code")
    for w in sorted(
        [x for x in extract_withholding(sI) + extract_withholding(sF) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Withholding Tax", "Data", w["currency"], w["date"].isoformat(), w["description"], w["amount"], "")

    add("Deposits & Withdrawals", "Header", "Currency", "Settle Date", "Description", "Amount")
    for src in (sI, sF):
        for kind, r in src.get("Deposits & Withdrawals", []):
            if kind != "Data" or not r or r[0] in ("Total",) or str(r[0]).startswith("Total"):
                continue
            if len(r) < 4 or not r[1]:
                continue
            d = datetime.strptime(r[1][:10], "%Y-%m-%d").date()
            if _in_cy2025(d):
                add("Deposits & Withdrawals", "Data", *r)

    add("Interest", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted(
        [x for x in extract_interest(sI) + extract_interest(sF) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Interest", "Data", d["currency"], d["date"].isoformat(), d["description"], d["amount"])

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
    seen = set()
    for src in (sI, sF):
        for kind, r in src.get("Financial Instrument Information", []):
            if kind != "Data":
                continue
            key = r[1]
            if key in seen:
                continue
            seen.add(key)
            add("Financial Instrument Information", "Data", *r)

    annual_out.parent.mkdir(parents=True, exist_ok=True)
    with open(annual_out, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    print(f"Synthesized annual → {annual_out} ({len(rows)} rows, YE symbols={len(mtm_rows)}, NAV~{ending})")


NOTES = [
    ("Assessee", "Gunjan Narulkar — IBKR U16931511 (Paid) — Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED"),
    ("Related account", "Also holds U22995548 (Free) — GOOG 179 shares transferred out on 03-Dec-2025; see separate workbook"),
    (
        "Statement coverage",
        "Earliest Paid statement starts 31-Dec-2024 with GOOG already held (prior qty 216). "
        "GOOG opening lot seeded 31-Dec-2024 cost USD 19,300.32 (backsolved from Jan-2025 sell bases + Mar-2025 open). "
        "True original buy date of GOOG pre-dates furnished statements — confirm for LTCG clock.",
    ),
    (
        "Annual CY2025",
        "Not furnished — synthesized. YE marks proxied from 31-Mar-2025 closes where still held, else cost. "
        "Prefer IBKR Annual / PortfolioAnalyst for filing.",
    ),
    ("Internal transfer", "GOOG Internal Out to U22995548 excluded from taxable CG (same beneficial owner); shown on FA as disposal/transfer"),
    ("NFLX split", "10-for-1 split applied in FIFO where present"),
]


def run():
    synthesize_annual(INCEPTION, FISCAL, ANNUAL)
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="on or before 2024-12-31 (GOOG held at Paid statement start)",
        assessee_label="Gunjan Narulkar (U16931511 Paid)",
        notes_extra=NOTES,
        seed_orders=[goog_seed_order()],
    )


if __name__ == "__main__":
    run()
