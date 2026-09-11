#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Vinodkrishna Poyyale — IBKR U15388294.

Inputs: FY Activity Statements only (10-Sep-2024–31-Mar-2025; 01-Apr-2025–31-Mar-2026).
CY2025 Annual is synthesized. YE2025 marks proxied from 31-Mar-2025 closes where still held.
A3 Col C = Symbol only (same as Ray Stephanos latest).
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
    fin_info,
    merge_orders,
    merge_splits,
    normalize_symbol,
    parse_ibkr,
    to_usd,
    transfers_to_buy_orders,
    main,
)

INCEPTION = ROOT / "input_Inception_FY2025-26_Abhisek_Banerjee.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Abhisek_Banerjee.csv"
ANNUAL = ROOT / "input_Annual_Statement_Abhisek_Banerjee.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Abhisek_Banerjee_AY2026-27.xlsx"


def _in_cy2025(d: date) -> bool:
    return date(2025, 1, 1) <= d <= date(2025, 12, 31)


def synthesize_annual(src: Path, annual_out: Path) -> None:
    """Build CY2025 annual from a single statement (avoid double-count when I==F)."""
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

    # No true 31-Dec-2025 marks (statement ends 31-Mar-2026) — YE at cost.
    close_px: dict = {}

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
        "Synthesized for Schedule FA from FY statement (opened May-2025); YE marks at cost (see Notes)",
    )

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
    for kind, r in s.get("Transfers", []):
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
    for kind, r in s.get("Trades", []):
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
    for kind, r in s.get("Corporate Actions", []):
        if kind != "Data" or not r or r[0] in ("Total", "Total in USD"):
            continue
        d = datetime.strptime(r[2][:10], "%Y-%m-%d").date()
        if _in_cy2025(d):
            add("Corporate Actions", "Data", *r)

    add("Dividends", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted(
        [x for x in extract_dividends(s) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Dividends", "Data", d["currency"], d["date"].isoformat(), d["description"], d["amount"])

    add("Withholding Tax", "Header", "Currency", "Date", "Description", "Amount", "Code")
    for w in sorted(
        [x for x in extract_withholding(s) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Withholding Tax", "Data", w["currency"], w["date"].isoformat(), w["description"], w["amount"], "")

    add("Deposits & Withdrawals", "Header", "Currency", "Settle Date", "Description", "Amount")
    for kind, r in s.get("Deposits & Withdrawals", []):
        if kind != "Data" or r[0] == "Total" or str(r[0]).startswith("Total"):
            continue
        d = datetime.strptime(r[1][:10], "%Y-%m-%d").date()
        if _in_cy2025(d):
            add("Deposits & Withdrawals", "Data", *r)

    add("Interest", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted(
        [x for x in extract_interest(s) if _in_cy2025(x["date"])],
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
    for kind, r in s.get("Financial Instrument Information", []):
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
    ("Assessee", "Abhisek Banerjee — IBKR U20221090 — Advisor: Financial Hospital Advisory LLP"),
    (
        "Account opening",
        "First deposits 16-May-2025 (USD 500) and 19-May-2025 (USD 11,300). "
        "Statement period 16-May-2025 – 31-Mar-2026; prior NAV 0 — treated as since inception.",
    ),
    (
        "Annual CY2025 statement",
        "Not furnished. CY2025 synthesized from the FY statement (single parse — no double-count). "
        "YE2025 closing prices proxied at cost (all lots acquired in CY2025 and still held) — "
        "prefer IBKR Annual / PortfolioAnalyst for filing-quality peak & YE values.",
    ),
    ("Capital gains", "Nil realized stock disposals in FY2025-26 (buys / holdings only; forex P/L ignored for Schedule FA equity)."),
    ("Aliases", "AIR↔AIRd, SHELL↔R6C0d normalized for FIFO; MC = LVMH (Euronext Paris)"),
    ("Multi-ccy", "EUR / JPY converted via IBKR YE FX then Rule 115 USD/INR (editable on CG)"),
    ("Holdings", "AIRd, MC, R6C0d, 4568.T, BAC, CAT, COPX, DB, GLD, GOOG, TGT, URA"),
]


def run():
    synthesize_annual(FISCAL, ANNUAL)
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-05-16 (first USD EFT deposit)",
        assessee_label="Abhisek Banerjee",
        notes_extra=NOTES,
    )


if __name__ == "__main__":
    run()
