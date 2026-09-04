#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Tushar Agrawal — IBKR U16595399 (Individual).

Inputs:
  - Fiscal FY 2025-26 only (01-Apr-2025 to 31-Mar-2026) — real IBKR Activity Statement
  - Inception stub (account metadata) — no pre-FY Activity Statement furnished
  - Annual CY 2025 — synthesized; YE marks at cost

Opening lots at 01-Apr-2025 (MTM Prior Quantity) are seeded 31-Dec-2024 with
costs backsolved from IBKR sell Basis (FIFO) or carried Open Positions cost.
A3 Col C = Symbol only (same as Ray Stephanos).
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

INCEPTION = ROOT / "input_Inception_FY2024-25_Tushar_Agrawal.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Tushar_Agrawal.csv"
ANNUAL = ROOT / "input_Annual_Statement_Tushar_Agrawal.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Tushar_Agrawal_AY2026-27.xlsx"

# Backsolved from IBKR sell Basis (FIFO) / end Open Positions cost for still-held.
# Qty = MTM Prior Quantity at 01-Apr-2025.
SEED_LOTS: list[tuple[str, str, float, float]] = [
    # symbol, currency, qty, total_cost_local
    ("4568.T", "JPY", 100.0, 372297.6),
    ("ACN", "USD", 3.3, 1181.7488382),
    ("AMZN", "USD", 3.3, 759.5842138),
    ("ASML", "USD", 2.0, 1507.750144),
    ("BLBD", "USD", 75.0, 2926.0036),
    ("BX", "USD", 6.0, 1127.689432),
    ("CAT", "USD", 3.0, 1168.000144),
    ("GD", "USD", 3.0, 806.018638),
    ("GE", "USD", 9.0, 1529.200432),
    ("LLY", "USD", 2.0, 1581.000096),
    ("NFLX", "USD", 1.6, 1488.9640906),
    ("NOVd", "EUR", 15.0, 1053.0),
    ("R6C0d", "EUR", 33.0, 1013.625),
    ("SLV", "USD", 159.0, 4524.598928),
    ("TGT", "USD", 12.0, 1633.000576),
    ("TIP", "USD", 57.0, 6185.502736),
    ("TSLA", "USD", 6.0, 2584.760354),
    ("TSM", "USD", 3.0, 580.300216),
    ("XOM", "USD", 24.0, 2679.401152),
]


def seed_orders() -> list[dict]:
    out = []
    for i, (sym, ccy, qty, cost) in enumerate(SEED_LOTS):
        out.append(
            {
                "currency": ccy,
                "symbol": normalize_symbol(sym),
                "datetime": f"2024-12-31, 00:00:{i:02d}",
                "date": date(2024, 12, 31),
                "qty": qty,
                "price": cost / qty if qty else 0.0,
                "proceeds": -cost,
                "comm_usd": 0.0,
                "basis": cost,
                "realized": 0.0,
                "code": "O",
            }
        )
    return out


def _in_cy2025(d: date) -> bool:
    return date(2025, 1, 1) <= d <= date(2025, 12, 31)


def synthesize_annual(fiscal: Path, annual_out: Path, seeds: list[dict]) -> None:
    _, sF = parse_ibkr(fiscal)
    open_fy = extract_open_positions(sF)
    orders = merge_orders(
        extract_orders(sF),
        transfers_to_buy_orders(extract_transfers(sF), [open_fy]),
        seeds,
    )
    for o in orders:
        o["symbol"] = normalize_symbol(o["symbol"])

    splits = merge_splits(extract_stock_splits(sF))
    lots_ye = build_lots_as_of(orders, splits, date(2025, 12, 31))

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
        "Synthesized for Schedule FA from FY2025-26 + seeded opening lots; YE marks at cost",
    )

    add("Account Information", "Header", "Field Name", "Field Value")
    for _, r in sF["Account Information"]:
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

    # Proxy starting NAV ≈ FY start NAV from fiscal Change in NAV
    start_nav = 0.0
    for kind, r in sF.get("Change in NAV", []):
        if kind == "Data" and r and r[0] == "Starting Value":
            try:
                start_nav = float(str(r[1]).replace(",", ""))
            except (TypeError, ValueError):
                pass
    ending = round(stock_usd, 2)

    add("Net Asset Value", "Header", "Asset Class", "Prior Total", "Current Long", "Current Short", "Current Total", "Change")
    add("Net Asset Value", "Data", "Stock", start_nav, ending, 0, ending, ending - start_nav)
    add("Net Asset Value", "Data", "Total", start_nav, ending, 0, ending, ending - start_nav)
    add("Change in NAV", "Header", "Field Name", "Field Value")
    add("Change in NAV", "Data", "Starting Value", start_nav)
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
    for o in sorted(extract_orders(sF), key=lambda x: (x["date"], x["datetime"])):
        if _in_cy2025(o["date"]):
            add(
                "Trades",
                "Data",
                "Order",
                "Stocks",
                o["currency"],
                o["symbol"],
                o["datetime"],
                o["qty"],
                o["price"],
                "",
                o["proceeds"],
                o["comm_usd"],
                o["basis"],
                o["realized"],
                "",
                o.get("code") or "",
            )

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
    for kind, r in sF.get("Corporate Actions", []):
        if kind != "Data" or not r or r[0] in ("Total", "Total in USD"):
            continue
        d = datetime.strptime(r[2][:10], "%Y-%m-%d").date()
        if _in_cy2025(d):
            add("Corporate Actions", "Data", *r)

    add("Dividends", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted(
        [x for x in extract_dividends(sF) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Dividends", "Data", d["currency"], d["date"].isoformat(), d["description"], d["amount"])

    add("Withholding Tax", "Header", "Currency", "Date", "Description", "Amount", "Code")
    for w in sorted(
        [x for x in extract_withholding(sF) if _in_cy2025(x["date"])],
        key=lambda x: x["date"],
    ):
        add("Withholding Tax", "Data", w["currency"], w["date"].isoformat(), w["description"], w["amount"], "")

    add("Interest", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted(
        [x for x in extract_interest(sF) if _in_cy2025(x["date"])],
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
    for kind, r in sF.get("Financial Instrument Information", []):
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
    (
        "Assessee",
        "Tushar Agrawal — IBKR U16595399 (Individual) — Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED "
        "(also spelled Agarwal in some correspondence)",
    ),
    (
        "Statement coverage",
        "Only FY 2025-26 Activity Statement furnished (01-Apr-2025–31-Mar-2026). "
        "19 symbols already held at FY start (MTM Prior Qty) — opening lots seeded 31-Dec-2024 with "
        "costs backsolved from IBKR sell Basis (FIFO) or end Open Positions cost (SLV/TIP/TSLA/4568.T). "
        "True original buy dates pre-date furnished statements — confirm for LTCG clock.",
    ),
    (
        "Annual CY2025",
        "Not furnished — synthesized from FY trades in CY2025 + seeds. YE2025 marks at cost. "
        "CY dividends/interest omit Jan–Mar 2025 (outside fiscal window). Prefer IBKR Annual / PortfolioAnalyst for filing.",
    ),
    (
        "Capital gains",
        "FY2025-26: multiple equity/ETF disposals (BLBD, AMZN, ASML, BX, CAT, GE, LLY, NFLX, TGT, TSM, XOM, "
        "NOVd, R6C0d/SHELL, ACN, CPNG, WCBR, 1810, GD, etc.). Multi-currency USD/EUR/HKD/JPY/DKK.",
    ),
    (
        "Peak NAV",
        "Activity Statement / synthesized peak = max(start, end) only. Obtain PortfolioAnalyst for true peak A2.",
    ),
]


def run():
    seeds = seed_orders()
    synthesize_annual(FISCAL, ANNUAL, seeds)
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="on or before 2024-12-31 (positions held at FY2025-26 start)",
        assessee_label="Tushar Agrawal",
        notes_extra=NOTES,
        seed_orders=seeds,
    )


if __name__ == "__main__":
    run()
