#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Yagyank Chadha — IBKR U20291582 (Individual).

Inputs: real IBKR Activity Statement FY2025-26 (23-May-2025–31-Mar-2026).
CY2025 Annual synthesized; YE marks from PortfolioAnalyst CY2025 Open Position Summary.
Inbound FOP/Internal only — nil taxable CG.
Linked: Internal In of SHOP + Lux UCITS from U15172057 on 23-May-2025.
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
    transfers_to_out_orders,
    main,
)

INCEPTION = ROOT / "input_Inception_FY2025-26_Yagyank_U20291582.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Yagyank_U20291582.csv"
ANNUAL = ROOT / "input_Annual_Statement_Yagyank_U20291582.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Yagyank_Chadha_U20291582_AY2026-27.xlsx"

# YE2025 closes from PortfolioAnalyst CY2025 (same account)
PA_YE_CLOSE = {
    "LU0099574567": (80.78, "EUR"),
    "LU0106831901": (77.41, "USD"),
    "GOOG": (313.8, "USD"),
    "SHOP": (160.97, "USD"),
}


def _in_cy2025(d: date) -> bool:
    return date(2025, 1, 1) <= d <= date(2025, 12, 31)


def synthesize_annual(src: Path, annual_out: Path) -> None:
    _, s = parse_ibkr(src)
    open_pos = extract_open_positions(s)
    orders = merge_orders(
        extract_orders(s),
        transfers_to_buy_orders(extract_transfers(s), [open_pos]),
        transfers_to_out_orders(extract_transfers(s)),
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
    add(
        "Statement",
        "Data",
        "WhenGenerated",
        "Synthesized for Schedule FA from FY Activity Statement; YE marks from PortfolioAnalyst CY2025",
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
        if sym in PA_YE_CLOSE and PA_YE_CLOSE[sym][1] == ccy:
            px = PA_YE_CLOSE[sym][0]
        value = px * qty
        stock_usd += to_usd(value, ccy)
        add(
            "Open Positions",
            "Data",
            "Summary",
            "Stocks" if not sym.startswith("LU") else "Mutual Funds",
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

    # Cash from FY ending cash is after Mar — use PA YE cash ~16.54
    cash = 16.54
    ending = round(stock_usd + cash, 2)

    add("Net Asset Value", "Header", "Asset Class", "Prior Total", "Current Long", "Current Short", "Current Total", "Change")
    add("Net Asset Value", "Data", "Cash ", 0, cash, 0, cash, cash)
    add("Net Asset Value", "Data", "Stock", 0, round(stock_usd, 2), 0, round(stock_usd, 2), round(stock_usd, 2))
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
        if kind != "Data" or r[0] == "Total" or str(r[0]).startswith("Total"):
            continue
        d = datetime.strptime(str(r[3])[:10], "%Y-%m-%d").date()
        if _in_cy2025(d):
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

    add("Dividends", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted([x for x in extract_dividends(s) if _in_cy2025(x["date"])], key=lambda x: x["date"]):
        add("Dividends", "Data", d["currency"], d["date"].isoformat(), d["description"], d["amount"])

    add("Withholding Tax", "Header", "Currency", "Date", "Description", "Amount", "Code")
    for w in sorted([x for x in extract_withholding(s) if _in_cy2025(x["date"])], key=lambda x: x["date"]):
        add("Withholding Tax", "Data", w["currency"], w["date"].isoformat(), w["description"], w["amount"], "")

    add("Deposits & Withdrawals", "Header", "Currency", "Settle Date", "Description", "Amount")
    add("Interest", "Header", "Currency", "Date", "Description", "Amount")
    for d in sorted([x for x in extract_interest(s) if _in_cy2025(x["date"])], key=lambda x: x["date"]):
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
    (
        "Assessee",
        "Yagyank Chadha — IBKR U20291582 (Individual) — "
        "Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED",
    ),
    (
        "Source data",
        "Real IBKR Activity Statement FY2025-26 (23-May-2025–31-Mar-2026). "
        "CY2025 Annual synthesized; YE marks from PortfolioAnalyst CY2025 Open Position Summary.",
    ),
    (
        "Account opening",
        "Opened 23-May-2025 via Internal In from U15172057: SHOP × 84, "
        "Fidelity Global Technology LU0099574567 × 511.42, BGF World Financials LU0106831901 × 792.45. "
        "FOP/ACATS In GOOG × 105 on 01-Dec-2025.",
    ),
    (
        "Related account",
        "Assets transferred in from main trading account U15172057 (same beneficial owner). "
        "See U15172057 workbook for Internal Out (code I — not taxable CG).",
    ),
    (
        "Capital gains",
        "Nil taxable stock/fund sales in FY2025-26 (inbound transfers only).",
    ),
    (
        "Dividends / FTC",
        "GOOG USD 22.05 on 15-Dec-2025 and 16-Mar-2026 (FY total USD 44.10). "
        "US withholding 25% (USD 5.51 × 2) — see FTC sheet.",
    ),
    (
        "Cost / acquisition dates",
        "Cost basis from IBKR Open Positions carryover. Acquisition dates = inbound transfer dates. "
        "Confirm original buy dates at U15172057 / prior broker for LTCG on future sales "
        "(SHOP/funds originally FOP into U15172057 in Sep-2024).",
    ),
]


def run():
    synthesize_annual(FISCAL, ANNUAL)
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-05-23 (Internal In from U15172057)",
        assessee_label="Yagyank Chadha (U20291582)",
        notes_extra=NOTES,
    )
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    run()
