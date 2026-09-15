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
FISCAL_FY2627 = ROOT / "input_Fiscal_Statement_FY2026-27_Tushar_Agrawal.csv"
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
        "FY 2025-26 Activity Statement (01-Apr-2025–31-Mar-2026) + FY 2026-27 YTD "
        "(01-Apr-2026–03-Sep-2026). 19 symbols held at FY2025-26 start — opening lots seeded "
        "31-Dec-2024 with costs backsolved from IBKR sell Basis (FIFO) / Open Positions cost. "
        "True original buy dates pre-date furnished statements — confirm for LTCG clock.",
    ),
    (
        "Annual CY2025",
        "Not furnished — synthesized from FY2025-26 trades in CY2025 + seeds. YE2025 marks at cost. "
        "CY dividends/interest omit Jan–Mar 2025. Prefer IBKR Annual / PortfolioAnalyst for filing.",
    ),
    (
        "FY 2026-27 (partial)",
        "Sheets 'Capital Gains FY2026-27' and 'Schedule OS - FY2026-27' cover 01-Apr-2026 to 03-Sep-2026 only "
        "(statement end). Full-year FY2026-27 requires Activity Statement through 31-Mar-2027. "
        "SBI TT Apr–Aug 2026 from public TTBR compilations — verify on sbi.co.in before filing.",
    ),
    (
        "Capital gains",
        "FY2025-26: BLBD/AMZN/ASML/BX/CAT/GE/LLY/NFLX/TGT/TSM/XOM/NOVd/R6C0d/ACN/CPNG/WCBR/1810/GD etc. "
        "FY2026-27 YTD: SLV/NBIS/TIP/URA/1919/4568.T/AAXJ/BNO/IEF/IWR/QQQ/VT/VTV etc.",
    ),
    (
        "Peak NAV",
        "Activity Statement / synthesized peak = max(start, end) only. Obtain PortfolioAnalyst for true peak A2.",
    ),
]


def _sales_to_cg_rows(sales):
    from prepare_schedule_fa_ray import (
        FX_TO_USD_YE2025,
        SBI_TT_EUR,
        month_end_on_or_before,
        rule115_usd,
        to_usd,
    )

    rows = []
    for s in sales:
        if (s.code or "").upper() in ("I",):
            continue
        if abs(s.qty) < 1e-6 or (abs(s.proceeds_local) < 1e-8 and abs(s.cost_local) < 1e-8):
            continue
        if s.currency == "USD":
            proc_usd = s.proceeds_local
            cost_usd = s.cost_local
        else:
            proc_usd = to_usd(s.proceeds_local, s.currency)
            cost_usd = to_usd(s.cost_local, s.currency)

        sale_rate_usd = rule115_usd(s.sell_date)
        cost_rate_usd = rule115_usd(s.acq_date)

        if s.currency == "EUR":
            sale_prev = (
                date(s.sell_date.year, s.sell_date.month, 1) - __import__("datetime").timedelta(days=1)
                if s.sell_date.month > 1
                else date(s.sell_date.year - 1, 12, 31)
            )
            cost_prev = (
                date(s.acq_date.year, s.acq_date.month, 1) - __import__("datetime").timedelta(days=1)
                if s.acq_date.month > 1
                else date(s.acq_date.year - 1, 12, 31)
            )
            sale_rate = SBI_TT_EUR[month_end_on_or_before(sale_prev, SBI_TT_EUR)]
            cost_rate = SBI_TT_EUR[month_end_on_or_before(cost_prev, SBI_TT_EUR)]
            sale_inr = s.proceeds_local * sale_rate
            cost_inr = s.cost_local * cost_rate
            comm_inr = abs(s.comm_usd) * sale_rate_usd
        else:
            sale_rate = sale_rate_usd
            cost_rate = cost_rate_usd
            sale_inr = proc_usd * sale_rate_usd
            cost_inr = cost_usd * cost_rate_usd
            comm_inr = abs(s.comm_usd) * sale_rate_usd

        rows.append(
            {
                "symbol": s.symbol,
                "currency": s.currency,
                "qty": round(s.qty, 6),
                "acq_date": s.acq_date,
                "sell_date": s.sell_date,
                "holding_days": s.holding_days,
                "type": "LTCG" if s.is_ltcg else "STCG",
                "proceeds_local": round(s.proceeds_local, 2),
                "cost_local": round(s.cost_local, 2),
                "proceeds_usd": round(proc_usd, 2),
                "cost_usd": round(cost_usd, 2),
                "comm_usd": round(abs(s.comm_usd), 2),
                "sale_rate": sale_rate,
                "cost_rate": cost_rate,
                "sale_inr": round(sale_inr),
                "comm_inr": round(comm_inr),
                "cost_inr": round(cost_inr),
                "gain_inr": round(sale_inr - comm_inr - cost_inr),
                "note": "",
                "fcy_to_usd": 1.0 if s.currency in ("USD", "EUR") else FX_TO_USD_YE2025.get(s.currency, 1.0),
            }
        )
    return rows


def _write_formula_cg(wb, title, subtitle, rows):
    from openpyxl.styles import Font, PatternFill
    from prepare_schedule_fa_ray import FX_TO_USD_YE2025, autosize, rule115_usd, style_header

    ws = wb.create_sheet(title)
    ws["A1"] = subtitle
    ws.merge_cells("A1:U1")
    ws["A2"] = (
        "RECONCILIATION SHEET (formula-driven). Edit yellow cells — P&L recalculates. "
        "Gain/(Loss) INR = Sale INR − Comm INR − Cost INR. Holding days >730 → LTCG else STCG."
    )
    ws.merge_cells("A2:U2")
    ws["A3"] = (
        "F=E-D | G=IF(F>730,\"LTCG\",\"STCG\") | L=H*K | M=I*K | "
        "Q=L*N (Sale INR) | R=J*P (Comm INR) | S=M*O (Cost INR) | T=Q-R-S (Gain)"
    )
    ws.merge_cells("A3:U3")
    headers = [
        "Symbol", "CCY", "Qty", "Acquisition Date", "Sale Date", "Holding Days", "Type",
        "Sale Proceeds (FCY)", "Cost Basis (FCY)", "Comm (USD)", "FCY→USD Rate",
        "Sale Proceeds (USD)", "Cost (USD)", "Sale FX Rate", "Cost FX Rate", "Comm USD/INR Rate",
        "Sale (INR)", "Comm (INR)", "Cost (INR)", "Gain/(Loss) INR", "Notes",
    ]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    style_header(ws, 4, len(headers))
    fill = PatternFill("solid", fgColor="FFF2CC")
    first = 5
    for i, r in enumerate(rows):
        rr = first + i
        usd_sale = rule115_usd(r["sell_date"])
        if r["currency"] == "EUR":
            fcy_to_usd = 1.0
            sale_fx, cost_fx = r["sale_rate"], r["cost_rate"]
        else:
            fcy_to_usd = 1.0 if r["currency"] == "USD" else r.get("fcy_to_usd", FX_TO_USD_YE2025.get(r["currency"], 1.0))
            sale_fx = r["sale_rate"] if r["currency"] == "USD" else rule115_usd(r["sell_date"])
            cost_fx = r["cost_rate"] if r["currency"] == "USD" else rule115_usd(r["acq_date"])
        ws.cell(rr, 1, r["symbol"])
        ws.cell(rr, 2, r["currency"])
        ws.cell(rr, 3, r["qty"])
        ws.cell(rr, 4, r["acq_date"])
        ws.cell(rr, 5, r["sell_date"])
        ws.cell(rr, 6, f"=E{rr}-D{rr}")
        ws.cell(rr, 7, f'=IF(F{rr}>730,"LTCG","STCG")')
        ws.cell(rr, 8, round(r["proceeds_local"], 6))
        ws.cell(rr, 9, round(r["cost_local"], 6))
        ws.cell(rr, 10, r["comm_usd"])
        ws.cell(rr, 11, fcy_to_usd)
        ws.cell(rr, 12, f"=H{rr}*K{rr}")
        ws.cell(rr, 13, f"=I{rr}*K{rr}")
        ws.cell(rr, 14, sale_fx)
        ws.cell(rr, 15, cost_fx)
        ws.cell(rr, 16, usd_sale)
        ws.cell(rr, 17, f"=L{rr}*N{rr}")
        ws.cell(rr, 18, f"=J{rr}*P{rr}")
        ws.cell(rr, 19, f"=M{rr}*O{rr}")
        ws.cell(rr, 20, f"=Q{rr}-R{rr}-S{rr}")
        ws.cell(rr, 21, r.get("note") or "")
        for col in (8, 9, 10, 11, 14, 15, 16):
            ws.cell(rr, col).fill = fill
    last = first + len(rows) - 1 if rows else first - 1
    tot = last + 2
    if rows:
        ws.cell(tot, 7, "TOTAL STCG (INR)")
        ws.cell(tot, 20, f'=SUMIF(G{first}:G{last},"STCG",T{first}:T{last})')
        ws.cell(tot + 1, 7, "TOTAL LTCG (INR)")
        ws.cell(tot + 1, 20, f'=SUMIF(G{first}:G{last},"LTCG",T{first}:T{last})')
        ws.cell(tot + 2, 7, "TOTAL CG (INR)")
        ws.cell(tot + 2, 20, f"=T{tot}+T{tot+1}")
        for r in range(tot, tot + 3):
            ws.cell(r, 7).font = Font(bold=True)
            ws.cell(r, 20).font = Font(bold=True)
    autosize(ws)


def _write_os_fy(wb, title, subtitle, dividends, interest):
    from prepare_schedule_fa_ray import autosize, rule115_usd, style_header, to_usd

    ws = wb.create_sheet(title)
    ws["A1"] = subtitle
    ws.append(["Date", "Description", "Currency", "Amount (FCY)", "Amount (USD)", "USD/INR (Rule 115)", "Amount (INR)"])
    style_header(ws, 2, 7)
    int_usd = int_inr = 0.0
    ws.append(["— INTEREST —"])
    for d in interest:
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        rate = rule115_usd(d["date"])
        inr = usd * rate
        int_usd += usd
        int_inr += inr
        ws.append([d["date"], d["description"], d["currency"], d["amount"], round(usd, 4), rate, round(inr)])
    ws.append([None, "TOTAL INTEREST", None, None, round(int_usd, 2), None, round(int_inr)])
    ws.append([])
    div_usd = div_inr = 0.0
    ws.append(["— DIVIDENDS —"])
    for d in dividends:
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        rate = rule115_usd(d["date"])
        inr = usd * rate
        div_usd += usd
        div_inr += inr
        ws.append([d["date"], d["description"], d["currency"], d["amount"], round(usd, 4), rate, round(inr)])
    ws.append([None, "TOTAL DIVIDENDS", None, None, round(div_usd, 2), None, round(div_inr)])
    autosize(ws)
    return div_usd, div_inr, int_usd, int_inr


def append_fy2026_27(out: Path, seeds: list[dict]) -> dict:
    """Add FY2026-27 (YTD to statement end) CG + OS sheets to the workbook."""
    from openpyxl import load_workbook
    from prepare_schedule_fa_ray import (
        build_fifo_sales,
        build_lots_as_of,
        extract_dividends,
        extract_interest,
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
        rule115_usd,
    )

    _, sF = parse_ibkr(FISCAL)
    _, sN = parse_ibkr(FISCAL_FY2627)
    open_fy = extract_open_positions_safe(sF)
    open_n = extract_open_positions_safe(sN)

    orders = merge_orders(
        extract_orders(sF),
        extract_orders(sN),
        transfers_to_buy_orders(extract_transfers(sF) + extract_transfers(sN), [open_fy, open_n]),
        seeds,
    )
    for o in orders:
        o["symbol"] = normalize_symbol(o["symbol"])
    splits = merge_splits(extract_stock_splits(sF), extract_stock_splits(sN))

    lots_fy2627_open = build_lots_as_of(orders, splits, date(2026, 3, 31))
    # Only in-window orders — opening lots already include pre-FY seeds (avoid double-apply)
    orders_fy = [o for o in orders if o["date"] >= date(2026, 4, 1)]
    sales, _ = build_fifo_sales(
        lots_fy2627_open,
        orders_fy,
        date(2026, 4, 1),
        date(2026, 9, 3),
        splits=splits,
    )
    cg_rows = _sales_to_cg_rows(sales)
    stcg = sum(r["gain_inr"] for r in cg_rows if r["type"] == "STCG")
    ltcg = sum(r["gain_inr"] for r in cg_rows if r["type"] == "LTCG")

    divs = [
        d
        for d in extract_dividends(sN)
        if date(2026, 4, 1) <= d["date"] <= date(2026, 9, 3)
    ]
    ints = [
        d
        for d in extract_interest(sN)
        if date(2026, 4, 1) <= d["date"] <= date(2026, 9, 3)
    ]
    wht = [
        w
        for w in extract_withholding(sN)
        if date(2026, 4, 1) <= w["date"] <= date(2026, 9, 3)
    ]

    wb = load_workbook(out)
    if "Capital Gains FY2026-27" in wb.sheetnames:
        del wb["Capital Gains FY2026-27"]
    if "Schedule OS - FY2026-27" in wb.sheetnames:
        del wb["Schedule OS - FY2026-27"]

    _write_formula_cg(
        wb,
        "Capital Gains FY2026-27",
        "Capital gains — FY 2026-27 YTD (01-Apr-2026 to 03-Sep-2026). Partial year — statement ends 03-Sep-2026. FIFO. Yellow cells editable.",
        cg_rows,
    )
    div_usd, div_inr, int_usd, int_inr = _write_os_fy(
        wb,
        "Schedule OS - FY2026-27",
        "Interest & Dividend income — FY 2026-27 YTD (01-Apr-2026 to 03-Sep-2026). Partial — full year needs statement through 31-Mar-2027.",
        divs,
        ints,
    )

    # Append FY26-27 withholding to FTC sheet
    if "FTC - Withholding Tax" in wb.sheetnames:
        wsw = wb["FTC - Withholding Tax"]
        for w in wht:
            usd = w["amount"] if w["currency"] == "USD" else to_usd(w["amount"], w["currency"])
            rate = rule115_usd(w["date"])
            wsw.append(
                [
                    "FY2026-27 YTD",
                    w["date"],
                    w["description"],
                    w["currency"],
                    w["amount"],
                    round(usd, 4),
                    rate,
                    round(usd * rate),
                ]
            )

    # Notes
    if "Notes for CA" in wb.sheetnames:
        wn = wb["Notes for CA"]
        wn.append([])
        wn.append(["FY 2026-27 YTD (01-Apr-2026 to 03-Sep-2026) — partial statement"])
        wn.append(["Dividends FY2026-27 YTD", f"USD {div_usd:,.2f} / INR {div_inr:,.0f}"])
        wn.append(["Interest FY2026-27 YTD", f"USD {int_usd:,.2f} / INR {int_inr:,.0f}"])
        wn.append(["STCG FY2026-27 YTD (INR)", f"{stcg:,.0f}"])
        wn.append(["LTCG FY2026-27 YTD (INR)", f"{ltcg:,.0f}"])
        wn.append(
            [
                "FY2026-27 coverage",
                "Activity Statement ends 03-Sep-2026 — not full FY. Obtain statement to 31-Mar-2027 for complete year.",
            ]
        )

    # CG summary by symbol for FY26-27
    if "CG Summary by Symbol" in wb.sheetnames:
        wss = wb["CG Summary by Symbol"]
        wss.append([])
        wss.append(["— FY 2026-27 YTD (to 03-Sep-2026) —"])
        wss.append(["Symbol", "STCG INR", "LTCG INR", "Total Gain/(Loss) INR", "Sale Proceeds USD", "Cost USD"])
        by = {}
        for r in cg_rows:
            e = by.setdefault(r["symbol"], {"stcg": 0, "ltcg": 0, "sale": 0.0, "cost": 0.0})
            if r["type"] == "STCG":
                e["stcg"] += r["gain_inr"]
            else:
                e["ltcg"] += r["gain_inr"]
            e["sale"] += r["proceeds_usd"]
            e["cost"] += r["cost_usd"]
        for sym in sorted(by):
            e = by[sym]
            wss.append([sym, e["stcg"], e["ltcg"], e["stcg"] + e["ltcg"], round(e["sale"], 2), round(e["cost"], 2)])

    wb.save(out)
    summary = {
        "stcg": stcg,
        "ltcg": ltcg,
        "div_usd": div_usd,
        "div_inr": div_inr,
        "int_usd": int_usd,
        "int_inr": int_inr,
        "cg_lots": len(cg_rows),
        "div_lines": len(divs),
    }
    print(
        f"FY2026-27 YTD: CG lots={len(cg_rows)} STCG={stcg:,.0f} LTCG={ltcg:,.0f} "
        f"Div USD={div_usd:,.2f} INR={div_inr:,.0f}"
    )
    return summary


def extract_open_positions_safe(sections):
    from prepare_schedule_fa_ray import extract_open_positions

    try:
        return extract_open_positions(sections)
    except Exception:
        return {}


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
    append_fy2026_27(OUT, seeds)


if __name__ == "__main__":
    run()
