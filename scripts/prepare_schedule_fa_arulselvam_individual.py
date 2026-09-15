#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Arulselvam Chandrasekaran — IBKR U22748155 (Individual).

Inputs: real IBKR Annual CY2025 + Fiscal FY2025-26 (opened Nov-2025; inception = annual).
Base currency SGD; A2 NAV converted to USD via YE USD.SGD.
GOOG FOP In cost set to IBKR sell Basis total (USD 52,506.73); same-day Internal
GOOG Out/In wash to Joint stripped from working CSVs (net zero; see Notes).
A3 Col C = Symbol only.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from prepare_schedule_fa_ray import main  # noqa: E402

INCEPTION = ROOT / "input_Inception_FY2025-26_Arulselvam_U22748155.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Arulselvam_U22748155.csv"
ANNUAL = ROOT / "input_Annual_Statement_Arulselvam_U22748155.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Arulselvam_Chandrasekaran_U22748155_AY2026-27.xlsx"

NOTES = [
    (
        "Assessee",
        "Arulselvam Chandrasekaran — IBKR U22748155 (Individual) — "
        "Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED",
    ),
    (
        "Account opening",
        "Funded via FOP In of GOOG × 447 on 05-Nov-2025 (external FOP). "
        "Base currency SGD. Cash Internal Transfer Out USD 60,000 to Joint U22929455 on 12-Nov-2025; "
        "disbursement USD 69,168.80 on 01-Dec-2025; EFT in USD 758.44 on 08-Dec-2025.",
    ),
    (
        "Related Joint account",
        "Cash USD 60,000 + same-day Internal GOOG × 219 wash Out/In with U22929455 (Dhanalakshmi S Joint) "
        "on 10-Nov-2025 — wash lines removed from working CSVs (net zero equity). See Joint workbook.",
    ),
    (
        "Capital gains",
        "Sold entire GOOG × 447 on 10-Nov-2025 in two tickets (228 @ 289.14; 219 @ 288.80). "
        "IBKR Realized ~USD 76,662 (codes C;HC — Highest Cost tax lot election). "
        "Workbook FIFO uses FOP date 05-Nov-2025 as acquisition → all STCG under India clock unless "
        "original buy dates at source are confirmed. IBKR tagged mostly Long Term in base-currency R&U "
        "(ST ~SGD 1,714 / LT ~SGD 98,161) — confirm original GOOG acquisition dates and edit yellow "
        "Acquisition Date on CG before filing. Cost basis seeded to sell Basis total USD 52,506.73 "
        "(not FOP Market Value USD 124,292.82).",
    ),
    (
        "Holdings at 31-Dec-2025 / 31-Mar-2026",
        "No open stock positions. Ending cash ~USD 831.49 (SGD NAV ~1,069.30).",
    ),
    (
        "Dividends / Interest",
        "Nil dividends. CY/FY interest: USD Credit Interest Nov-2025 USD 73.05.",
    ),
    (
        "SGD base / A2 NAV",
        "Account base SGD. A2 USD NAV = SGD Ending Value ÷ 1.286 (IBKR YE USD.SGD) ≈ USD 831.49. "
        "Peak shown as max(start, end) understates mid-year peak while GOOG was held "
        "(FOP MV ~USD 124k) — obtain PortfolioAnalyst for true peak. Prefer PA USD NAV for filing.",
    ),
]


if __name__ == "__main__":
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-11-05 (FOP In GOOG × 447)",
        notes_extra=NOTES,
        assessee_label="Arulselvam Chandrasekaran (U22748155)",
    )
    print(f"Wrote {OUT}")
