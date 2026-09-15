#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Vineet Dev — IBKR U18045590.

Inputs (real IBKR Activity Statements):
  - Inception / FY 2024-25 (01-Apr-2024 to 31-Mar-2025) — CRM FOP In 24-Feb-2025
  - Annual CY 2025 (01-Jan-2025 to 31-Dec-2025) — Schedule FA
  - Fiscal FY 2025-26 (01-Apr-2025 to 31-Mar-2026) — ITR income & capital gains

A3 Col C = Symbol only (same as Ray Stephanos).
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from prepare_schedule_fa_ray import main  # noqa: E402

INCEPTION = ROOT / "input_Inception_FY2024-25_Vineet_Dev.csv"
ANNUAL = ROOT / "input_Annual_Statement_Vineet_Dev.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Vineet_Dev.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Vineet_Dev_AY2026-27.xlsx"

NOTES = [
    ("Assessee", "Vineet Dev — IBKR U18045590 (Individual) — Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED"),
    (
        "Account opening",
        "Opened via FOP In of CRM × 419 on 24-Feb-2025 (MV USD 129,806.20). "
        "No cash deposits in furnished statements. Cost basis from IBKR Open Positions at 31-Mar-2025 "
        "(unit cost applied to full 419). Confirm original broker acquisition dates for CRM LTCG "
        "(IBKR sale of 1 CRM on 05-Mar-2025 coded Long-term).",
    ),
    (
        "Capital gains",
        "FY2024-25: CRM × 1 sold 05-Mar-2025. FY2025-26: multiple equity/ETF disposals "
        "(incl. CRM, AMZN, ASML, BTI, NOVd/NOVN, ROG, R6C0d/SHELL, SIEd, NFLX, TSM, etc.) "
        "plus CHF-traded Swiss names (NOVN, ROG) converted via IBKR YE CHF→USD.",
    ),
    (
        "FX",
        "CHF→USD 1.2615 and EUR/JPY from IBKR YE2025 Forex Balances — editable on FX Lookup. "
        "Multi-currency: USD, EUR, CHF, JPY.",
    ),
    (
        "Peak NAV",
        "Activity Statement peak = max(start, end) only. Obtain PortfolioAnalyst for true peak A2.",
    ),
]


def run():
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-02-24 (CRM FOP In)",
        assessee_label="Vineet Dev",
        notes_extra=NOTES,
    )


if __name__ == "__main__":
    run()
