#!/usr/bin/env python3
"""
Schedule FA / Capital Gains / Dividends for Puneet Kohli (IBKR U17334752).

Inputs (FY Activity Statements only — no separate CY Annual was furnished):
  - inception: Apr 1, 2024 – Mar 31, 2025 (FOP transfer of MSFT)
  - fiscal:    Apr 1, 2025 – Mar 31, 2026
  - annual:    synthesized CY 2025 from the above + YE MSFT close 481.48 (Yahoo/StatMuse)

A3 Col C = Symbol only (same convention as Ray Stephanos latest output).
"""

from __future__ import annotations

from pathlib import Path

from prepare_schedule_fa_ray import main

ROOT = Path(__file__).resolve().parents[1]

NOTES = [
    ("Assessee", "Puneet Kohli — IBKR U17334752 — Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED"),
    (
        "Acquisition",
        "MSFT 77 shares acquired via FOP transfer In on 2025-01-29. "
        "IBKR cost basis USD 29,495.26 used (not transfer market value USD 34,434.40). "
        "Original purchase date outside IBKR not available — confirm for LTCG holding-period clock.",
    ),
    (
        "Annual CY2025 statement",
        "Not furnished. CY2025 dividends/WHT synthesized from FY statements. "
        "YE2025 MSFT close USD 481.48 (Yahoo Finance / StatMuse 31-Dec-2025) used for A3 closing; "
        "prefer replacing with IBKR Annual Activity Statement / PortfolioAnalyst when available.",
    ),
    ("Capital gains FY2025-26", "Nil — no disposals / sales in the fiscal year."),
    ("Capital gains FY2024-25", "Nil — only inbound FOP transfer; no sales."),
    ("Holdings", "Single security: MSFT (Microsoft Corp) qty 77 throughout after 29-Jan-2025."),
]


def run():
    main(
        inception=ROOT / "input_Inception_FY2024-25_Kohli.csv",
        annual=ROOT / "input_Annual_Statement_Kohli.csv",
        fiscal=ROOT / "input_Fiscal_Statement_Kohli.csv",
        out=ROOT / "output" / "Foreign_Assets_Schedule_FA_Puneet_Kohli_AY2026-27.xlsx",
        account_open_date="2025-01-29 (first inbound FOP transfer of MSFT)",
        assessee_label="Puneet Kohli",
        notes_extra=NOTES,
    )


if __name__ == "__main__":
    run()
