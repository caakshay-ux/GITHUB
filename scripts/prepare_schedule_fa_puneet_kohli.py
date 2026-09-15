#!/usr/bin/env python3
"""
Schedule FA / Capital Gains / Dividends for Puneet Kohli (IBKR U17334752).

Inputs (FY Activity Statements — no separate CY Annual was furnished):
  - inception: Apr 1, 2024 – Mar 31, 2025 (FOP transfer of MSFT)
  - fiscal:    Apr 1, 2025 – Mar 31, 2026
  - FY2026-27 YTD: Apr 1, 2026 – Sep 14, 2026
  - annual:    synthesized CY 2025 from the above + YE MSFT close 481.48 (Yahoo/StatMuse)

A3 Col C = Symbol only (same convention as Ray Stephanos latest output).
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from fy2026_27_append import append_fy2026_27  # noqa: E402
from prepare_schedule_fa_ray import main  # noqa: E402

INCEPTION = ROOT / "input_Inception_FY2024-25_Kohli.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Kohli.csv"
FISCAL_FY2627 = ROOT / "input_Fiscal_Statement_FY2026-27_Kohli.csv"
ANNUAL = ROOT / "input_Annual_Statement_Kohli.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Puneet_Kohli_AY2026-27.xlsx"
FY2627_END = date(2026, 9, 14)

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
    (
        "Capital gains FY2026-27 YTD",
        "Nil — no disposals / sales through 14-Sep-2026. Still holding MSFT × 77.",
    ),
    ("Holdings", "Single security: MSFT (Microsoft Corp) qty 77 throughout after 29-Jan-2025."),
    (
        "Dividends FY2026-27 YTD",
        "MSFT USD 0.91 × 77 on 11-Jun-2026 and 10-Sep-2026 (total USD 140.14). "
        "US WHT 25% (USD 17.52 × 2) — see FTC sheet / Schedule OS - FY2026-27.",
    ),
    ("Related account", "Also holds main trading account U16755051."),
]


def run():
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-01-29 (first inbound FOP transfer of MSFT)",
        assessee_label="Puneet Kohli",
        notes_extra=NOTES,
    )
    append_fy2026_27(
        OUT,
        statement_paths=[INCEPTION, FISCAL],
        fy_statement=FISCAL_FY2627,
        end=FY2627_END,
        extra_notes=[
            (
                "FY2026-27 coverage",
                "Activity Statement ends 14-Sep-2026 — not full FY. "
                "Obtain statement to 31-Mar-2027 for complete year.",
            ),
        ],
    )
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    run()
