#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Arulselvam Chandrasekaran & Dhanalakshmi S —
IBKR U22929455 (Joint).

Inputs: real IBKR Annual CY2025 + Fiscal FY2025-26 (opened Nov-2025; inception = annual).
Base currency SGD; A2 NAV converted to USD via YE USD.SGD.
Same-day Internal GOOG × 219 In/Out wash with U22748155 stripped (net zero).
Nil taxable stock CG (buys only through 31-Mar-2026).
A3 Col C = Symbol only.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from prepare_schedule_fa_ray import main  # noqa: E402

INCEPTION = ROOT / "input_Inception_FY2025-26_Arulselvam_Joint_U22929455.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Arulselvam_Joint_U22929455.csv"
ANNUAL = ROOT / "input_Annual_Statement_Arulselvam_Joint_U22929455.csv"
OUT = (
    ROOT
    / "output"
    / "Foreign_Assets_Schedule_FA_Arulselvam_Dhanalakshmi_Joint_U22929455_AY2026-27.xlsx"
)

NOTES = [
    (
        "Assessee",
        "Arulselvam Chandrasekaran and Dhanalakshmi S — IBKR U22929455 (Joint) — "
        "Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED",
    ),
    (
        "Account opening",
        "Opened Nov-2025. Internal Transfer In USD 60,000 from Individual U22748155 on 12-Nov-2025. "
        "Base currency SGD. Same-day Internal GOOG × 219 In/Out wash with U22748155 on 10-Nov-2025 "
        "stripped from working CSVs (net zero equity).",
    ),
    (
        "Related Individual account",
        "See Individual workbook U22748155 (GOOG FOP + full sale; cash transfer to this Joint).",
    ),
    (
        "Capital gains",
        "Nil taxable stock sales in FY2025-26 — only buys (COPX, GLD, URA in CY2025; "
        "additional COPX/GLD/PPLT/SLV/GSL in Jan–Feb 2026).",
    ),
    (
        "Holdings at 31-Dec-2025",
        "COPX × 36, GLD × 12, URA × 60 (YE marks from IBKR Annual).",
    ),
    (
        "Holdings at 31-Mar-2026",
        "COPX × 75, GLD × 27, GSL × 54, PPLT × 30, SLV × 42, URA × 60. "
        "Ending NAV ~SGD 76,244 (~USD 59,288 at YE USD.SGD 1.286).",
    ),
    (
        "Dividends / Interest",
        "CY2025: nil dividends; interest USD 53.60 (Nov). "
        "FY2025-26 dividends USD 219.02 (COPX/URA/GSL); interest USD 224 + small SGD debit. "
        "US withholding on dividends / interest — see FTC sheet (includes later cancellations).",
    ),
    (
        "SGD base / A2 NAV",
        "Account base SGD. A2 USD NAV = SGD Ending Value ÷ 1.286 (IBKR YE USD.SGD). "
        "Prefer PortfolioAnalyst USD NAV for filing.",
    ),
]


if __name__ == "__main__":
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-11-12 (Internal Transfer In USD 60,000 from U22748155)",
        notes_extra=NOTES,
        assessee_label="Arulselvam Chandrasekaran & Dhanalakshmi S (Joint U22929455)",
    )
    print(f"Wrote {OUT}")
