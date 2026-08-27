#!/usr/bin/env python3
"""
Schedule FA / CG / Dividends for Yagyank Chadha — IBKR U20291582 (Individual).

Source: IBKR PortfolioAnalyst reports (not full Activity Statements) for
CY2025 (23-May-2025–31-Dec-2025) and FY2025-26 (23-May-2025–31-Mar-2026),
advisor MATCAP WEALTH ADVISORS PRIVATE LIMITED.

Working CSVs under input_*_Yagyank_Chadha.csv are synthesized into the
standard IBKR Activity Statement layout so the shared Ray generator can run.
Cost basis / YE marks taken from PortfolioAnalyst Open Position Summary.
Inbound FOP/ACATS only — no stock sales → nil taxable CG.
A3 Col C = Symbol only (ISIN for Lux UCITS funds).
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from prepare_schedule_fa_ray import main  # noqa: E402

INCEPTION = ROOT / "input_Inception_FY2025-26_Yagyank_Chadha.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Yagyank_Chadha.csv"
ANNUAL = ROOT / "input_Annual_Statement_Yagyank_Chadha.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Yagyank_Chadha_U20291582_AY2026-27.xlsx"

NOTES = [
    (
        "Assessee",
        "Yagyank Chadha — IBKR U20291582 (Individual) — "
        "Advisor: MATCAP WEALTH ADVISORS PRIVATE LIMITED",
    ),
    (
        "Source data",
        "IBKR PortfolioAnalyst reports furnished (CY2025 + FY2025-26), not full Activity "
        "Statements. Working CSVs synthesized for the Schedule FA generator. Prefer full "
        "IBKR Activity / Trade Confirmation statements if available for filing.",
    ),
    (
        "Account opening",
        "First funding 23-May-2025 via FOP In of SHOP × 84. Base currency USD. "
        "Further FOP In 26-May-2025: BGF World Financials (LU0106831901) × 792.45 and "
        "Fidelity Global Technology (LU0099574567) × 511.42. ACATS In GOOG × 105 on 01-Dec-2025.",
    ),
    (
        "Holdings CY2025 YE",
        "LU0099574567 (EUR) × 511.42; LU0106831901 × 792.45; GOOG × 105; SHOP × 84; cash ~USD 16.54. "
        "Ending NAV USD 156,356.24 (PortfolioAnalyst).",
    ),
    (
        "Holdings 31-Mar-2026",
        "Same quantities; Ending NAV ~USD 138,777 (PortfolioAnalyst).",
    ),
    (
        "Capital gains",
        "Nil taxable stock/fund sales in FY2025-26 (inbound transfers only).",
    ),
    (
        "Dividends",
        "GOOG USD 22.05 on 15-Dec-2025 (CY + FY) and USD 22.05 on 16-Mar-2026 (FY only). "
        "Total FY dividends USD 44.10. No withholding lines in PortfolioAnalyst — confirm FTC.",
    ),
    (
        "Cost / acquisition dates",
        "Cost basis from PortfolioAnalyst Open Position Summary (carryover). Acquisition dates "
        "in workbook = IBKR inbound transfer dates. Confirm original buy dates at the transferring "
        "broker for LTCG clock if any future sales.",
    ),
    (
        "EUR fund",
        "Fidelity LU0099574567 is EUR-denominated; A3/A2 use IBKR YE EUR→USD 1.1746 then Rule 115.",
    ),
]


if __name__ == "__main__":
    main(
        inception=INCEPTION,
        annual=ANNUAL,
        fiscal=FISCAL,
        out=OUT,
        account_open_date="2025-05-23 (FOP In SHOP × 84)",
        notes_extra=NOTES,
        assessee_label="Yagyank Chadha (U20291582)",
    )
    print(f"Wrote {OUT}")
