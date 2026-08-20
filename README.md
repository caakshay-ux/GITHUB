# Schedule FA / Capital Gains / Dividends — IBKR to ITR (India)

This repository contains workings to convert Interactive Brokers (IBKR) Activity Statements into Indian ITR schedules for a resident individual assessee.

## Clients (AY 2026-27)

| Assessee | IBKR | Output |
| --- | --- | --- |
| Ray G Stephanos | U15124027 | [`output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx) |
| Puneet Kohli | U17334752 (MSFT) | [`output/Foreign_Assets_Schedule_FA_Puneet_Kohli_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Puneet_Kohli_AY2026-27.xlsx) |
| Puneet Kohli | U16755051 (main) | [`output/Foreign_Assets_Schedule_FA_Puneet_Kohli_U16755051_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Puneet_Kohli_U16755051_AY2026-27.xlsx) |

### Workbook sheets

| Sheet | Purpose |
| --- | --- |
| A3 | Schedule FA – Foreign equity/debt interests (INR) for CY 2025; Col C = Symbol only |
| A3 (USD reference) | Same holdings in USD before FX conversion |
| FX Rates (USD-INR) | SBI TT Buy rates used + methodology notes |
| A2 Custodial Account | Schedule FA – IBKR custodial account |
| Schedule OS - CY2025 | Interest & dividends for calendar year 2025 |
| Schedule OS - FY2025-26 | Interest & dividends for FY 2025-26 (ITR income) |
| FTC - Withholding Tax | Foreign tax withheld (Schedule FSI/TR input) |
| Capital Gains FY2025-26 | **Formula-driven** FIFO CG — editable FX (yellow); Gain = Sale−Comm−Cost INR |
| Capital Gains FY2024-25 | **Formula-driven** FIFO CG for prior FY (AY 2025-26) |
| FX Lookup | Editable SBI TT USD/INR & EUR/INR rates + FCY→USD cross rates |
| FIFO Lots Register | Lot qty / acquisition date / cost as of 31-Dec-2024 & 31-Mar-2025 |
| CG Summary by Symbol | Symbol-wise STCG/LTCG summary (FY2025-26) |
| Notes for CA | Assumptions, caveats, action items |

### Source inputs — Ray G Stephanos

- `input_Inception_FY2024-25_Ray.csv` — IBKR Activity Statement since funding (01-Apr-2024 to 31-Mar-2025)
- `input_Annual_Statement_Ray.csv` — IBKR Activity Statement CY 2025
- `input_Fiscal_Statement_Ray.csv` — IBKR Activity Statement FY 2025-26

### Source inputs — Puneet Kohli U16755051 (main)

- `input_Inception_FY2024-25_Kohli_U16755051.csv` — FY 2024-25
- `input_Fiscal_Statement_Kohli_U16755051.csv` — FY 2025-26
- `input_Annual_Statement_Kohli_U16755051.csv` — **synthesized** CY 2025
- Also under `clients/Puneet_Kohli_U16755051/`

### Source inputs — Puneet Kohli U17334752 (MSFT-only)

- `input_Inception_FY2024-25_Kohli.csv` — FY 2024-25 (FOP transfer of MSFT)
- `input_Fiscal_Statement_Kohli.csv` — FY 2025-26
- `input_Annual_Statement_Kohli.csv` — **synthesized** CY 2025 (no Annual furnished; YE MSFT close 481.48)
- Also under `clients/Puneet_Kohli/`

### FIFO method

Lots are rebuilt from inception buys **and inbound FOP transfers** (IBKR cost basis). Splits applied where present (e.g. LRCX, NFLX for Ray).  
Cost = IBKR trade/transfer Basis; sale = Proceeds; sell commission deducted. INR via Rule 115.

### Regenerate

```bash
pip install openpyxl
# Ray
python3 scripts/prepare_schedule_fa_ray.py
# Puneet Kohli — U17334752 (MSFT)
python3 scripts/prepare_schedule_fa_puneet_kohli.py
# Puneet Kohli — U16755051 (main)
python3 scripts/prepare_schedule_fa_puneet_kohli_U16755051.py
```

### Key caveats (read Notes for CA)

1. SBI TT Buy rates are month-end card compilations — confirm exact card rates on sbi.co.in for filing.
2. Peak NAV uses max(start, end) — use PortfolioAnalyst for true peak.
3. Foreign listed shares (no Indian STT): holding period > 24 months → LTCG @ 12.5% (post 23-Jul-2024); else STCG at slab rates.
4. Kohli: MSFT via FOP 29-Jan-2025 — original buy date outside IBKR not available; confirm for LTCG clock. Prefer IBKR Annual CY2025 for YE closing.