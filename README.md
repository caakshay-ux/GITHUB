# Schedule FA / Capital Gains / Dividends — IBKR to ITR (India)

This repository contains workings to convert Interactive Brokers (IBKR) Activity Statements into Indian ITR schedules for a resident individual assessee.

## Client deliverable (AY 2026-27 / AY 2025-26)

**Assessee:** Ray G Stephanos  
**IBKR Account:** U15124027 (Advisor: Financial Hospital Advisory LLP)  
**Output file:** [`output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx)

### Workbook sheets

| Sheet | Purpose |
| --- | --- |
| A3 | Schedule FA – Foreign equity/debt interests (INR) for CY 2025; Col C = Symbol only |
| A3 (USD reference) | Same holdings in USD before FX conversion |
| FX Rates (USD-INR) | SBI TT Buy rates used + methodology notes |
| A2 Custodial Account | Schedule FA – IBKR custodial account (opened 06-Sep-2024) |
| Schedule OS - CY2025 | Interest & dividends for calendar year 2025 |
| Schedule OS - FY2025-26 | Interest & dividends for FY 2025-26 (ITR income) |
| FTC - Withholding Tax | Foreign tax withheld (Schedule FSI/TR input) |
| Capital Gains FY2025-26 | **Formula-driven** FIFO CG — editable FX (yellow); Gain = Sale−Comm−Cost INR |
| Capital Gains FY2024-25 | **Formula-driven** FIFO CG for prior FY (AY 2025-26) |
| FX Lookup | Editable SBI TT USD/INR & EUR/INR rates + FCY→USD cross rates |
| FIFO Lots Register | Lot qty / acquisition date / cost as of 31-Dec-2024 & 31-Mar-2025 |
| CG Summary by Symbol | Symbol-wise STCG/LTCG summary (FY2025-26) |
| Notes for CA | Assumptions, caveats, action items |

### Source inputs

- `input_Inception_FY2024-25_Ray.csv` — IBKR Activity Statement since funding (01-Apr-2024 to 31-Mar-2025)
- `input_Annual_Statement_Ray.csv` — IBKR Activity Statement CY 2025
- `input_Fiscal_Statement_Ray.csv` — IBKR Activity Statement FY 2025-26

### FIFO method

Lots are rebuilt from inception buys. Splits applied (LRCX 10:1 on 02-Oct-2024; NFLX 10:1 on 17-Nov-2025).  
Cost = IBKR trade Basis; sale = Proceeds; sell commission deducted. INR via Rule 115.

### Regenerate

```bash
pip install openpyxl
python3 scripts/prepare_schedule_fa_ray.py
```

### Key caveats (read Notes for CA)

1. SBI TT Buy rates are month-end card compilations — confirm exact card rates on sbi.co.in for filing.
2. Peak NAV uses max(start, end) — use PortfolioAnalyst for true peak.
3. Foreign listed shares (no Indian STT): holding period > 24 months → LTCG @ 12.5% (post 23-Jul-2024); else STCG at slab rates. Account funded Sep-2024 → all FY2025-26 disposals are STCG.
