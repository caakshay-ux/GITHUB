# Schedule FA / Capital Gains / Dividends — IBKR to ITR (India)

This repository contains workings to convert Interactive Brokers (IBKR) Activity Statements into Indian ITR schedules for a resident individual assessee.

## Client deliverable (AY 2026-27)

**Assessee:** Ray G Stephanos  
**IBKR Account:** U15124027 (Advisor: Financial Hospital Advisory LLP)  
**Output file:** [`output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx)

### Workbook sheets

| Sheet | Purpose |
| --- | --- |
| A3 | Schedule FA – Foreign equity/debt interests (INR) for CY 2025 |
| A3 (USD reference) | Same holdings in USD before FX conversion |
| FX Rates (USD-INR) | SBI TT Buy rates used + methodology notes |
| A2 Custodial Account | Schedule FA – IBKR custodial account |
| Schedule OS - CY2025 | Interest & dividends for calendar year 2025 |
| Schedule OS - FY2025-26 | Interest & dividends for FY 2025-26 (ITR income) |
| FTC - Withholding Tax | Foreign tax withheld (Schedule FSI/TR input) |
| Capital Gains FY2025-26 | FIFO capital gains with Rule 115 INR conversion |
| CG Summary by Symbol | Symbol-wise STCG/LTCG summary |
| Notes for CA | Assumptions, caveats, action items |

### Source inputs

- `input_Annual_Statement_Ray.csv` — IBKR Activity Statement CY 2025 (01-Jan-2025 to 31-Dec-2025)
- `input_Fiscal_Statement_Ray.csv` — IBKR Activity Statement FY 2025-26 (01-Apr-2025 to 31-Mar-2026)

### Regenerate

```bash
pip install openpyxl
python3 scripts/prepare_schedule_fa_ray.py
```

### Key caveats (read Notes for CA)

1. Pre-2025 lot acquisition dates are placeholders (`2024-07-01`) — replace from prior IBKR statements before finalising LTCG vs STCG.
2. SBI TT Buy rates are month-end card compilations — confirm exact card rates on sbi.co.in for filing.
3. Peak NAV uses max(start, end) because the Activity Statement has no daily NAV series — use PortfolioAnalyst for true peak.
4. Foreign listed shares (no Indian STT): holding period > 24 months → LTCG @ 12.5% (post 23-Jul-2024); else STCG at slab rates.
