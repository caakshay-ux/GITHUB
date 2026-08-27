# Schedule FA / Capital Gains / Dividends — IBKR to ITR (India)

This repository contains workings to convert Interactive Brokers (IBKR) Activity Statements into Indian ITR schedules for a resident individual assessee.

## Clients (AY 2026-27)

| Assessee | IBKR | Output |
| --- | --- | --- |
| Ray G Stephanos | U15124027 | [`output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx) |
| Puneet Kohli | U17334752 (MSFT) | [`output/Foreign_Assets_Schedule_FA_Puneet_Kohli_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Puneet_Kohli_AY2026-27.xlsx) |
| Puneet Kohli | U16755051 (main) | [`output/Foreign_Assets_Schedule_FA_Puneet_Kohli_U16755051_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Puneet_Kohli_U16755051_AY2026-27.xlsx) |
| Aditya Malik | U15181144 | [`output/Foreign_Assets_Schedule_FA_Aditya_Malik_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Aditya_Malik_AY2026-27.xlsx) |
| Vinodkrishna Poyyale | U15388294 | [`output/Foreign_Assets_Schedule_FA_Vinod_Krishna_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Vinod_Krishna_AY2026-27.xlsx) |
| Abhisek Banerjee | U20221090 | [`output/Foreign_Assets_Schedule_FA_Abhisek_Banerjee_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Abhisek_Banerjee_AY2026-27.xlsx) |
| Phanindra V Gottipati | U15817316 | [`output/Foreign_Assets_Schedule_FA_Phanindra_V_Gottipati_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Phanindra_V_Gottipati_AY2026-27.xlsx) |
| Gunjan Narulkar (Paid) | U16931511 | [`output/Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U16931511_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U16931511_AY2026-27.xlsx) |
| Gunjan Narulkar (Free) | U22995548 | [`output/Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U22995548_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U22995548_AY2026-27.xlsx) |
| Rajani J Vallath (Single) | U16003525 | [`output/Foreign_Assets_Schedule_FA_Rajani_J_Vallath_U16003525_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Rajani_J_Vallath_U16003525_AY2026-27.xlsx) |
| Rajani J Vallath & Sanjeev K Nair (Joint) | U21864112 | [`output/Foreign_Assets_Schedule_FA_Rajani_Sanjeev_Joint_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Rajani_Sanjeev_Joint_AY2026-27.xlsx) |
| Arulselvam Chandrasekaran (Individual) | U22748155 | [`output/Foreign_Assets_Schedule_FA_Arulselvam_Chandrasekaran_U22748155_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Arulselvam_Chandrasekaran_U22748155_AY2026-27.xlsx) |
| Arulselvam Chandrasekaran & Dhanalakshmi S (Joint) | U22929455 | [`output/Foreign_Assets_Schedule_FA_Arulselvam_Dhanalakshmi_Joint_U22929455_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Arulselvam_Dhanalakshmi_Joint_U22929455_AY2026-27.xlsx) |
| Yagyank Chadha | U20291582 | [`output/Foreign_Assets_Schedule_FA_Yagyank_Chadha_U20291582_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Yagyank_Chadha_U20291582_AY2026-27.xlsx) |
| Yagyank Chadha | U15172057 | [`output/Foreign_Assets_Schedule_FA_Yagyank_Chadha_U15172057_AY2026-27.xlsx`](output/Foreign_Assets_Schedule_FA_Yagyank_Chadha_U15172057_AY2026-27.xlsx) |

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

### Source inputs — Vinodkrishna Poyyale U15388294

- `input_Inception_FY2024-25_Vinod_Krishna.csv` — 10-Sep-2024 to 31-Mar-2025
- `input_Fiscal_Statement_Vinod_Krishna.csv` — FY 2025-26
- `input_Annual_Statement_Vinod_Krishna.csv` — **synthesized** CY 2025
- Also under `clients/Vinod_Krishna_U15388294/`

### Source inputs — Aditya Malik U15181144

- `input_Inception_FY2024-25_Aditya_Malik.csv` — 20-Aug-2024 to 31-Mar-2025
- `input_Fiscal_Statement_Aditya_Malik.csv` — FY 2025-26
- `input_Annual_Statement_Aditya_Malik.csv` — **synthesized** CY 2025
- Also under `clients/Aditya_Malik_U15181144/`

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
# Aditya Malik — U15181144
python3 scripts/prepare_schedule_fa_aditya_malik.py
# Vinodkrishna Poyyale — U15388294
python3 scripts/prepare_schedule_fa_vinod_krishna.py
# Gunjan Narulkar — U16931511 (Paid)
python3 scripts/prepare_schedule_fa_gunjan_paid.py
# Gunjan Narulkar — U22995548 (Free)
python3 scripts/prepare_schedule_fa_gunjan_free.py
# Abhisek Banerjee — U20221090
python3 scripts/prepare_schedule_fa_abhisek_banerjee.py
# Phanindra V Gottipati — U15817316
python3 scripts/prepare_schedule_fa_phanindra.py
# Rajani / Sanjeev Joint — U21864112
python3 scripts/prepare_schedule_fa_rajani_sanjeev_joint.py
# Rajani J Vallath — U16003525 (Single)
python3 scripts/prepare_schedule_fa_rajani_single.py
# Arulselvam Chandrasekaran — U22748155 (Individual)
python3 scripts/prepare_schedule_fa_arulselvam_individual.py
# Arulselvam & Dhanalakshmi — U22929455 (Joint)
python3 scripts/prepare_schedule_fa_arulselvam_joint.py
# Yagyank Chadha — U20291582 (satellite)
python3 scripts/prepare_schedule_fa_yagyank_u20291582.py
# Yagyank Chadha — U15172057 (main)
python3 scripts/prepare_schedule_fa_yagyank_u15172057.py
# Both Yagyank accounts
python3 scripts/prepare_schedule_fa_yagyank_chadha.py
```




### Source inputs — Rajani J Vallath (Single) U16003525

- `input_Inception_FY2025-26_Rajani_Single.csv` / `input_Fiscal_Statement_Rajani_Single.csv` — 18-Nov-2025–31-Mar-2026
- `input_Annual_Statement_Rajani_Single.csv` — **synthesized** CY 2025 (YE AVGO mark from ACATS 18-Dec-2025)
- Also under `clients/Rajani_J_Vallath_U16003525/`
- Output: `output/Foreign_Assets_Schedule_FA_Rajani_J_Vallath_U16003525_AY2026-27.xlsx`

### Source inputs — Yagyank Chadha U20291582

- Real IBKR Activity Statement FY2025-26 (23-May-2025–31-Mar-2026): `input_Fiscal_Statement_Yagyank_U20291582.csv`
- CY2025 Annual synthesized (YE marks from PortfolioAnalyst CY2025)
- Also under `clients/Yagyank_Chadha_U20291582/`
- Output: `output/Foreign_Assets_Schedule_FA_Yagyank_Chadha_U20291582_AY2026-27.xlsx`

### Source inputs — Yagyank Chadha U15172057

- `input_Inception_FY2024-25_Yagyank_U15172057.csv` — 03-Sep-2024–31-Mar-2025
- `input_Fiscal_Statement_Yagyank_U15172057.csv` — FY 2025-26
- `input_Annual_Statement_Yagyank_U15172057.csv` — **synthesized** CY2025 (YE marks from 31-Mar-2025)
- Also under `clients/Yagyank_Chadha_U15172057/`
- Output: `output/Foreign_Assets_Schedule_FA_Yagyank_Chadha_U15172057_AY2026-27.xlsx`

### Source inputs — Arulselvam Chandrasekaran (Individual) U22748155

- `input_Annual_Statement_Arulselvam_U22748155.csv` / `input_Inception_FY2025-26_Arulselvam_U22748155.csv` — real IBKR Annual CY2025 (opened Nov-2025)
- `input_Fiscal_Statement_Arulselvam_U22748155.csv` — FY 2025-26
- Working CSVs: FOP GOOG cost set to sell Basis total USD 52,506.73; same-day Internal GOOG wash stripped
- Also under `clients/Arulselvam_Chandrasekaran_U22748155/`
- Output: `output/Foreign_Assets_Schedule_FA_Arulselvam_Chandrasekaran_U22748155_AY2026-27.xlsx`

### Source inputs — Arulselvam Chandrasekaran & Dhanalakshmi S (Joint) U22929455

- `input_Annual_Statement_Arulselvam_Joint_U22929455.csv` / inception copy — real IBKR Annual CY2025
- `input_Fiscal_Statement_Arulselvam_Joint_U22929455.csv` — FY 2025-26
- Working CSVs: same-day Internal GOOG wash with U22748155 stripped
- Also under `clients/Arulselvam_Dhanalakshmi_Joint_U22929455/`
- Output: `output/Foreign_Assets_Schedule_FA_Arulselvam_Dhanalakshmi_Joint_U22929455_AY2026-27.xlsx`

### Source inputs — Rajani J Vallath & Sanjeev K Nair (Joint) U21864112

- `input_Inception_FY2025-26_Rajani_Sanjeev_Joint.csv` / `input_Fiscal_Statement_Rajani_Sanjeev_Joint.csv` — 30-Jan-2026–31-Mar-2026 (opened via AVGO Internal In)
- `input_Annual_Statement_Rajani_Sanjeev_Joint.csv` — **synthesized empty** CY 2025 (account not open in CY2025)
- Also under `clients/Rajani_Sanjeev_Joint_U21864112/`
- Output: `output/Foreign_Assets_Schedule_FA_Rajani_Sanjeev_Joint_AY2026-27.xlsx`

### Source inputs — Phanindra V Gottipati U15817316

- `input_Inception_FY2024-25_Phanindra.csv` — 09-Oct-2024–31-Mar-2025
- `input_Fiscal_Statement_Phanindra.csv` — FY 2025-26
- `input_Annual_Statement_Phanindra.csv` — **synthesized** CY 2025 (YE marks from 31-Mar-2025)
- Also under `clients/Phanindra_V_Gottipati_U15817316/`
- Output: `output/Foreign_Assets_Schedule_FA_Phanindra_V_Gottipati_AY2026-27.xlsx`

### Source inputs — Abhisek Banerjee U20221090

- `input_Inception_FY2025-26_Abhisek_Banerjee.csv` / `input_Fiscal_Statement_Abhisek_Banerjee.csv` — same FY statement 16-May-2025–31-Mar-2026 (prior NAV 0)
- `input_Annual_Statement_Abhisek_Banerjee.csv` — **synthesized** CY 2025 (YE at cost)
- Also under `clients/Abhisek_Banerjee_U20221090/`
- Output: `output/Foreign_Assets_Schedule_FA_Abhisek_Banerjee_AY2026-27.xlsx`

### Source inputs — Gunjan Narulkar

**U16931511 (Paid)** — `clients/Gunjan_Narulkar_U16931511/` + root copies:
- `input_Inception_FY2024-25_Gunjan_Paid.csv` — 31-Dec-2024–31-Mar-2025
- `input_Fiscal_Statement_Gunjan_Paid.csv` — FY 2025-26
- `input_Annual_Statement_Gunjan_Paid.csv` — **synthesized** CY 2025
- Output: `output/Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U16931511_AY2026-27.xlsx`

**U22995548 (Free)** — `clients/Gunjan_Narulkar_U22995548/`:
- `input_Since_Inception_Gunjan_Free.csv` — 03-Dec-2025–19-Aug-2026 (GOOG × 179 via Internal In)
- `input_Annual_Statement_Gunjan_Free.csv` — **synthesized** CY 2025
- Output: `output/Foreign_Assets_Schedule_FA_Gunjan_Narulkar_U22995548_AY2026-27.xlsx`

### Key caveats (read Notes for CA)

1. SBI TT Buy rates are month-end card compilations — confirm exact card rates on sbi.co.in for filing.
2. Peak NAV uses max(start, end) — use PortfolioAnalyst for true peak.
3. Foreign listed shares (no Indian STT): holding period > 24 months → LTCG @ 12.5% (post 23-Jul-2024); else STCG at slab rates.
4. Kohli: MSFT via FOP 29-Jan-2025 — original buy date outside IBKR not available; confirm for LTCG clock. Prefer IBKR Annual CY2025 for YE closing.
5. Gunjan Paid: GOOG 216 held before statement start — seeded 31-Dec-2024 @ USD 19,300.32; confirm original buy date for LTCG. Internal Out of GOOG 179 to Free (03-Dec-2025) excluded from taxable CG.
6. Gunjan Free: YE close proxied at cost; no trades / nil CG.
7. Abhisek: opened mid-May 2025; nil CG (buys only); YE2025 marks at cost — prefer IBKR Annual for filing.
8. Phanindra: nil CG (buys only); YE2025 marks proxied from 31-Mar-2025 (GLD at cost); prefer IBKR Annual for filing.
9. Rajani/Sanjeev Joint: opened 30-Jan-2026 — FA CY2025 N/A; AVGO sale 04-Mar-2026 — confirm original buy date at U16003525 for LTCG (IBKR tagged L).
10. Rajani Single U16003525: ACATS AVGO In 18-Dec-2025; Internal Out to Joint excluded from CG; YE mark from ACATS price; confirm original AVGO buy date at 27960235 for LTCG on Mar-2026 sale.
11. Arulselvam Individual U22748155 (SGD base): GOOG FOP × 447 then full sale 10-Nov-2025; cost = IBKR Basis USD 52,506.73; IBKR used Highest Cost (HC) — confirm original buy dates for LTCG (workbook uses FOP date → STCG). Peak NAV max(start,end) understates mid-year GOOG holding — use PortfolioAnalyst.
12. Arulselvam Joint U22929455 (SGD base): nil stock CG; YE marks from real Annual; A2 NAV SGD→USD via YE USD.SGD 1.286.
13. Yagyank Chadha U20291582: real FY Activity Statement; Internal In from U15172057; FOP GOOG; nil CG; YE from PortfolioAnalyst.
14. Yagyank Chadha U15172057: multi-currency portfolio; taxable CG on BLBD/NOVd/ORCL/1810; Internal Out of SHOP+Lux UCITS to U20291582 excluded (code I); YE marks proxied from 31-Mar-2025.