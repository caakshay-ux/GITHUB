#!/usr/bin/env python3
"""
Prepare Schedule FA (A2/A3), capital gains and dividend/interest workings
for Ray G Stephanos (IBKR U15124027) from Activity Statements.

Source statements:
  - Inception / FY 2024-25 (01-Apr-2024 to 31-Mar-2025) — true buy dates & opening lots
  - Annual CY 2025 (01-Jan-2025 to 31-Dec-2025) — Schedule FA
  - Fiscal FY 2025-26 (01-Apr-2025 to 31-Mar-2026) — ITR income & capital gains

FIFO capital gains use actual acquisition dates from inception trades (no placeholders).
"""

from __future__ import annotations

import csv
import re
from collections import defaultdict, deque
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime, date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter

ROOT = Path("/workspace")
INCEPTION = ROOT / "input_Inception_FY2024-25_Ray.csv"
ANNUAL = ROOT / "input_Annual_Statement_Ray.csv"
FISCAL = ROOT / "input_Fiscal_Statement_Ray.csv"
OUT = ROOT / "output" / "Foreign_Assets_Schedule_FA_Ray_G_Stephanos_AY2026-27.xlsx"

# ---------------------------------------------------------------------------
# SBI TT Buying rates (USD & EUR) — month-end card rates
# Source: public SBI TTBR compilations (CA must verify exact card rate for the date)
# ---------------------------------------------------------------------------
SBI_TT_USD = {
    date(2024, 8, 31): 83.50,
    date(2024, 9, 30): 83.30,
    date(2024, 10, 31): 83.68,
    date(2024, 11, 30): 84.15,
    date(2024, 12, 31): 85.20,
    date(2025, 1, 31): 86.20,
    date(2025, 2, 28): 86.95,
    date(2025, 3, 29): 85.10,  # month-end card
    date(2025, 3, 31): 85.10,
    date(2025, 4, 30): 84.25,
    date(2025, 5, 31): 85.10,
    date(2025, 6, 30): 85.10,
    date(2025, 7, 31): 87.15,
    date(2025, 8, 30): 87.70,
    date(2025, 8, 31): 87.70,
    date(2025, 9, 30): 88.35,
    date(2025, 10, 31): 88.20,
    date(2025, 11, 29): 88.95,
    date(2025, 11, 30): 88.95,
    date(2025, 12, 31): 89.47,
    date(2026, 1, 31): 91.35,
    date(2026, 2, 28): 90.56,
    date(2026, 3, 31): 93.15,
}

SBI_TT_EUR = {
    date(2024, 8, 31): 91.72,
    date(2024, 9, 30): 92.34,
    date(2024, 10, 31): 90.07,
    date(2024, 11, 30): 88.26,
    date(2024, 12, 31): 87.93,
    date(2025, 1, 31): 88.78,
    date(2025, 2, 28): 89.55,
    date(2025, 3, 31): 91.86,
    date(2025, 4, 30): 94.97,
    date(2025, 5, 31): 95.78,
    date(2025, 6, 30): 98.92,
    date(2025, 7, 31): 99.02,
    date(2025, 8, 31): 101.70,
    date(2025, 9, 30): 103.09,
    date(2025, 10, 31): 101.23,
    date(2025, 11, 30): 102.40,
    date(2025, 12, 31): 104.20,
    date(2026, 1, 31): 105.00,
    date(2026, 2, 28): 104.50,
    date(2026, 3, 31): 106.00,
}

# Approximate IBKR year-end FX to USD (from Annual Forex Balances close)
FX_TO_USD_YE2025 = {
    "USD": 1.0,
    "EUR": 1.1746,
    "HKD": 0.12849,
    "JPY": 0.0063827,
    "DKK": 0.15726,
    # IBKR YE USD.SGD ≈ 1.2860 → SGD per USD; invert for SGD→USD
    "SGD": 1.0 / 1.2860,
}

# Country metadata for Schedule FA
COUNTRY = {
    "United States of America": ("United States of America", 2),
    "Netherlands": ("Netherlands", 156),
    "Denmark": ("Denmark", 61),
    "Japan": ("Japan", 111),
    "United Kingdom": ("United Kingdom", 234),
    "Germany": ("Germany", 78),
    "France": ("France", 75),
    "Hong Kong": ("Hong Kong", 94),
    "China": ("China", 44),
    "South Korea": ("Korea, Republic of", 116),
    "Ireland": ("Ireland", 101),
    "Cayman Islands": ("Cayman Islands", 40),
    "Taiwan": ("Taiwan", 212),
    "Australia": ("Australia", 12),
    "Luxembourg": ("Luxembourg", 127),
    "Canada": ("Canada", 36),
}

# Entity master: symbol -> (country_key, nature, address, zip, legal_name)
ENTITY = {
    "AAPL": ("United States of America", "Listed Foreign Equity Share (Company)",
             "One Apple Park Way, Cupertino, CA", "95014", "Apple Inc - Common Stock"),
    "AAXJ": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
             "iShares MSCI All Country Asia ex Japan ETF"),
    "AEP": ("United States of America", "Listed Foreign Equity Share (Company)",
            "1 Riverside Plaza, Columbus, OH", "43215", "American Electric Power Co Inc"),
    "ACN": ("Ireland", "Listed Foreign Equity Share (Company)",
            "1 Grand Canal Square, Grand Canal Harbour, Dublin", "D02 P820", "Accenture plc - Class A"),
    "ADBE": ("United States of America", "Listed Foreign Equity Share (Company)",
             "345 Park Avenue, San Jose, CA", "95110", "Adobe Inc"),
    "AIQ": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o Global X Management Company LLC, 605 Third Avenue, New York, NY", "10158",
            "Global X Artificial Intelligence & Technology ETF"),
    "AIRd": ("Netherlands", "Listed Foreign Equity Share (Company)",
             "2 rond-point Emile Dewoitine, Blagnac / Airbus SE registered NL", "31700",
             "Airbus SE"),
    "AIR": ("Netherlands", "Listed Foreign Equity Share (Company)",
            "2 rond-point Emile Dewoitine, Blagnac / Airbus SE registered NL", "31700",
            "Airbus SE"),
    "AMD": ("United States of America", "Listed Foreign Equity Share (Company)",
            "2485 Augustine Drive, Santa Clara, CA", "95054", "Advanced Micro Devices Inc"),
    "AMZN": ("United States of America", "Listed Foreign Equity Share (Company)",
             "410 Terry Avenue North, Seattle, WA", "98109", "Amazon.com Inc"),
    "ASML": ("Netherlands", "American Depository Receipt of Listed Foreign Company",
             "De Run 6501, 5504 DR Veldhoven", "5504 DR",
             "ASML Holding NV - NY Registered Shares"),
    "AVGO": ("United States of America", "Listed Foreign Equity Share (Company)",
             "3421 Hillview Avenue, Palo Alto, CA", "94304", "Broadcom Inc"),
    "AXP": ("United States of America", "Listed Foreign Equity Share (Company)",
            "200 Vesey Street, New York, NY", "10285", "American Express Co"),
    "BAC": ("United States of America", "Listed Foreign Equity Share (Company)",
            "100 North Tryon Street, Charlotte, NC", "28255", "Bank of America Corp"),
    "BLBD": ("United States of America", "Listed Foreign Equity Share (Company)",
             "3920 Arkwright Road, Macon, GA", "31210", "Blue Bird Corp"),
    "BNO": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o United States Commodity Funds LLC", "", "United States Brent Oil Fund LP"),
    "BOTZ": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o Global X Management Company LLC, 605 Third Avenue, New York, NY", "10158",
             "Global X Robotics & Artificial Intelligence ETF"),
    "BTI": ("United Kingdom", "American Depository Receipt of Listed Foreign Company",
            "Globe House, 4 Temple Place, London", "WC2R 2PG",
            "British American Tobacco plc - Sponsored ADR"),
    "BX": ("United States of America", "Listed Foreign Equity Share (Company)",
           "345 Park Avenue, New York, NY", "10154", "Blackstone Inc"),
    "BYDDY": ("China", "American Depository Receipt of Listed Foreign Company",
              "No. 3009 BYD Road, Pingshan, Shenzhen", "518118",
              "BYD Co Ltd - Unsponsored ADR"),
    "CAT": ("United States of America", "Listed Foreign Equity Share (Company)",
            "5205 N O'Connor Boulevard, Irving, TX", "75039", "Caterpillar Inc"),
    "CLPT": ("United States of America", "Listed Foreign Equity Share (Company)",
             "120 S Sierra Avenue, Solana Beach, CA", "92075", "ClearPoint Neuro Inc"),
    "CPNG": ("United States of America", "Listed Foreign Equity Share (Company)",
             "720 Olive Way, Suite 600, Seattle, WA", "98101",
             "Coupang Inc (Delaware) - Class A"),
    "COPX": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o Global X Management Company LLC, 605 Third Avenue, New York, NY", "10158",
             "Global X Copper Miners ETF"),
    "CRCL": ("United States of America", "Listed Foreign Equity Share (Company)",
             "99 High Street, Boston, MA", "02110", "Circle Internet Group Inc"),
    "CNUA": (
        "Ireland",
        "Exchange Traded Fund (Investment Trust)",
        "UBS Fund Management (Ireland) Ltd, Dublin",
        "D02 H738",
        "UBS ETF MSCI China A USD Acc",
    ),
    "CSCO": ("United States of America", "Listed Foreign Equity Share (Company)",
            "170 West Tasman Drive, San Jose, CA", "95134", "Cisco Systems Inc"),
    "DB": ("Germany", "Listed Foreign Equity Share (Company)",
           "Taunusanlage 12, Frankfurt am Main", "60325",
           "Deutsche Bank AG - Registered Shares"),
    "DUOL": ("United States of America", "Listed Foreign Equity Share (Company)",
             "5900 Penn Avenue, Pittsburgh, PA", "15206", "Duolingo Inc"),
    "DXJ": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o WisdomTree Investments Inc, 250 West 34th Street, New York, NY", "10119",
            "WisdomTree Japan Hedged Equity Fund"),
    "EMXC": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
             "iShares MSCI Emerging Markets ex China ETF"),
    "GD": ("United States of America", "Listed Foreign Equity Share (Company)",
           "11011 Sunset Hills Road, Reston, VA", "20190", "General Dynamics Corp"),
    "GE": ("United States of America", "Listed Foreign Equity Share (Company)",
           "1 Neumann Way, Cincinnati, OH", "45215", "GE Aerospace / General Electric"),
    "GLD": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o World Gold Trust Services LLC, 685 Third Avenue, New York, NY", "10017",
            "SPDR Gold Shares (World Gold Trust)"),
    "GOOG": ("United States of America", "Listed Foreign Equity Share (Company)",
             "1600 Amphitheatre Parkway, Mountain View, CA", "94043",
             "Alphabet Inc - Class C"),
    "GSL": ("United States of America", "Listed Foreign Equity Share (Company)",
            "c/o Global Ship Lease Inc / Marshall Islands registered (NYSE)", "",
            "Global Ship Lease Inc - Class A"),
    "LU0099574567": (
        "Luxembourg",
        "Foreign Mutual Fund (Investment Trust / UCITS)",
        "Fidelity Funds SICAV, 2a rue Albert Borschette, Luxembourg",
        "L-1246",
        "Fidelity Funds - Global Technology Fund A-INC (EUR)",
    ),
    "LU0106831901": (
        "Luxembourg",
        "Foreign Mutual Fund (Investment Trust / UCITS)",
        "BlackRock Global Funds, 2-4 rue Eugène Ruppert, Luxembourg",
        "L-2453",
        "BGF World Financials Fund A2 Acc (USD)",
    ),
    "SHOP": (
        "Canada",
        "Listed Foreign Equity Share (Company)",
        "151 O'Connor Street, Ground Floor, Ottawa, ON",
        "K2P 2L8",
        "Shopify Inc - Class A",
    ),
    "HON": ("United States of America", "Listed Foreign Equity Share (Company)",
            "855 S Mint Street, Charlotte, NC", "28202", "Honeywell International Inc"),
    "HYU": ("South Korea", "Global Depository Receipt of Listed Foreign Company",
            "12 Heolleung-ro, Seocho-gu, Seoul", "06797",
            "Hyundai Motor Co - Reg S GDR"),
    "IBM": ("United States of America", "Listed Foreign Equity Share (Company)",
            "1 New Orchard Road, Armonk, NY", "10504", "International Business Machines Corp"),
    "ICLN": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
             "iShares Global Clean Energy ETF"),
    "IEF": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
            "iShares 7-10 Year Treasury Bond ETF"),
    "IGPT": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o Invesco Capital Management LLC, Downers Grove, IL", "60515",
             "Invesco AI and Next Gen Software ETF"),
    "IJH": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
            "iShares Core S&P Mid-Cap ETF"),
    "INDA": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
             "iShares MSCI India ETF"),
    "INTC": ("United States of America", "Listed Foreign Equity Share (Company)",
             "2200 Mission College Boulevard, Santa Clara, CA", "95054", "Intel Corp"),
    "IWR": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
            "iShares Russell Mid-Cap ETF"),
    "IYW": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
            "iShares U.S. Technology ETF"),
    "LLY": ("United States of America", "Listed Foreign Equity Share (Company)",
            "Lilly Corporate Center, Indianapolis, IN", "46285", "Eli Lilly & Company"),
    "LRCX": ("United States of America", "Listed Foreign Equity Share (Company)",
             "4650 Cushing Parkway, Fremont, CA", "94538", "Lam Research Corp"),
    "MCHI": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
             "iShares MSCI China ETF"),
    "MC": ("France", "Listed Foreign Equity Share (Company)",
           "22 Avenue Montaigne, Paris", "75008",
           "LVMH Moet Hennessy Louis Vuitton SE"),
    "META": ("United States of America", "Listed Foreign Equity Share (Company)",
             "1 Meta Way, Menlo Park, CA", "94025", "Meta Platforms Inc - Class A"),
    "MSFT": ("United States of America", "Listed Foreign Equity Share (Company)",
             "One Microsoft Way, Redmond, WA", "98052", "Microsoft Corp"),
    "NBIS": ("Netherlands", "Listed Foreign Equity Share (Company)",
             "Schiphol Boulevard 165, 1118 BG Schiphol", "1118 BG", "Nebius Group NV"),
    "NFLX": ("United States of America", "Listed Foreign Equity Share (Company)",
             "121 Albright Way, Los Gatos, CA", "95032", "Netflix Inc"),
    "NOVd": ("Denmark", "Listed Foreign Equity Share (Company)",
             "Novo Alle 1, 2880 Bagsvaerd", "2880", "Novo Nordisk A/S - B Shares"),
    "NVDA": ("United States of America", "Listed Foreign Equity Share (Company)",
             "2788 San Tomas Expressway, Santa Clara, CA", "95051", "NVIDIA Corp"),
    "ORCL": ("United States of America", "Listed Foreign Equity Share (Company)",
             "2300 Oracle Way, Austin, TX", "78741", "Oracle Corp"),
    "PPLT": ("United States of America", "Exchange Traded Fund (Investment Trust)",
             "c/o abrdn ETFs Sponsor LLC / Platinum Trust, Philadelphia, PA", "19103",
             "abrdn Physical Platinum Shares ETF"),
    "QQQ": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o Invesco Capital Management LLC, Downers Grove, IL", "60515",
            "Invesco QQQ Trust Series 1"),
    "R6C0d": ("United Kingdom", "Listed Foreign Equity Share (Company)",
              "Shell Centre, 2 York Road, London", "SE1 7NA", "Shell plc"),
    "SHELL": ("United Kingdom", "Listed Foreign Equity Share (Company)",
              "Shell Centre, 2 York Road, London", "SE1 7NA", "Shell plc"),
    "RBRK": ("United States of America", "Listed Foreign Equity Share (Company)",
             "3495 Deer Creek Road, Palo Alto, CA", "94304", "Rubrik Inc - Class A"),
    "REGN": ("United States of America", "Listed Foreign Equity Share (Company)",
             "777 Old Saw Mill River Road, Tarrytown, NY", "10591",
             "Regeneron Pharmaceuticals Inc"),
    "RIO": ("United Kingdom", "American Depository Receipt of Listed Foreign Company",
            "6 St James's Square, London", "SW1Y 4AD", "Rio Tinto plc - Sponsored ADR"),
    "RTX": ("United States of America", "Listed Foreign Equity Share (Company)",
            "1000 Wilson Boulevard, Arlington, VA", "22209", "RTX Corp"),
    "SAUS": ("Ireland", "Exchange Traded Fund (Investment Trust)",
             "c/o BlackRock Asset Management Ireland Ltd, Dublin", "D01 W5P4",
             "iShares MSCI Australia UCITS ETF"),
    "SIEd": ("Germany", "Listed Foreign Equity Share (Company)",
             "Werner-von-Siemens-Strasse 1, Munich", "80333", "Siemens AG - Registered"),
    "SIE": ("Germany", "Listed Foreign Equity Share (Company)",
            "Werner-von-Siemens-Strasse 1, Munich", "80333", "Siemens AG - Registered"),
    "SLV": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o iShares Delaware Trust Sponsor LLC / BlackRock, San Francisco, CA", "94105",
            "iShares Silver Trust"),
    "SNOW": ("United States of America", "Listed Foreign Equity Share (Company)",
             "106 East Babcock Street, Bozeman, MT", "59715", "Snowflake Inc"),
    "TGT": ("United States of America", "Listed Foreign Equity Share (Company)",
            "1000 Nicollet Mall, Minneapolis, MN", "55403", "Target Corp"),
    "TIP": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o BlackRock Fund Advisors, 400 Howard Street, San Francisco, CA", "94105",
            "iShares TIPS Bond ETF"),
    "TSLA": ("United States of America", "Listed Foreign Equity Share (Company)",
             "1 Tesla Road, Austin, TX", "78725", "Tesla Inc"),
    "TSM": ("Taiwan", "American Depository Receipt of Listed Foreign Company",
            "8 Li-Hsin Road 6, Hsinchu Science Park, Hsinchu", "30078",
            "Taiwan Semiconductor Manufacturing Co - Sponsored ADR"),
    "URA": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o Global X Management Company LLC, 605 Third Avenue, New York, NY", "10158",
            "Global X Uranium ETF"),
    "V": ("United States of America", "Listed Foreign Equity Share (Company)",
          "P.O. Box 8999, San Francisco, CA", "94128", "Visa Inc - Class A"),
    "VOO": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o The Vanguard Group Inc, 100 Vanguard Boulevard, Malvern, PA", "19355",
            "Vanguard S&P 500 ETF"),
    "VKTX": ("United States of America", "Listed Foreign Equity Share (Company)",
             "9920 Pacific Heights Boulevard, San Diego, CA", "92121",
             "Viking Therapeutics Inc"),
    "VT": ("United States of America", "Exchange Traded Fund (Investment Trust)",
           "c/o The Vanguard Group Inc, 100 Vanguard Boulevard, Malvern, PA", "19355",
           "Vanguard Total World Stock ETF"),
    "VTV": ("United States of America", "Exchange Traded Fund (Investment Trust)",
            "c/o The Vanguard Group Inc, 100 Vanguard Boulevard, Malvern, PA", "19355",
            "Vanguard Value ETF"),
    "WCBR": ("Ireland", "Exchange Traded Fund (Investment Trust)",
             "c/o WisdomTree Issuer ICAV, Dublin", "D02 T860",
             "WisdomTree Cybersecurity UCITS ETF"),
    "W1TB": ("Ireland", "Exchange Traded Fund (Investment Trust)",
             "c/o WisdomTree Issuer ICAV, Dublin", "D02 T860",
             "WisdomTree Cybersecurity UCITS ETF"),
    "WFC": ("United States of America", "Listed Foreign Equity Share (Company)",
            "420 Montgomery Street, San Francisco, CA", "94104", "Wells Fargo & Company"),
    "XOM": ("United States of America", "Listed Foreign Equity Share (Company)",
            "22777 Springwoods Village Parkway, Spring, TX", "77389", "Exxon Mobil Corp"),
    "1810": ("Cayman Islands", "Listed Foreign Equity Share (Company)",
             "Xiaomi Corporation, Cayman Islands / Beijing ops", "", "Xiaomi Corp - Class B"),
    "1919": ("China", "Listed Foreign Equity Share (Company)",
             "COSCO SHIPPING Holdings Co Ltd, Shanghai", "", "COSCO Shipping Holdings Co - H"),
    "2914.T": ("Japan", "Listed Foreign Equity Share (Company)",
               "2-2-1 Toranomon, Minato-ku, Tokyo", "105-8422", "Japan Tobacco Inc"),
    "3069": ("Hong Kong", "Exchange Traded Fund (Investment Trust)",
             "ChinaAMC, Hong Kong", "", "ChinaAMC Hang Seng Biotech ETF"),
    "3088": ("Hong Kong", "Exchange Traded Fund (Investment Trust)",
             "ChinaAMC, Hong Kong", "", "ChinaAMC Hang Seng TECH Index ETF"),
    "3110": ("Hong Kong", "Exchange Traded Fund (Investment Trust)",
             "Global X ETFs / Mirae Asset, Hong Kong", "", "Global X Hang Seng High Dividend Yield ETF"),
    "3188": ("Hong Kong", "Exchange Traded Fund (Investment Trust)",
             "ChinaAMC, Hong Kong", "", "ChinaAMC CSI 300 Index ETF"),
    "3416": ("Hong Kong", "Exchange Traded Fund (Investment Trust)",
             "Global X ETFs / Mirae Asset, Hong Kong", "", "Global X HSCEI Covered Call ETF"),
    "4568.T": ("Japan", "Listed Foreign Equity Share (Company)",
               "3-5-1 Nihonbashi Honcho, Chuo-ku, Tokyo", "103-8426", "Daiichi Sankyo Co Ltd"),
    "7203.T": ("Japan", "Listed Foreign Equity Share (Company)",
               "1 Toyota-cho, Toyota City, Aichi", "471-8571", "Toyota Motor Corp"),
    "ZEALc": ("Denmark", "Listed Foreign Equity Share (Company)",
              "Sydmarken 11, 2860 Soeborg", "2860", "Zealand Pharma A/S"),
    "SHELL.DRS": ("United Kingdom", "Other Interest (Dividend Right)",
                  "Shell Centre, London", "SE1 7NA", "Shell plc - Dividend Rights"),
    "SHELL.DVD": ("United Kingdom", "Other Interest (Dividend Right)",
                  "Shell Centre, London", "SE1 7NA", "Shell plc - Dividend Rights"),
    "SHELL1.DI": ("United Kingdom", "Other Interest (Dividend Right)",
                  "Shell Centre, London", "SE1 7NA", "Shell plc - Dividend Rights"),
    "SHELL.DDR": ("United Kingdom", "Other Interest (Dividend Right)",
                  "Shell Centre, London", "SE1 7NA", "Shell plc - Dividend Rights"),
}


def parse_ibkr(path: Path):
    sections = defaultdict(list)
    headers = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            if len(row) < 2:
                continue
            # Some IBKR CSV exports pad trailing empty columns
            while len(row) > 2 and row[-1] == "":
                row = row[:-1]
            sec, kind = row[0], row[1]
            if kind == "Header":
                headers[sec] = row[2:]
            elif kind in ("Data", "Total", "SubTotal"):
                sections[sec].append((kind, row[2:]))
    return headers, sections


def fnum(x):
    if x is None or x == "" or x == "--":
        return None
    return float(str(x).replace(",", "").replace('"', ""))


def parse_dt(s: str) -> date:
    s = (s or "").strip()
    if not s:
        raise ValueError("empty date")
    # ISO first; also accept DD-MM-YYYY from some IBKR exports
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return datetime.strptime(s[:10], "%Y-%m-%d").date()


def month_end_on_or_before(d: date, table: dict) -> date:
    keys = sorted(k for k in table if k <= d)
    if not keys:
        return min(table)
    return keys[-1]


def sbi_usd(d: date) -> float:
    return SBI_TT_USD[month_end_on_or_before(d, SBI_TT_USD)]


def rule115_usd(event_date: date) -> float:
    """TT buy rate on last day of month preceding the month of receipt/transfer."""
    if event_date.month == 1:
        prev = date(event_date.year - 1, 12, 31)
    else:
        # last day of previous month
        prev = date(event_date.year, event_date.month, 1) - __import__("datetime").timedelta(days=1)
    return sbi_usd(prev)


def to_usd(amount: float, currency: str, fx_map: dict | None = None) -> float:
    fx_map = fx_map or FX_TO_USD_YE2025
    return amount * fx_map.get(currency, 1.0)


def acct_info(sections):
    return {r[0]: r[1] for _, r in sections["Account Information"] if len(r) >= 2}


def fin_info(sections):
    out = {}
    for _, r in sections["Financial Instrument Information"]:
        if len(r) < 3:
            continue
        sym = r[1].split(",")[0].strip()  # handle "W1TB, WCBR"
        out[sym] = {
            "category": r[0],
            "symbol": sym,
            "description": r[2],
            "security_id": r[4] if len(r) > 4 else "",
            "exchange": r[6] if len(r) > 6 else "",
            "type": r[8] if len(r) > 8 else "",
        }
        # also map aliases
        if "," in r[1]:
            for a in r[1].split(","):
                out[a.strip()] = out[sym]
    return out


@dataclass
class Lot:
    qty: float
    cost_local: float  # total cost in trade currency
    currency: str
    acq_date: date
    source: str = ""


@dataclass
class SaleMatch:
    symbol: str
    currency: str
    sell_date: date
    qty: float
    proceeds_local: float
    comm_usd: float
    cost_local: float
    acq_date: date
    holding_days: int
    is_ltcg: bool
    realized_local: float
    code: str = ""


def extract_orders(sections):
    """Return list of stock Order rows as dicts."""
    orders = []
    for kind, r in sections["Trades"]:
        if kind != "Data" or r[0] != "Order":
            continue
        if r[1] != "Stocks":
            continue
        qty = fnum(r[5])
        orders.append(
            {
                "currency": r[2],
                "symbol": normalize_symbol(r[3]),
                "datetime": normalize_datetime(r[4]),
                "date": parse_dt(r[4]),
                "qty": qty,
                "price": fnum(r[6]),
                "proceeds": fnum(r[8]),
                "comm_usd": fnum(r[9]) or 0.0,
                "basis": fnum(r[10]) if len(r) > 10 else None,
                "realized": fnum(r[11]) if len(r) > 11 else None,
                "code": r[13] if len(r) > 13 else "",
            }
        )
    return orders


def extract_mtm(sections):
    out = {}
    for kind, r in sections["Mark-to-Market Performance Summary"]:
        if kind != "Data" or r[0] != "Stocks":
            continue
        sym = r[1]
        out[sym] = {
            "prior_qty": fnum(r[2]) or 0.0,
            "curr_qty": fnum(r[3]) or 0.0,
            "prior_price": fnum(r[4]),
            "curr_price": fnum(r[5]),
        }
    return out


def extract_open_positions(sections):
    """Year-end open positions (Summary rows) keyed by symbol."""
    out = {}
    for kind, r in sections["Open Positions"]:
        if kind != "Data" or r[0] != "Summary":
            continue
        sym = normalize_symbol(r[3])
        out[sym] = {
            "currency": r[2],
            "qty": fnum(r[4]) or 0.0,
            "cost_price": fnum(r[6]),
            "cost_basis": fnum(r[7]),
            "close_price": fnum(r[8]),
            "value": fnum(r[9]),
            "upl": fnum(r[10]),
        }
    return out


def extract_dividends(sections):
    rows = []
    for kind, r in sections["Dividends"]:
        if kind != "Data":
            continue
        if r[0] in ("Total",) or str(r[0]).startswith("Total"):
            continue
        if len(r) < 4 or not r[1]:
            continue
        m = re.match(r"^([A-Z0-9.\-]+)(?:\(|$)", r[2].replace(" ", ""))
        # Symbol is before '(' in description like AAPL(US037...)
        m = re.match(r"^([A-Za-z0-9.\-]+)\(", r[2])
        sym = m.group(1) if m else ""
        rows.append(
            {
                "currency": r[0],
                "date": parse_dt(r[1]),
                "description": r[2],
                "amount": fnum(r[3]) or 0.0,
                "symbol": sym,
            }
        )
    return rows


def extract_interest(sections):
    rows = []
    for kind, r in sections["Interest"]:
        if kind != "Data":
            continue
        if r[0] in ("Total",) or str(r[0]).startswith("Total"):
            continue
        if len(r) < 4 or not r[1]:
            continue
        rows.append(
            {
                "currency": r[0],
                "date": parse_dt(r[1]),
                "description": r[2],
                "amount": fnum(r[3]) or 0.0,
            }
        )
    return rows


def extract_withholding(sections):
    rows = []
    for kind, r in sections["Withholding Tax"]:
        if kind != "Data":
            continue
        if r[0] in ("Total",) or str(r[0]).startswith("Total"):
            continue
        if len(r) < 4 or not r[1]:
            continue
        rows.append(
            {
                "currency": r[0],
                "date": parse_dt(r[1]),
                "description": r[2],
                "amount": fnum(r[3]) or 0.0,  # negative = tax withheld
            }
        )
    return rows


def normalize_symbol(sym: str) -> str:
    if sym == "W1TB":
        return "WCBR"
    if sym == "SHELL":
        return "R6C0d"  # IBKR alias for Shell plc on AEB
    if sym == "AIR":
        return "AIRd"
    if sym == "SIE":
        return "SIEd"
    if sym == "ZEAL":
        return "ZEALc"
    # Lux UCITS short codes on transfers vs full ISIN on open positions
    if sym in ("009957456", "LU0099574567"):
        return "LU0099574567"
    if sym in ("010683190", "LU0106831901"):
        return "LU0106831901"
    return sym


def extract_transfers(sections):
    """Stock FOP / ACATS transfers (In = acquisition into IBKR)."""
    out = []
    for kind, r in sections.get("Transfers", []):
        if kind != "Data" or not r or r[0] in ("Total",):
            continue
        # Stocks, Funds, Mutual Funds — all Schedule FA equity/fund interests
        if r[0] not in ("Stocks", "Funds", "Mutual Funds"):
            continue
        direction = (r[5] if len(r) > 5 else "").strip()
        qty = fnum(r[8]) if len(r) > 8 else None
        if qty is None:
            continue
        out.append(
            {
                "currency": r[1],
                "symbol": normalize_symbol(r[2]),
                "date": parse_dt(r[3]),
                "type": r[4] if len(r) > 4 else "",
                "direction": direction,
                "qty": qty,
                "market_value": fnum(r[10]) if len(r) > 10 else None,
                "realized": fnum(r[11]) if len(r) > 11 else None,
            }
        )
    return out


def transfers_to_buy_orders(transfers, open_positions_by_statement):
    """
    Convert inbound stock transfers into synthetic buy orders for FIFO.
    Cost basis prefers IBKR Open Positions cost_basis (carryover); else transfer MV.
    """
    basis_by_sym = {}
    for op in open_positions_by_statement:
        for sym, info in op.items():
            if info.get("cost_basis") is not None and info.get("qty"):
                basis_by_sym[sym] = (info["cost_basis"], info["qty"], info.get("currency") or "USD")

    orders = []
    for t in transfers:
        if t["direction"].lower() != "in":
            continue
        sym = t["symbol"]
        qty = t["qty"]
        if qty <= 0:
            continue
        if sym in basis_by_sym:
            total_basis, basis_qty, ccy = basis_by_sym[sym]
            cost = abs(total_basis) * (qty / basis_qty) if basis_qty else abs(total_basis)
            currency = ccy
        else:
            cost = abs(t["market_value"] or 0.0)
            currency = t["currency"]
        orders.append(
            {
                "currency": currency,
                "symbol": sym,
                "datetime": f"{t['date'].isoformat()}, 00:00:00",
                "date": t["date"],
                "qty": qty,
                "price": (cost / qty) if qty else 0.0,
                "proceeds": -cost,
                "comm_usd": 0.0,
                "basis": cost,
                "realized": 0.0,
                "code": "T",
            }
        )
    return orders


def transfers_to_out_orders(transfers):
    """
    Outbound stock transfers (Internal/FOP Out) as synthetic sells so FIFO lots leave the account.
    Tagged code='I' — excluded from taxable Capital Gains (same beneficial owner).
    """
    orders = []
    for t in transfers:
        if t["direction"].lower() != "out":
            continue
        qty = abs(t["qty"])
        if qty <= 0:
            continue
        proceeds = abs(t["market_value"] or 0.0)
        orders.append(
            {
                "currency": t["currency"],
                "symbol": t["symbol"],
                "datetime": f"{t['date'].isoformat()}, 00:00:00",
                "date": t["date"],
                "qty": -qty,
                "price": (proceeds / qty) if qty else 0.0,
                "proceeds": proceeds,
                "comm_usd": 0.0,
                "basis": proceeds,  # placeholder; CG skipped for code I
                "realized": 0.0,
                "code": "I",
            }
        )
    return orders


def extract_stock_splits(sections):
    out = []
    for kind, r in sections.get("Corporate Actions", []):
        if kind != "Data":
            continue
        desc = r[4] if len(r) > 4 else ""
        m = re.search(r"Split\s+(\d+)\s+for\s+(\d+)", desc, re.I)
        if not m:
            continue
        new, old = int(m.group(1)), int(m.group(2))
        sm = re.match(r"^([A-Za-z0-9.\-]+)\(", desc)
        if not sm:
            continue
        out.append({"date": parse_dt(r[2]), "symbol": normalize_symbol(sm.group(1)), "ratio": new / old})
    return out


def extract_cash_mergers(sections, realized_usd_by_sym=None):
    """Cash merger proceeds (USD leg) with cost inferred from IBKR realized P/L when available."""
    realized_usd_by_sym = realized_usd_by_sym or {}
    out = []
    for kind, r in sections.get("Corporate Actions", []):
        if kind != "Data" or r[1] != "USD":
            continue
        proceeds = fnum(r[6])
        if not proceeds:
            continue
        desc = r[4]
        sm = re.match(r"^([A-Za-z0-9.\-]+)\(", desc)
        if not sm:
            continue
        sym = normalize_symbol(sm.group(1))
        realized = realized_usd_by_sym.get(sym)
        cost = (proceeds - realized) if realized is not None else None
        out.append(
            {
                "date": parse_dt(r[2]),
                "symbol": sym,
                "proceeds_usd": proceeds,
                "cost_usd": cost,
                "description": desc,
            }
        )
    return out


def extract_realized_usd(sections):
    out = {}
    for kind, r in sections.get("Realized & Unrealized Performance Summary", []):
        if kind != "Data" or not r[0] or r[0] in ("Forex",):
            continue
        if r[0] != "Stocks":
            continue
        sym = r[1]
        try:
            out[sym] = float(str(r[7]).replace(",", "") or 0)
        except ValueError:
            pass
    return out


def build_fifo_sales(opening_lots: dict, orders: list, start: date, end: date, splits=None, cash_mergers=None):
    """
    FIFO match sells in [start, end] for acquisition dates / holding period.
    Cost/proceeds from IBKR Order Basis/Proceeds (pro-rata by lot). Splits adjust qty.
    Cash mergers recorded as redemptions using IBKR proceeds/cost.
    """
    books = {k: deque(deepcopy(v)) for k, v in opening_lots.items()}
    sales = []
    splits = splits or []
    cash_mergers = cash_mergers or []

    events = []
    for o in orders:
        events.append(("trade", o["date"], o["datetime"], o))
    for sp in splits:
        events.append(("split", sp["date"], "00:00:00", sp))
    for cm in cash_mergers:
        events.append(("merger", cm["date"], "23:59:59", cm))
    events.sort(key=lambda e: (e[1], e[2], 0 if e[0] == "split" else 1 if e[0] == "trade" else 2))

    for etype, edate, _etime, payload in events:
        if edate < start:
            if etype == "trade":
                o = payload
                sym = normalize_symbol(o["symbol"])
                if o["qty"] > 0:
                    cost = abs(o["basis"]) if o["basis"] is not None else abs(o["proceeds"])
                    books.setdefault(sym, deque()).append(
                        Lot(o["qty"], cost, o["currency"], o["date"], "pre-window buy")
                    )
                elif o["qty"] < 0:
                    rem = abs(o["qty"])
                    while rem > 1e-10 and books.get(sym):
                        lot = books[sym][0]
                        take = min(lot.qty, rem)
                        lot.qty -= take
                        rem -= take
                        if lot.qty <= 1e-10:
                            books[sym].popleft()
            elif etype == "merger":
                books[payload["symbol"]] = deque()
            # Pre-window splits are NOT reapplied: opening_lots are already as-of start
            # (from build_lots_as_of, which applied historical splits).
            continue
        if edate > end:
            continue

        if etype == "split":
            for lot in books.get(payload["symbol"], []):
                lot.qty *= payload["ratio"]
            continue

        if etype == "merger":
            sym = payload["symbol"]
            rem_lots = list(books.get(sym, []))
            qty = sum(L.qty for L in rem_lots) or 1.0
            acq = rem_lots[0].acq_date if rem_lots else date(2024, 7, 1)
            hold = (edate - acq).days
            cost = payload["cost_usd"] if payload["cost_usd"] is not None else 0.0
            sales.append(
                SaleMatch(
                    symbol=sym,
                    currency="USD",
                    sell_date=edate,
                    qty=qty,
                    proceeds_local=payload["proceeds_usd"],
                    comm_usd=0.0,
                    cost_local=cost,
                    acq_date=acq,
                    holding_days=hold,
                    is_ltcg=hold > 730,
                    realized_local=payload["proceeds_usd"] - cost,
                )
            )
            books[sym] = deque()
            continue

        o = payload
        sym = normalize_symbol(o["symbol"])
        if o["qty"] > 0:
            cost = abs(o["basis"]) if o["basis"] is not None else abs(o["proceeds"])
            books.setdefault(sym, deque()).append(
                Lot(o["qty"], cost, o["currency"], o["date"], "buy")
            )
            continue

        sell_qty = abs(o["qty"])
        proceeds_total = abs(o["proceeds"] or 0)
        basis_total = abs(o["basis"]) if o["basis"] is not None else proceeds_total
        rem = sell_qty
        while rem > 1e-10:
            if not books.get(sym):
                books.setdefault(sym, deque()).append(
                    Lot(rem, basis_total * (rem / sell_qty), o["currency"], date(2024, 1, 1), "missing-lot-assumed")
                )
            lot = books[sym][0]
            take = min(lot.qty, rem)
            frac = take / sell_qty
            hold = (o["date"] - lot.acq_date).days
            sales.append(
                SaleMatch(
                    symbol=sym,
                    currency=o["currency"],
                    sell_date=o["date"],
                    qty=take,
                    proceeds_local=proceeds_total * frac,
                    comm_usd=o["comm_usd"] * frac,
                    cost_local=basis_total * frac,
                    acq_date=lot.acq_date,
                    holding_days=hold,
                    is_ltcg=hold > 730,
                    realized_local=proceeds_total * frac - basis_total * frac,
                    code=o.get("code") or "",
                )
            )
            lot.qty -= take
            rem -= take
            if lot.qty <= 1e-10:
                books[sym].popleft()
    return sales, books


def normalize_datetime(s: str) -> str:
    """Normalize 'YYYY-MM-DD, H:MM:SS' vs 'YYYY-MM-DD, HH:MM:SS' for dedupe."""
    try:
        part = s.strip()
        if "," in part:
            d, t = part.split(",", 1)
            bits = t.strip().split(":")
            if bits and len(bits[0]) == 1:
                bits[0] = bits[0].zfill(2)
            return f"{d.strip()}, {':'.join(bits)}"
        return part
    except Exception:
        return s


def merge_orders(*order_lists):
    """Deduplicate stock orders across statements by symbol/datetime/qty/proceeds."""
    seen = {}
    for orders in order_lists:
        for o in orders:
            o = dict(o)
            o["datetime"] = normalize_datetime(o["datetime"])
            key = (o["symbol"], o["datetime"], o["qty"], round(o["proceeds"] or 0, 6), o["currency"])
            seen[key] = o
    return list(seen.values())


def merge_splits(*split_lists):
    seen = {}
    for splits in split_lists:
        for s in splits:
            key = (s["symbol"], s["date"], s["ratio"])
            seen[key] = s
    return list(seen.values())


def build_lots_as_of(orders, splits, as_of: date):
    """
    Replay all buys/sells/splits up to and including as_of, returning FIFO lot books.
    Cost on lots = IBKR trade Basis (actual cost). Acquisition date = trade date.
    """
    books = defaultdict(deque)
    events = []
    for o in orders:
        if o["date"] <= as_of:
            events.append(("trade", o["date"], o["datetime"], o))
    for sp in splits:
        if sp["date"] <= as_of:
            events.append(("split", sp["date"], "00:00:00", sp))
    events.sort(key=lambda e: (e[1], e[2], 0 if e[0] == "split" else 1))

    for etype, edate, _etime, payload in events:
        if etype == "split":
            for lot in books.get(payload["symbol"], []):
                lot.qty *= payload["ratio"]
                # cost unchanged on split
            continue
        o = payload
        sym = normalize_symbol(o["symbol"])
        if o["qty"] > 0:
            cost = abs(o["basis"]) if o["basis"] is not None else abs(o["proceeds"])
            books[sym].append(Lot(o["qty"], cost, o["currency"], o["date"], "buy"))
        else:
            rem = abs(o["qty"])
            while rem > 1e-10 and books.get(sym):
                lot = books[sym][0]
                take = min(lot.qty, rem)
                if lot.qty:
                    lot.cost_local *= (lot.qty - take) / lot.qty
                lot.qty -= take
                rem -= take
                if lot.qty <= 1e-10:
                    books[sym].popleft()
            if rem > 1e-6:
                # oversell without lots — should not happen with complete history
                pass
    # drop empty
    return {k: v for k, v in books.items() if v and sum(L.qty for L in v) > 1e-10}


def earliest_acq(books, sym):
    lots = books.get(sym)
    if not lots:
        return None
    return min(L.acq_date for L in lots)


def opening_lots_from_mtm(mtm, open_ye, orders_before):
    """Deprecated fallback — prefer build_lots_as_of from inception trades."""
    return build_lots_as_of(orders_before, [], date(2024, 12, 31))


def entity_row(sym, fin):
    meta = ENTITY.get(sym)
    if meta:
        ckey, nature, addr, zipc, name = meta
    else:
        info = fin.get(sym, {})
        ckey = "United States of America"
        nature = "Listed Foreign Equity / ETF"
        addr = ""
        zipc = ""
        name = info.get("description", sym)
    cname, ccode = COUNTRY[ckey]
    return cname, ccode, name, addr, zipc, nature


def style_header(ws, row, cols):
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(color="FFFFFF", bold=True, size=10)
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def autosize(ws, max_width=48):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = 10
        for cell in col[:80]:
            if cell.value is not None:
                width = min(max_width, max(width, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = width


def main(
    inception: Path | None = None,
    annual: Path | None = None,
    fiscal: Path | None = None,
    out: Path | None = None,
    account_open_date: str | None = None,
    notes_extra: list | None = None,
    assessee_label: str | None = None,
    seed_orders: list | None = None,
):
    inception = inception or INCEPTION
    annual = annual or ANNUAL
    fiscal = fiscal or FISCAL
    out = out or OUT

    hA, sA = parse_ibkr(annual)
    hF, sF = parse_ibkr(fiscal)
    hI, sI = parse_ibkr(inception)
    acct = acct_info(sA) or acct_info(sI)
    fin = fin_info(sA)
    fin.update(fin_info(sF))
    fin.update(fin_info(sI))

    mtm = extract_mtm(sA)
    open_ye = extract_open_positions(sA)
    open_inc = extract_open_positions(sI)
    open_fy = extract_open_positions(sF)
    orders_cy = extract_orders(sA)
    orders_fy_all = extract_orders(sF)
    orders_inc = extract_orders(sI)
    all_xfers = extract_transfers(sI) + extract_transfers(sA) + extract_transfers(sF)
    xfer_orders = transfers_to_buy_orders(all_xfers, [open_inc, open_ye, open_fy])
    xfer_out_orders = transfers_to_out_orders(all_xfers)
    orders_merged = merge_orders(
        orders_inc, orders_cy, orders_fy_all, xfer_orders, xfer_out_orders, seed_orders or []
    )

    splits = merge_splits(
        extract_stock_splits(sI),
        extract_stock_splits(sA),
        extract_stock_splits(sF),
    )

    # True FIFO opening books from inception trades
    lots_ye2024 = build_lots_as_of(orders_merged, splits, date(2024, 12, 31))  # as of 31-Dec-2024 / 01-Jan-2025
    lots_fy_start = build_lots_as_of(orders_merged, splits, date(2025, 3, 31))  # as of 31-Mar-2025 / 01-Apr-2025
    # first acquisition date per symbol (ever)
    first_acq = {}
    for o in sorted(orders_merged, key=lambda x: (x["date"], x["datetime"])):
        if o["qty"] <= 0:
            continue
        sym = normalize_symbol(o["symbol"])
        first_acq.setdefault(sym, o["date"])

    div_cy = extract_dividends(sA)
    div_fy_all = extract_dividends(sF)
    # Prefer annual for overlap accuracy
    div_fy_map = {(d["date"], d["description"], d["amount"]): d for d in div_cy if d["date"] >= date(2025, 4, 1)}
    for d in div_fy_all:
        if date(2025, 4, 1) <= d["date"] <= date(2026, 3, 31):
            div_fy_map[(d["date"], d["description"], d["amount"])] = d
    div_fy = sorted(div_fy_map.values(), key=lambda x: (x["date"], x["description"]))

    # FY2024-25 dividends (for completeness)
    div_fy2425_map = {}
    for d in extract_dividends(sI):
        if date(2024, 4, 1) <= d["date"] <= date(2025, 3, 31):
            div_fy2425_map[(d["date"], d["description"], d["amount"])] = d
    for d in div_cy:
        if date(2024, 4, 1) <= d["date"] <= date(2025, 3, 31):
            div_fy2425_map[(d["date"], d["description"], d["amount"])] = d
    div_fy2425 = sorted(div_fy2425_map.values(), key=lambda x: (x["date"], x["description"]))

    int_cy = extract_interest(sA)
    int_fy_all = extract_interest(sF)
    int_fy_map = {(d["date"], d["description"], d["amount"]): d for d in int_cy if d["date"] >= date(2025, 4, 1)}
    for d in int_fy_all:
        if date(2025, 4, 1) <= d["date"] <= date(2026, 3, 31):
            int_fy_map[(d["date"], d["description"], d["amount"])] = d
    int_fy = sorted(int_fy_map.values(), key=lambda x: (x["date"], x["description"]))

    int_fy2425 = [
        d for d in extract_interest(sI)
        if date(2024, 4, 1) <= d["date"] <= date(2025, 3, 31)
    ]

    wht_cy = extract_withholding(sA)
    wht_fy_all = extract_withholding(sF)
    wht_fy = [w for w in wht_fy_all if date(2025, 4, 1) <= w["date"] <= date(2026, 3, 31)]
    wht_fy2425 = [
        w for w in extract_withholding(sI)
        if date(2024, 4, 1) <= w["date"] <= date(2025, 3, 31)
    ]

    nav = {r[0]: fnum(r[1]) for _, r in sA.get("Change in NAV", [])}
    deposits = []
    for kind, r in sA.get("Deposits & Withdrawals", []):
        if kind != "Data" or not r or r[0] in ("Total",) or str(r[0]).startswith("Total"):
            continue
        if len(r) < 4 or not r[1]:
            continue
        deposits.append({"currency": r[0], "date": parse_dt(r[1]), "desc": r[2], "amount": fnum(r[3])})
    deposits_inc = []
    for kind, r in sI.get("Deposits & Withdrawals", []):
        if kind != "Data" or not r or r[0] in ("Total",) or str(r[0]).startswith("Total"):
            continue
        if len(r) < 4 or not r[1]:
            continue
        deposits_inc.append({"currency": r[0], "date": parse_dt(r[1]), "desc": r[2], "amount": fnum(r[3])})
    # Infer account open date: first deposit, else first inbound transfer, else override
    open_date_note = account_open_date
    if not open_date_note:
        first_dates = []
        if deposits_inc:
            first_dates.append(min(d["date"] for d in deposits_inc))
        if deposits:
            first_dates.append(min(d["date"] for d in deposits))
        xfers_in = [t for t in extract_transfers(sI) + extract_transfers(sA) if t["direction"].lower() == "in"]
        if xfers_in:
            first_dates.append(min(t["date"] for t in xfers_in))
        if first_dates:
            open_date_note = min(first_dates).isoformat() + " (first funding / inbound transfer per IBKR)"
        else:
            open_date_note = "Confirm from IBKR account opening documents"

    base_ccy = (acct.get("Base Currency") or "USD").strip().upper()

    def _amt_usd(amount: float, currency: str) -> float:
        if (currency or "USD").upper() == "USD":
            return amount
        return to_usd(amount, currency)

    def _deposits_usd(deps: list) -> float:
        return sum(_amt_usd(d["amount"], d["currency"]) for d in deps)

    # Opening lots for FA/CG dating come from inception replay
    opening = lots_ye2024

    # ---------------- Schedule FA A3: FIFO lot-wise (aligns with Capital Gains) ----------------
    # Replay CY2025 trades so each acquisition date is a separate A3 line; INR uses same
    # Rule 115 / EUR methodology as the Capital Gains sheets.

    def fa_amount_inr(local_amt, currency, event_date, for_sale=True):
        """Convert FCY amount to INR consistently with CG sheet."""
        if currency == "EUR":
            prev = (
                date(event_date.year, event_date.month, 1) - timedelta(days=1)
                if event_date.month > 1
                else date(event_date.year - 1, 12, 31)
            )
            rate = SBI_TT_EUR[month_end_on_or_before(prev, SBI_TT_EUR)]
            return local_amt * rate, rate, local_amt, 1.0
        if currency == "USD":
            rate = rule115_usd(event_date) if for_sale else sbi_usd(event_date)
            # For Schedule FA initial of assets acquired in year, SBI TT on acq date is used;
            # for income/sale Rule 115. Peak/closing use event-date SBI.
            if not for_sale:
                rate = sbi_usd(event_date)
            return local_amt * rate, rate, local_amt, 1.0
        # HKD/JPY etc: local → USD (IBKR YE) → INR
        usd = to_usd(local_amt, currency)
        rate = rule115_usd(event_date) if for_sale else sbi_usd(event_date)
        if not for_sale:
            rate = sbi_usd(event_date)
        return usd * rate, rate, usd, FX_TO_USD_YE2025.get(currency, 1.0)

    rate_close = sbi_usd(date(2025, 12, 31))
    rate_peak = sbi_usd(date(2025, 12, 31))
    rate_init_prior = sbi_usd(date(2024, 12, 31))

    # CY2025 FIFO: opening = lots at 31-Dec-2024
    lots_cy_open = {k: deque(deepcopy(list(v))) for k, v in lots_ye2024.items()}
    cy_orders = [o for o in orders_merged if date(2025, 1, 1) <= o["date"] <= date(2025, 12, 31)]
    sales_cy, lots_cy_end = build_fifo_sales(
        lots_cy_open,
        cy_orders,
        date(2025, 1, 1),
        date(2025, 12, 31),
        splits=splits,
        cash_mergers=[m for m in extract_cash_mergers(sA, extract_realized_usd(sA))
                      if date(2025, 1, 1) <= m["date"] <= date(2025, 12, 31)],
    )

    # Dividends by symbol (CY)
    div_inr_by_sym = defaultdict(float)
    div_usd_by_sym = defaultdict(float)
    div_local_by_sym = defaultdict(float)
    for d in div_cy:
        sym = normalize_symbol(d["symbol"]) if d["symbol"] else "UNKNOWN"
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        div_usd_by_sym[sym] += usd
        div_local_by_sym[sym] += d["amount"]
        div_inr_by_sym[sym] += usd * rule115_usd(d["date"])

    # Close prices / currency at YE2025 for remaining lots
    close_info = {}
    for sym, pos in open_ye.items():
        close_info[sym] = {
            "currency": pos["currency"],
            "close_price": pos["close_price"],
            "qty": pos["qty"],
            "value": pos["value"],
        }
    # also from MTM curr
    for sym, m in mtm.items():
        if sym not in close_info and (m.get("curr_qty") or 0) and m.get("curr_price") is not None:
            # currency from trades
            ccy = "USD"
            for o in orders_merged:
                if normalize_symbol(o["symbol"]) == sym:
                    ccy = o["currency"]
                    break
            close_info[sym] = {
                "currency": ccy,
                "close_price": m["curr_price"],
                "qty": m["curr_qty"],
                "value": m["curr_qty"] * m["curr_price"],
            }

    # Jan-1 FMV per symbol (for opening lots' Schedule FA initial)
    jan1_fmv = {}
    for sym, m in mtm.items():
        pq, pp = m.get("prior_qty") or 0, m.get("prior_price")
        if pq and pp is not None:
            ccy = close_info.get(sym, {}).get("currency")
            if not ccy:
                for o in orders_merged:
                    if normalize_symbol(o["symbol"]) == sym:
                        ccy = o["currency"]
                        break
                ccy = ccy or "USD"
            jan1_fmv[sym] = {"qty": pq, "price": pp, "value": pq * pp, "currency": ccy}

    fa_rows = []
    skip_syms = {"SHELL.DRS", "SHELL.DVD", "SHELL.DVR", "SHELL1.DI", "SHELL.DDR"}
    # Track which symbols already got dividend attributed (put on first lot row)
    div_assigned = set()

    def append_fa_lot(sym, acq_date, qty, currency, initial_local, sale_local, closing_local,
                      sale_date=None, note=""):
        if sym in skip_syms:
            return
        if abs(qty) < 1e-9 and (initial_local or 0) == 0 and (sale_local or 0) == 0 and (closing_local or 0) == 0:
            return
        cname, ccode, name, addr, zipc, nature = entity_row(sym, fin)
        # Initial / cost INR
        # - Sold lots: use same Rule 115 / EUR methodology as Capital Gains (so A3 ↔ CG reconcile)
        # - Opening lots still held: Schedule FA beginning value at 31-Dec-2024 SBI TT
        # - Acquired in CY and still held: SBI TT on acquisition date
        if sale_local and sale_date:
            # Align with CG cost conversion
            if currency == "EUR":
                prev = (
                    date(acq_date.year, acq_date.month, 1) - timedelta(days=1)
                    if acq_date.month > 1
                    else date(acq_date.year - 1, 12, 31)
                )
                r_init = SBI_TT_EUR[month_end_on_or_before(prev, SBI_TT_EUR)]
                init_inr = initial_local * r_init
                init_usd = to_usd(initial_local, "EUR")
            elif currency == "USD":
                r_init = rule115_usd(acq_date)
                init_usd = initial_local
                init_inr = initial_local * r_init
            else:
                init_usd = to_usd(initial_local, currency)
                r_init = rule115_usd(acq_date)
                init_inr = init_usd * r_init
        elif acq_date < date(2025, 1, 1) and sym in jan1_fmv and jan1_fmv[sym]["qty"]:
            unit = jan1_fmv[sym]["value"] / jan1_fmv[sym]["qty"]
            init_local_use = unit * qty
            ccy0 = jan1_fmv[sym]["currency"]
            init_usd = init_local_use if ccy0 == "USD" else to_usd(init_local_use, ccy0)
            r_init = rate_init_prior
            init_inr = init_usd * r_init
        else:
            if currency == "EUR":
                r_init = SBI_TT_EUR[month_end_on_or_before(acq_date, SBI_TT_EUR)]
                init_inr = initial_local * r_init
                init_usd = to_usd(initial_local, "EUR")
            elif currency == "USD":
                r_init = sbi_usd(acq_date)
                init_usd = initial_local
                init_inr = initial_local * r_init
            else:
                init_usd = to_usd(initial_local, currency)
                r_init = sbi_usd(acq_date)
                init_inr = init_usd * r_init

        # Sale INR — match CG: Rule 115 / EUR Rule 115
        if sale_local and sale_date:
            if currency == "EUR":
                prev = (
                    date(sale_date.year, sale_date.month, 1) - timedelta(days=1)
                    if sale_date.month > 1
                    else date(sale_date.year - 1, 12, 31)
                )
                r_sale = SBI_TT_EUR[month_end_on_or_before(prev, SBI_TT_EUR)]
                sale_inr = sale_local * r_sale
                sale_usd = to_usd(sale_local, "EUR")
            elif currency == "USD":
                r_sale = rule115_usd(sale_date)
                sale_usd = sale_local
                sale_inr = sale_local * r_sale
            else:
                sale_usd = to_usd(sale_local, currency)
                r_sale = rule115_usd(sale_date)
                sale_inr = sale_usd * r_sale
        else:
            sale_usd = sale_inr = 0.0

        # Closing
        if closing_local:
            if currency == "EUR":
                close_usd = to_usd(closing_local, "EUR")
            elif currency == "USD":
                close_usd = closing_local
            else:
                close_usd = to_usd(closing_local, currency)
            close_inr = close_usd * rate_close
        else:
            close_usd = close_inr = 0.0

        peak_usd = max(init_usd, close_usd, sale_usd)
        peak_inr = peak_usd * rate_peak

        # Dividends once per symbol on first lot row
        if sym not in div_assigned:
            d_usd = div_usd_by_sym.get(sym, 0.0)
            d_inr = div_inr_by_sym.get(sym, 0.0)
            div_assigned.add(sym)
        else:
            d_usd = d_inr = 0.0

        fa_rows.append(
            {
                "country": cname,
                "code": ccode,
                "name": sym,  # A3 Col C — Symbol only
                "address": addr,
                "zip": zipc,
                "nature": nature,
                "acq": acq_date.isoformat(),
                "acq_date": acq_date,
                "initial_usd": round(init_usd, 2),
                "peak_usd": round(peak_usd, 2),
                "closing_usd": round(close_usd, 2),
                "div_usd": round(d_usd, 2),
                "sale_usd": round(sale_usd, 2),
                "initial_inr": round(init_inr),
                "peak_inr": round(peak_inr),
                "closing_inr": round(close_inr),
                "div_inr": round(d_inr),
                "sale_inr": round(sale_inr),
                "symbol": sym,
                "qty": qty,
                "r_init": r_init,
                "r_peak": rate_peak,
                "r_close": rate_close,
                "lot_note": note,
                "initial_local": round(initial_local, 4),
                "sale_local": round(sale_local or 0, 4),
                "currency": currency,
            }
        )

    # 1) Sold lots in CY2025 — one A3 row per FIFO lot (matches Capital Gains lot lines)
    for s in sales_cy:
        append_fa_lot(
            s.symbol,
            s.acq_date,
            s.qty,
            s.currency,
            initial_local=s.cost_local,
            sale_local=s.proceeds_local,
            closing_local=0.0,
            sale_date=s.sell_date,
            note=f"FIFO lot sold {s.sell_date.isoformat()}; aligns with Capital Gains",
        )

    # 2) Remaining lots at 31-Dec-2025
    for sym, lots in sorted(lots_cy_end.items()):
        info = close_info.get(sym, {})
        ccy = info.get("currency") or (lots[0].currency if lots else "USD")
        px = info.get("close_price")
        for L in lots:
            if L.qty <= 1e-10:
                continue
            closing_local = (px * L.qty) if px is not None else 0.0
            append_fa_lot(
                sym,
                L.acq_date,
                L.qty,
                L.currency or ccy,
                initial_local=L.cost_local,
                sale_local=0.0,
                closing_local=closing_local,
                sale_date=None,
                note="Holding at 31-Dec-2025",
            )

    # 3) Symbols with dividends but no lots/sales caught (edge)
    for sym, d_usd in div_usd_by_sym.items():
        if sym in div_assigned or sym in skip_syms:
            continue
        cname, ccode, name, addr, zipc, nature = entity_row(sym, fin)
        fa_rows.append(
            {
                "country": cname,
                "code": ccode,
                "name": sym,  # A3 Col C — Symbol only
                "address": addr,
                "zip": zipc,
                "nature": nature,
                "acq": first_acq.get(sym, date(2025, 1, 1)).isoformat(),
                "acq_date": first_acq.get(sym, date(2025, 1, 1)),
                "initial_usd": 0,
                "peak_usd": 0,
                "closing_usd": 0,
                "div_usd": round(d_usd, 2),
                "sale_usd": 0,
                "initial_inr": 0,
                "peak_inr": 0,
                "closing_inr": 0,
                "div_inr": round(div_inr_by_sym.get(sym, 0)),
                "sale_inr": 0,
                "symbol": sym,
                "r_init": rate_init_prior,
                "r_peak": rate_peak,
                "r_close": rate_close,
            }
        )
        div_assigned.add(sym)

    fa_rows.sort(key=lambda r: (r["symbol"], r["acq_date"], r.get("sale_local", 0)))

    # ---------------- Capital gains FY 2025-26 (FIFO from inception) ----------------
    realized_map = extract_realized_usd(sA)
    realized_map.update(extract_realized_usd(sF))
    realized_map.update(extract_realized_usd(sI))
    mergers = extract_cash_mergers(sA, realized_map)
    # Also mergers in inception period (none expected with cash) — include if any
    mergers += [m for m in extract_cash_mergers(sI, realized_map)
                if not any(x["symbol"] == m["symbol"] and x["date"] == m["date"] for x in mergers)]

    # Opening books as of 01-Apr-2025 = lots after replaying through 31-Mar-2025
    lots_apr = {k: deque(deepcopy(list(v))) for k, v in lots_fy_start.items()}

    fy_orders = [o for o in orders_merged if date(2025, 4, 1) <= o["date"] <= date(2026, 3, 31)]
    sales_fy, _ = build_fifo_sales(
        lots_apr,
        fy_orders,
        date(2025, 4, 1),
        date(2026, 3, 31),
        splits=splits,
        cash_mergers=[m for m in mergers if date(2025, 4, 1) <= m["date"] <= date(2026, 3, 31)],
    )

    # FY 2024-25 capital gains (AY 2025-26) — BLBD/TSLA etc.
    lots_apr2425 = defaultdict(deque)  # account started from zero in FY24-25
    fy2425_orders = [o for o in orders_merged if date(2024, 4, 1) <= o["date"] <= date(2025, 3, 31)]
    sales_fy2425, _ = build_fifo_sales(
        lots_apr2425,
        fy2425_orders,
        date(2024, 4, 1),
        date(2025, 3, 31),
        splits=splits,
        cash_mergers=[m for m in mergers if date(2024, 4, 1) <= m["date"] <= date(2025, 3, 31)],
    )

    def sales_to_cg_rows(sales):
        rows = []
        skip_cg = {"SHELL.DRS", "SHELL.DVD", "SHELL.DVR", "SHELL1.DI", "SHELL.DDR"}
        for s in sales:
            if s.symbol in skip_cg:
                continue
            # Internal / inter-account transfers — not taxable CG (same beneficial owner)
            if (s.code or "").upper() in ("I",):
                continue
            if abs(s.qty) < 1e-6 or (abs(s.proceeds_local) < 1e-8 and abs(s.cost_local) < 1e-8):
                continue
            if s.currency == "USD":
                proc_usd = s.proceeds_local
                cost_usd = s.cost_local
            elif s.currency == "EUR":
                proc_usd = to_usd(s.proceeds_local, "EUR")
                cost_usd = to_usd(s.cost_local, "EUR")
            else:
                proc_usd = to_usd(s.proceeds_local, s.currency)
                cost_usd = to_usd(s.cost_local, s.currency)

            sale_rate_usd = rule115_usd(s.sell_date)
            cost_rate_usd = rule115_usd(s.acq_date)

            if s.currency == "EUR":
                sale_prev = (
                    date(s.sell_date.year, s.sell_date.month, 1) - timedelta(days=1)
                    if s.sell_date.month > 1
                    else date(s.sell_date.year - 1, 12, 31)
                )
                cost_prev = (
                    date(s.acq_date.year, s.acq_date.month, 1) - timedelta(days=1)
                    if s.acq_date.month > 1
                    else date(s.acq_date.year - 1, 12, 31)
                )
                sale_rate = SBI_TT_EUR[month_end_on_or_before(sale_prev, SBI_TT_EUR)]
                cost_rate = SBI_TT_EUR[month_end_on_or_before(cost_prev, SBI_TT_EUR)]
                sale_inr = s.proceeds_local * sale_rate
                cost_inr = s.cost_local * cost_rate
                comm_inr = abs(s.comm_usd) * sale_rate_usd
            else:
                sale_rate = sale_rate_usd
                cost_rate = cost_rate_usd
                sale_inr = proc_usd * sale_rate_usd
                cost_inr = cost_usd * cost_rate_usd
                comm_inr = abs(s.comm_usd) * sale_rate_usd

            gain_inr = sale_inr - comm_inr - cost_inr
            note = ""
            if s.symbol == "HYU" and s.sell_date == date(2025, 7, 30):
                note = "Cash merger / redemption — cost from IBKR realized P/L"
            elif s.symbol == "LRCX" and s.acq_date <= date(2024, 10, 3):
                note = "LRCX 10-for-1 split on 02/03-Oct-2024 reflected in FIFO qty"
            rows.append(
                {
                    "symbol": s.symbol,
                    "currency": s.currency,
                    "qty": round(s.qty, 6),
                    "acq_date": s.acq_date,
                    "sell_date": s.sell_date,
                    "holding_days": s.holding_days,
                    "type": "LTCG" if s.is_ltcg else "STCG",
                    "proceeds_local": round(s.proceeds_local, 2),
                    "cost_local": round(s.cost_local, 2),
                    "proceeds_usd": round(proc_usd, 2),
                    "cost_usd": round(cost_usd, 2),
                    "comm_usd": round(abs(s.comm_usd), 2),
                    "sale_rate": sale_rate,
                    "cost_rate": cost_rate,
                    "sale_inr": round(sale_inr),
                    "comm_inr": round(comm_inr),
                    "cost_inr": round(cost_inr),
                    "gain_inr": round(gain_inr),
                    "note": note,
                }
            )
        return rows

    cg_rows = sales_to_cg_rows(sales_fy)
    cg_rows_2425 = sales_to_cg_rows(sales_fy2425)
    stcg = sum(r["gain_inr"] for r in cg_rows if r["type"] == "STCG")
    ltcg = sum(r["gain_inr"] for r in cg_rows if r["type"] == "LTCG")
    stcg2425 = sum(r["gain_inr"] for r in cg_rows_2425 if r["type"] == "STCG")
    ltcg2425 = sum(r["gain_inr"] for r in cg_rows_2425 if r["type"] == "LTCG")

    def write_fx_lookup(wb):
        """Editable SBI TT rate table used by VLOOKUP in capital-gain formulas."""
        ws = wb.create_sheet("FX Lookup", 0)
        ws["A1"] = (
            "EDITABLE FX LOOKUP — change rates in column B/C; capital-gain sheets recalculate via VLOOKUP. "
            "Rule 115: use TT Buy on last day of month preceding the month of sale / acquisition."
        )
        ws.merge_cells("A1:E1")
        ws["A3"] = "Rate Date (month-end)"
        ws["B3"] = "USD/INR (SBI TT Buy)"
        ws["C3"] = "EUR/INR (SBI TT Buy)"
        ws["D3"] = "Notes"
        style_header(ws, 3, 4)
        # build sorted unique dates from both tables
        all_dates = sorted(set(SBI_TT_USD) | set(SBI_TT_EUR))
        r = 4
        for d in all_dates:
            ws.cell(r, 1, d)
            ws.cell(r, 2, SBI_TT_USD.get(d))
            ws.cell(r, 3, SBI_TT_EUR.get(d))
            ws.cell(r, 4, "Month-end card — replace with exact SBI card rate if different")
            r += 1
        last = r - 1
        ws["A2"] = f"Table range for VLOOKUP: A4:C{last}"
        # FCY→USD cross rates used for non-USD/non-EUR (year-end IBKR)
        ws.cell(r + 1, 1, "FCY→USD cross rates (IBKR 31-Dec-2025 close — edit if using trade-date FX)")
        ws.cell(r + 2, 1, "Currency")
        ws.cell(r + 2, 2, "Units per 1 FCY → USD")
        style_header(ws, r + 2, 2)
        cross_start = r + 3
        for i, (ccy, rate) in enumerate(sorted(FX_TO_USD_YE2025.items())):
            ws.cell(cross_start + i, 1, ccy)
            ws.cell(cross_start + i, 2, rate)
        ws.cell(cross_start + len(FX_TO_USD_YE2025) + 1, 1,
                "Yellow/input cells on CG sheets: Sale/Cost FCY, Comm USD, FX rates. All P&L columns are formulas.")
        # light fill on rate columns
        fill = PatternFill("solid", fgColor="FFF2CC")
        for row in range(4, last + 1):
            ws.cell(row, 2).fill = fill
            ws.cell(row, 3).fill = fill
        for row in range(cross_start, cross_start + len(FX_TO_USD_YE2025)):
            ws.cell(row, 2).fill = fill
        autosize(ws)
        return last, cross_start

    def rule115_prev_month(d: date) -> date:
        if d.month == 1:
            return date(d.year - 1, 12, 31)
        return date(d.year, d.month, 1) - timedelta(days=1)

    def write_formula_cg_sheet(wb, title, subtitle, rows, fx_last_row, cross_start):
        """
        Formula-driven capital gains for reconciliation.
        Yellow inputs: H Sale FCY, I Cost FCY, J Comm USD, K FCY→USD, N Sale FX, O Cost FX, P Comm USD/INR
        Formulas: F holding days, G type, L/M amounts, Q/R/S INR, T Gain = Q-R-S
        """
        ws = wb.create_sheet(title)
        ws["A1"] = subtitle
        ws.merge_cells("A1:U1")
        ws["A2"] = (
            "RECONCILIATION SHEET (formula-driven). Edit yellow cells — P&L recalculates. "
            "Gain/(Loss) INR = Sale INR − Comm INR − Cost INR. Holding days >730 → LTCG else STCG. "
            "Click any calculated cell to see the formula. FX Lookup has the full SBI TT rate table."
        )
        ws.merge_cells("A2:U2")
        ws["A3"] = (
            "F=E-D | G=IF(F>730,\"LTCG\",\"STCG\") | L=H*K | M=I*K | "
            "Q=L*N (Sale INR) | R=J*P (Comm INR) | S=M*O (Cost INR) | T=Q-R-S (Gain)"
        )
        ws.merge_cells("A3:U3")

        headers = [
            "Symbol", "CCY", "Qty", "Acquisition Date", "Sale Date", "Holding Days", "Type",
            "Sale Proceeds (FCY)", "Cost Basis (FCY)", "Comm (USD)", "FCY→USD Rate",
            "Sale Proceeds (USD)", "Cost (USD)", "Sale FX Rate", "Cost FX Rate", "Comm USD/INR Rate",
            "Sale (INR)", "Comm (INR)", "Cost (INR)", "Gain/(Loss) INR", "Notes",
        ]
        for i, h in enumerate(headers, 1):
            ws.cell(4, i, h)
        style_header(ws, 4, len(headers))

        input_fill = PatternFill("solid", fgColor="FFF2CC")
        first_data = 5
        for i, r in enumerate(rows):
            rr = first_data + i
            usd_sale = rule115_usd(r["sell_date"])
            if r["currency"] == "EUR":
                fcy_to_usd = 1.0
                sale_fx = r["sale_rate"]
                cost_fx = r["cost_rate"]
                sale_fcy = r["proceeds_local"]
                cost_fcy = r["cost_local"]
                note = (r["note"] + " | " if r["note"] else "") + (
                    "EUR: FCY→USD=1; Sale/Cost FX = EUR/INR (Rule 115); Comm uses USD/INR"
                )
            else:
                fcy_to_usd = 1.0 if r["currency"] == "USD" else FX_TO_USD_YE2025.get(r["currency"], 1.0)
                sale_fx = r["sale_rate"] if r["currency"] == "USD" else rule115_usd(r["sell_date"])
                cost_fx = r["cost_rate"] if r["currency"] == "USD" else rule115_usd(r["acq_date"])
                sale_fcy = r["proceeds_local"]
                cost_fcy = r["cost_local"]
                note = r["note"]
                if r["currency"] != "USD":
                    note = (note + " | " if note else "") + (
                        f"{r['currency']}: FCY→USD editable; Sale/Cost FX = USD/INR (Rule 115)"
                    )

            ws.cell(rr, 1, r["symbol"])
            ws.cell(rr, 2, r["currency"])
            ws.cell(rr, 3, r["qty"])
            ws.cell(rr, 4, r["acq_date"])
            ws.cell(rr, 5, r["sell_date"])
            ws.cell(rr, 6, f"=E{rr}-D{rr}")
            ws.cell(rr, 7, f'=IF(F{rr}>730,"LTCG","STCG")')
            ws.cell(rr, 8, round(sale_fcy, 6))
            ws.cell(rr, 9, round(cost_fcy, 6))
            ws.cell(rr, 10, r["comm_usd"])
            ws.cell(rr, 11, fcy_to_usd)
            ws.cell(rr, 12, f"=H{rr}*K{rr}")
            ws.cell(rr, 13, f"=I{rr}*K{rr}")
            ws.cell(rr, 14, sale_fx)
            ws.cell(rr, 15, cost_fx)
            ws.cell(rr, 16, usd_sale)
            ws.cell(rr, 17, f"=L{rr}*N{rr}")
            ws.cell(rr, 18, f"=J{rr}*P{rr}")
            ws.cell(rr, 19, f"=M{rr}*O{rr}")
            ws.cell(rr, 20, f"=Q{rr}-R{rr}-S{rr}")
            ws.cell(rr, 21, note)
            for col in (8, 9, 10, 11, 14, 15, 16):
                ws.cell(rr, col).fill = input_fill

        last_data = first_data + len(rows) - 1 if rows else first_data - 1
        tot = last_data + 2
        if rows:
            ws.cell(tot, 7, "TOTAL STCG (INR)")
            ws.cell(tot, 20, f'=SUMIF(G{first_data}:G{last_data},"STCG",T{first_data}:T{last_data})')
            ws.cell(tot + 1, 7, "TOTAL LTCG (INR)")
            ws.cell(tot + 1, 20, f'=SUMIF(G{first_data}:G{last_data},"LTCG",T{first_data}:T{last_data})')
            ws.cell(tot + 2, 7, "TOTAL CG (INR)")
            ws.cell(tot + 2, 20, f"=T{tot}+T{tot+1}")
        for r in range(tot, tot + 3):
            ws.cell(r, 7).font = Font(bold=True)
            ws.cell(r, 20).font = Font(bold=True)

        ws.cell(tot + 4, 1, "How to reconcile before filing:")
        ws.cell(tot + 5, 1, "1. Check Sale Proceeds (FCY) / Cost Basis (FCY) vs IBKR trade Proceeds & Basis.")
        ws.cell(tot + 6, 1, "2. Edit yellow FX cells (N Sale FX, O Cost FX, P Comm USD/INR) to exact SBI TT Buy card rates.")
        ws.cell(tot + 7, 1, "3. For HKD/JPY, edit FCY→USD (K) if using trade-date FX instead of 31-Dec-2025 IBKR close.")
        ws.cell(tot + 8, 1, "4. Gain/(Loss) cell formula: =Q-R-S i.e. Sale INR − Comm INR − Cost INR.")
        ws.cell(tot + 9, 1, "5. Bottom totals are SUMIF formulas on Type (STCG/LTCG).")
        autosize(ws)
        return ws

    # ---------------- Build workbook ----------------
    wb = Workbook()

    # A3
    ws = wb.active
    ws.title = "A3"
    ws["A1"] = (
        "A3. Details of Foreign Equity and Debt Interest held (including any beneficial interest) "
        "in any entity at any time during the calendar year 2025"
    )
    ws.merge_cells("A1:L1")
    headers = [
        "Country name", "Country code", "Name of entity", "Address of entity", "ZIP code",
        "Nature of entity", "Date of acquiring the interest", "Initial value of the investment",
        "Peak value of investment during the period", "Closing balance",
        "Total gross amount paid/credited with respect to the holding during the period",
        "Total gross proceeds from sale or redemption of investment during the period",
    ]
    for i, h in enumerate(headers, 1):
        ws.cell(2, i, h)
    style_header(ws, 2, 12)
    for i, h in enumerate(range(1, 13), 1):
        ws.cell(4, i, i)
    r0 = 5
    for i, row in enumerate(fa_rows):
        rr = r0 + i
        vals = [
            row["country"], row["code"], row["name"], row["address"], row["zip"], row["nature"],
            row["acq"], row["initial_inr"], row["peak_inr"], row["closing_inr"], row["div_inr"], row["sale_inr"],
        ]
        for c, v in enumerate(vals, 1):
            ws.cell(rr, c, v)
    tot_r = r0 + len(fa_rows) + 1
    ws.cell(tot_r, 7, "TOTAL (Rs.)")
    for col, key in [(8, "initial_inr"), (9, "peak_inr"), (10, "closing_inr"), (11, "div_inr"), (12, "sale_inr")]:
        ws.cell(tot_r, col, sum(r[key] for r in fa_rows))
        ws.cell(tot_r, col).font = Font(bold=True)
    autosize(ws)

    # A3 USD reference
    ws2 = wb.create_sheet("A3 (USD reference)")
    h2 = [
        "S.No", "Country", "Country Code", "Symbol", "Description", "Nature", "Qty Close",
        "Date of Acquisition", "Initial Value (USD)", "Peak Value (USD)", "Closing Value (USD)",
        "Dividends CY2025 (USD)", "Sale Proceeds (USD)",
    ]
    for i, h in enumerate(h2, 1):
        ws2.cell(1, i, h)
    style_header(ws2, 1, len(h2))
    for i, row in enumerate(fa_rows, 1):
        qty = row.get("qty", mtm.get(row["symbol"], {}).get("curr_qty") or open_ye.get(row["symbol"], {}).get("qty") or 0)
        _, _, name, _, _, nature = entity_row(row["symbol"], fin)
        ws2.append([
            i, row["country"], row["code"], row["symbol"], name, nature, qty, row["acq"],
            row["initial_usd"], row["peak_usd"], row["closing_usd"], row["div_usd"], row["sale_usd"],
        ])
    ws2.append([
        None, None, None, None, None, None, None, "TOTAL (USD)",
        round(sum(r["initial_usd"] for r in fa_rows), 2),
        round(sum(r["peak_usd"] for r in fa_rows), 2),
        round(sum(r["closing_usd"] for r in fa_rows), 2),
        round(sum(r["div_usd"] for r in fa_rows), 2),
        round(sum(r["sale_usd"] for r in fa_rows), 2),
    ])
    autosize(ws2)

    # FX Lookup (formula source) + legacy notes sheet
    fx_last, cross_start = write_fx_lookup(wb)

    # FX Rates (documentation)
    wsx = wb.create_sheet("FX Rates (USD-INR)")
    wsx["A1"] = "USD / INR SBI TT Buying rates used (month-end card rates — CA to verify exact date rates). Prefer editing 'FX Lookup' sheet for live CG formulas."
    wsx.append([])
    wsx.append(["Date", "Purpose", "USD/INR (Rs.)", "Source"])
    style_header(wsx, 3, 4)
    fx_notes = [
        (date(2024, 12, 31), "Initial value for securities held on 01-Jan-2025", 85.20),
        (date(2025, 12, 31), "Closing balance (year-end) / peak proxy", 89.47),
        (date(2025, 3, 31), "Rule 115 rate for Apr-2025 events", 85.10),
        (date(2025, 6, 30), "Rule 115 rate for Jul-2025 events", 85.10),
        (date(2025, 7, 31), "Rule 115 rate for Aug-2025 events", 87.15),
        (date(2025, 11, 30), "Rule 115 rate for Dec-2025 events", 88.95),
        (date(2026, 2, 28), "Rule 115 rate for Mar-2026 events", 90.56),
    ]
    for d, purpose, rate in fx_notes:
        wsx.append([d, purpose, rate, "SBI TTBR month-end compilation — verify on sbi.co.in"])
    wsx.append([])
    wsx.append(["NOTES:"])
    wsx.append(["1. Schedule FA: convert at SBI TT Buy on acquisition date (initial), peak date (peak), and 31-Dec (closing)."])
    wsx.append(["2. Income (dividends/interest/capital gains): Rule 115 — TT Buy on last day of month preceding the month of receipt/transfer."])
    wsx.append(["3. Non-USD holdings converted to USD using IBKR 31-Dec-2025 FX (EUR 1.1746, HKD 0.12849, JPY 0.0063827, DKK 0.15726), then to INR — editable on CG sheets."])
    wsx.append(["4. Peak per security estimated as MAX(initial, closing, sale proceeds) — IBKR Activity Statement does not publish daily per-symbol peaks."])
    autosize(wsx)

    # A2
    wsa2 = wb.create_sheet("A2 Custodial Account")
    wsa2["A1"] = (
        "A2. Details of Foreign Custodial Accounts held (including any beneficial interest) "
        "at any time during the calendar year 2025"
    )
    wsa2.append([
        "Country name", "Country code", "Name of Financial Institution", "Address of Financial Institution",
        "ZIP Code", "Account Number", "Status", "Account opening date",
    ])
    style_header(wsa2, 2, 8)
    a2_status = (
        "Owner - Joint beneficial owners"
        if (acct.get("Customer Type") or "").strip().lower() == "joint"
        else "Owner - Sole beneficial owner"
    )
    wsa2.append([
        "United States of America", 2,
        "Interactive Brokers LLC (Clearing Broker), Advisor Client via Interactive Brokers (India) / "
        f"Investment Advisor: {acct.get('Investment Advisor', '')}",
        "One Pickwick Plaza, Greenwich, CT, USA", "06830", acct.get("Account", ""),
        a2_status, open_date_note,
    ])
    wsa2.append([])
    wsa2.append([
        "Peak balance in account during period (USD)", "Peak balance (INR)",
        "Closing balance (USD)", "Closing balance (INR)",
        "Gross interest received (USD)", "Gross interest received (INR)",
        "Gross amount paid/credited (USD)", "Gross amount paid/credited (INR)",
    ])
    style_header(wsa2, 5, 8)
    closing_nav = nav.get("Ending Value") or 0
    starting_nav = nav.get("Starting Value") or 0
    # Change in NAV is in account base currency — convert to USD for Schedule FA A2
    if base_ccy != "USD":
        closing_nav = to_usd(closing_nav, base_ccy)
        starting_nav = to_usd(starting_nav, base_ccy)
    # Peak unknown — use max(start, end) + note
    peak_nav = max(closing_nav, starting_nav)
    int_usd = sum(
        (d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])) for d in int_cy
    )
    int_inr = sum(
        (d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])) * rule115_usd(d["date"])
        for d in int_cy
    )
    div_usd_tot = sum(
        (d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])) for d in div_cy
    )
    div_inr_tot = sum(
        (d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])) * rule115_usd(d["date"])
        for d in div_cy
    )
    gross_credited_usd = int_usd + div_usd_tot
    gross_credited_inr = int_inr + div_inr_tot
    wsa2.append([
        round(peak_nav, 2), round(peak_nav * rate_close),
        round(closing_nav, 2), round(closing_nav * rate_close),
        round(int_usd, 2), round(int_inr),
        round(gross_credited_usd, 2), round(gross_credited_inr),
    ])
    wsa2.append([])
    wsa2.append(["NOTES:"])
    wsa2.append([
        f"Account: IBKR {acct.get('Account')} — {acct.get('Name')} "
        f"({acct.get('Customer Type') or 'Individual'}, {base_ccy} base, Cash)."
    ])
    wsa2.append([f"Starting NAV 01-Jan-2025: USD {starting_nav:,.2f}; Ending NAV 31-Dec-2025: USD {closing_nav:,.2f}."])
    wsa2.append(["Peak NAV: Activity Statement does not include daily NAV — shown as max(start, end). Obtain PortfolioAnalyst for true peak."])
    dep_usd = _deposits_usd(deposits)
    dep_note = ", ".join(
        f"{d['date'].isoformat()}: {d['amount']} {d['currency']}" for d in deposits
    )
    wsa2.append([f"Deposits CY2025: USD {dep_usd:,.2f} ({dep_note})."])
    if base_ccy != "USD":
        wsa2.append([
            f"Base currency is {base_ccy}; A2 NAV converted to USD using IBKR YE cross-rate "
            f"(FX Lookup FCY→USD). Prefer PortfolioAnalyst USD NAV for filing."
        ])
    wsa2.append(["Gross amount paid/credited = interest + dividends credited in the custodial account during CY2025."])
    autosize(wsa2)

    # Schedule OS - CY dividends & interest
    wso = wb.create_sheet("Schedule OS - CY2025")
    wso["A1"] = "Interest & Dividend income from foreign sources — Calendar Year 2025 (Schedule FA cross-ref / Schedule OS)"
    wso.append(["Date", "Description", "Currency", "Amount (FCY)", "Amount (USD)", "USD/INR (Rule 115)", "Amount (INR)"])
    style_header(wso, 2, 7)
    wso.append(["— INTEREST —"])
    for d in int_cy:
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        rate = rule115_usd(d["date"])
        wso.append([d["date"], d["description"], d["currency"], d["amount"], round(usd, 4), rate, round(usd * rate)])
    wso.append([None, "TOTAL INTEREST", None, None, round(int_usd, 2), None, round(int_inr)])
    wso.append([])
    wso.append(["— DIVIDENDS —"])
    for d in div_cy:
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        rate = rule115_usd(d["date"])
        wso.append([d["date"], d["description"], d["currency"], d["amount"], round(usd, 4), rate, round(usd * rate)])
    wso.append([None, "TOTAL DIVIDENDS", None, None, round(div_usd_tot, 2), None, round(div_inr_tot)])
    autosize(wso)

    # Schedule OS FY
    wsof = wb.create_sheet("Schedule OS - FY2025-26")
    wsof["A1"] = "Interest & Dividend income — Financial Year 2025-26 (01-Apr-2025 to 31-Mar-2026) for ITR income schedules"
    wsof.append(["Date", "Description", "Currency", "Amount (FCY)", "Amount (USD)", "USD/INR (Rule 115)", "Amount (INR)"])
    style_header(wsof, 2, 7)
    int_fy_usd = int_fy_inr = 0.0
    wsof.append(["— INTEREST —"])
    for d in int_fy:
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        rate = rule115_usd(d["date"])
        inr = usd * rate
        int_fy_usd += usd
        int_fy_inr += inr
        wsof.append([d["date"], d["description"], d["currency"], d["amount"], round(usd, 4), rate, round(inr)])
    wsof.append([None, "TOTAL INTEREST", None, None, round(int_fy_usd, 2), None, round(int_fy_inr)])
    wsof.append([])
    div_fy_usd = div_fy_inr = 0.0
    wsof.append(["— DIVIDENDS —"])
    for d in div_fy:
        usd = d["amount"] if d["currency"] == "USD" else to_usd(d["amount"], d["currency"])
        rate = rule115_usd(d["date"])
        inr = usd * rate
        div_fy_usd += usd
        div_fy_inr += inr
        wsof.append([d["date"], d["description"], d["currency"], d["amount"], round(usd, 4), rate, round(inr)])
    wsof.append([None, "TOTAL DIVIDENDS", None, None, round(div_fy_usd, 2), None, round(div_fy_inr)])
    autosize(wsof)

    # Withholding / FSI
    wsw = wb.create_sheet("FTC - Withholding Tax")
    wsw["A1"] = "Foreign tax withheld (for Schedule FSI / TR) — review credit eligibility under DTAA"
    wsw.append(["Period", "Date", "Description", "Currency", "Amount (FCY)", "Amount (USD)", "USD/INR", "Amount (INR)"])
    style_header(wsw, 2, 8)
    for label, rows in [("CY2025", wht_cy), ("FY2025-26", wht_fy)]:
        for w in rows:
            # positive amounts in IBKR for reversals; tax withheld is negative
            usd = w["amount"] if w["currency"] == "USD" else to_usd(w["amount"], w["currency"])
            rate = rule115_usd(w["date"])
            wsw.append([label, w["date"], w["description"], w["currency"], w["amount"], round(usd, 4), rate, round(usd * rate)])
    autosize(wsw)

    # Capital Gains — formula-driven for reconciliation
    write_formula_cg_sheet(
        wb,
        "Capital Gains FY2025-26",
        "Capital gains — FY 2025-26 (AY 2026-27). FIFO. Yellow cells editable. Formulas compute USD, INR and Gain/(Loss).",
        cg_rows,
        fx_last,
        cross_start,
    )
    write_formula_cg_sheet(
        wb,
        "Capital Gains FY2024-25",
        "Capital gains — FY 2024-25 (AY 2025-26). FIFO from inception. Yellow cells editable.",
        cg_rows_2425,
        fx_last,
        cross_start,
    )

    # FIFO Opening lot register
    wsl = wb.create_sheet("FIFO Lots Register")
    wsl["A1"] = "FIFO lot register rebuilt from inception statement (actual acquisition dates & IBKR cost basis)"
    wsl.append(["As-of", "Symbol", "Qty", "Acquisition Date", "Cost (local)", "Currency", "Source"])
    style_header(wsl, 2, 7)
    for label, books in [("31-Dec-2024 (CY2025 open)", lots_ye2024), ("31-Mar-2025 (FY2025-26 open)", lots_fy_start)]:
        for sym in sorted(books):
            for L in books[sym]:
                wsl.append([label, sym, round(L.qty, 6), L.acq_date, round(L.cost_local, 4), L.currency, L.source])
    autosize(wsl)

    # CG Summary — formula links to FY2025-26 sheet where possible; values as cross-check
    wss = wb.create_sheet("CG Summary by Symbol")
    wss.append(["Symbol", "STCG INR", "LTCG INR", "Total Gain/(Loss) INR", "Sale Proceeds USD", "Cost USD"])
    style_header(wss, 1, 6)
    by = defaultdict(lambda: {"STCG": 0, "LTCG": 0, "proc": 0.0, "cost": 0.0})
    for r in cg_rows:
        by[r["symbol"]][r["type"]] += r["gain_inr"]
        by[r["symbol"]]["proc"] += r["proceeds_usd"]
        by[r["symbol"]]["cost"] += r["cost_usd"]
    for sym in sorted(by):
        wss.append([sym, by[sym]["STCG"], by[sym]["LTCG"], by[sym]["STCG"] + by[sym]["LTCG"],
                    round(by[sym]["proc"], 2), round(by[sym]["cost"], 2)])
    wss.append([])
    wss.append(["Note", "Detail P&L with editable FX is on 'Capital Gains FY2025-26' (formula-driven). This summary is a static cross-check."])
    autosize(wss)

    # Notes
    wsn = wb.create_sheet("Notes for CA")
    notes = [
        ("Assessee", acct.get("Name")),
        ("Account", f"{acct.get('Account')} — Interactive Brokers LLC (Advisor Client; Advisor: {acct.get('Investment Advisor')})"),
        ("Customer type", acct.get("Customer Type")),
        ("Base currency", acct.get("Base Currency")),
        ("Account funded", f"First deposit {deposits_inc[0]['date'].isoformat() if deposits_inc else 'n/a'} (inception statement)"),
        ("Schedule FA period", "Calendar Year 2025 (01-Jan-2025 to 31-Dec-2025) — for ITR AY 2026-27"),
        ("Income / CG period", "FY 2025-26 (01-Apr-2025 to 31-Mar-2026); also FY 2024-25 sheet for prior year"),
        ("Source data", "IBKR Activity Statements: Inception FY2024-25 + Annual CY2025 + Fiscal FY2025-26"),
        ("FIFO method", "Actual buy dates & IBKR Basis from inception replay; splits (LRCX 10:1, NFLX 10:1) applied"),
        ("CG reconciliation", "Sheets 'Capital Gains FY2025-26' / 'FY2024-25' are formula-driven — yellow cells editable; Gain = Sale INR − Comm INR − Cost INR"),
        ("FX Lookup", "Editable SBI TT USD/INR & EUR/INR month-end rates + FCY→USD cross rates"),
        ("Starting NAV CY2025", f"USD {starting_nav:,.2f}"),
        ("Ending NAV CY2025", f"USD {closing_nav:,.2f}"),
        ("Deposits CY2025", f"USD {_deposits_usd(deposits):,.2f}"),
        ("Deposits since inception (to 31-Mar-2025)", f"USD {_deposits_usd(deposits_inc):,.2f}"),
        ("Dividends CY2025", f"USD {div_usd_tot:,.2f} / INR {div_inr_tot:,.0f} (Rule 115)"),
        ("Interest CY2025", f"USD {int_usd:,.2f} / INR {int_inr:,.0f} (Rule 115)"),
        ("Dividends FY2025-26", f"USD {div_fy_usd:,.2f} / INR {div_fy_inr:,.0f}"),
        ("Interest FY2025-26", f"USD {int_fy_usd:,.2f} / INR {int_fy_inr:,.0f}"),
        ("STCG FY2025-26 (INR)", f"{stcg:,.0f}"),
        ("LTCG FY2025-26 (INR)", f"{ltcg:,.0f}"),
        ("STCG FY2024-25 (INR)", f"{stcg2425:,.0f}"),
        ("LTCG FY2024-25 (INR)", f"{ltcg2425:,.0f}"),
        ("A1 Foreign Bank/Depository", "Generally N/A separately — multi-currency cash sits inside IBKR custodial account (A2)"),
        ("A2 Custodial Account", f"Applicable — see sheet; opening: {open_date_note}"),
        ("A3 Equity & Debt Interest", f"Applicable — {len(fa_rows)} FIFO lot lines (one row per acquisition lot; sold lots match Capital Gains)"),
        ("A3 Col C (Name of entity)", "Symbol only (e.g. NOVd, AAPL). Legal name / address remain in Address & Nature columns."),
        ("A3 vs CG", "Sold lots: A3 Initial = CG Cost (INR), A3 Sale = CG Sale (INR), same Rule 115/EUR FX. Multiple buys (e.g. NOVd 21-Mar & 21-Aug) appear as separate A3 rows."),
        ("ADR country practice", "Underlying issuer country used for ADRs/GDRs (ASML NL, BTI UK, TSM TW, BYDDY CN, HYU KR, RIO UK)"),
        ("CPNG", "USA (Delaware corporation, NYSE) despite Korea operations"),
        ("NFLX split", "10-for-1 split on 14/17-Nov-2025 reflected in FIFO"),
        ("LRCX split", "10-for-1 split on 02/03-Oct-2024 reflected in FIFO (bought 1 pre-split → 10)"),
        ("HYU", "Cash merger acquisition Jul-2025 — sale proceeds USD 1,339.02 reported as redemption proceeds"),
        ("Action required", "1) Confirm exact SBI TT Buy card rates; 2) Obtain PortfolioAnalyst for true peak NAV; 3) Map withholding to Schedule FSI/TR under DTAA."),
    ]
    label = assessee_label or acct.get("Name") or "Assessee"
    wsn.append([f"Preparation notes — Schedule FA / CG / Dividends for {label}"])
    for a, b in (notes_extra or []):
        wsn.append([a, b])
    for a, b in notes:
        wsn.append([a, b])
    autosize(wsn)

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}")
    print(f"A3 rows: {len(fa_rows)}")
    print(f"CG lots: {len(cg_rows)} STCG={stcg} LTCG={ltcg}")
    print(f"Div CY USD={div_usd_tot:.2f} INR={div_inr_tot:.0f}")
    print(f"Int CY USD={int_usd:.2f} INR={int_inr:.0f}")
    print(f"Div FY USD={div_fy_usd:.2f} INR={div_fy_inr:.0f}")
    print(f"NAV end={closing_nav:.2f}")


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="Prepare Schedule FA / CG workbook from IBKR Activity Statements")
    p.add_argument("--inception", type=Path, default=None)
    p.add_argument("--annual", type=Path, default=None)
    p.add_argument("--fiscal", type=Path, default=None)
    p.add_argument("--out", type=Path, default=None)
    p.add_argument("--account-open-date", default=None)
    p.add_argument("--assessee", default=None)
    args = p.parse_args()
    main(
        inception=args.inception,
        annual=args.annual,
        fiscal=args.fiscal,
        out=args.out,
        account_open_date=args.account_open_date,
        assessee_label=args.assessee,
    )
