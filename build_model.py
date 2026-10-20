#!/usr/bin/env python3
"""Build ADBE_3statement_model.xlsx — a linked three-statement model with
live formulas. Historical inputs come from data/ADBE_edgar_series_*.csv
(SEC EDGAR company facts, pulled by pull_data.py). All $ in millions,
shares in millions, unless noted."""
import csv
import glob
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
CSV = sorted(glob.glob(os.path.join(HERE, "data", "ADBE_edgar_series_*.csv")))[-1]
rows = list(csv.DictReader(open(CSV)))

Q25 = ["2025-02-28", "2025-05-30", "2025-08-29"]
Q26 = ["2026-02-27", "2026-05-29", "2026-08-28"]
FY_END = {2021: "2021-12-03", 2022: "2022-12-02", 2023: "2023-12-01",
          2024: "2024-11-29", 2025: "2025-11-28"}
CF_ITEMS = {"da", "sbc", "cfo", "capex", "buybacks"}


def get(item, kind, end=None, fy=None):
    for r in rows:
        if r["item"] == item and r["kind"] == kind and \
                (end is None or r["end"] == end) and \
                (fy is None or int(r["fy"]) == fy):
            return float(r["val"]) / 1e6  # -> $M (shares -> millions)
    return None


def annual(item, fy):
    return get(item, "annual", fy=fy)


def nine_m(item, year):
    """9M value: IS items = sum of 3 quarters; CF items = YTD row."""
    ends = Q25 if year == 2025 else Q26
    if item in CF_ITEMS:
        return get(item, "ytd9m", end=ends[-1])
    return sum(get(item, "quarter", end=e) for e in ends)


def instant(item, end):
    return get(item, "instant", end=end)


HIST_ITEMS = ["revenue", "cost_of_revenue", "gross_profit", "rd", "sm", "ga",
              "amort_intangibles", "operating_income", "nonoperating_net",
              "pretax_income", "tax", "net_income", "da", "sbc", "cfo",
              "capex", "buybacks", "shares_diluted"]
H = {it: {fy: annual(it, fy) for fy in range(2021, 2026)} for it in HIST_ITEMS}
N25 = {it: nine_m(it, 2025) for it in HIST_ITEMS}
N26 = {it: nine_m(it, 2026) for it in HIST_ITEMS}
# 9M diluted WA shares = mean of the three quarterly weighted averages
for N, ends in ((N25, Q25), (N26, Q26)):
    N["shares_diluted"] = sum(get("shares_diluted", "quarter", end=e)
                              for e in ends) / 3
BS_ITEMS = ["cash", "short_term_investments", "receivables", "total_assets",
            "long_term_debt", "current_debt", "deferred_rev_current",
            "total_liabilities", "equity"]
BS = {it: {fy: instant(it, FY_END[fy]) for fy in range(2021, 2026)}
      for it in BS_ITEMS}
BSL = {it: instant(it, "2026-08-28") for it in BS_ITEMS}
SH_OUT = {2023: instant("shares_outstanding", "2024-01-05"),
          2024: instant("shares_outstanding", "2025-01-03"),
          2025: instant("shares_outstanding", "2026-01-09")}
SH_LATEST = instant("shares_outstanding", "2026-09-18")

# ---------------------------------------------------------------- styles
NAVY = "203764"
HDR = PatternFill("solid", fgColor=NAVY)
HDRF = Font(bold=True, color="FFFFFF")
TITLEF = Font(bold=True, size=14, color=NAVY)
SUBF = Font(bold=True, size=11, color=NAVY)
BOLD = Font(bold=True)
INP = PatternFill("solid", fgColor="FFF2CC")
SECT = PatternFill("solid", fgColor="D9E2F3")
THIN = Border(*[Side(style="thin", color="B0B0B0")] * 4)
MONEY = "#,##0"
MONEY1 = "#,##0.0"
PCT = "0.0%"
PX = '"$"#,##0.00'
MULT = '0.0"x"'

wb = Workbook()

def put(ws, cell, value, fmt=None, font=None, fill=None, border=False):
    c = ws[cell]
    c.value = value
    if fmt: c.number_format = fmt
    if font: c.font = font
    if fill: c.fill = fill
    if border: c.border = THIN
    return c

def header_row(ws, row, labels, widths=None):
    for i, lab in enumerate(labels):
        put(ws, f"{get_column_letter(1+i)}{row}", lab, font=HDRF, fill=HDR)
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(1+i)].width = w

# ================================================================ HISTORICALS
ws = wb.active
ws.title = "Historicals"
put(ws, "A1", "Adobe Inc. (NASDAQ: ADBE) — Historical Financials", font=TITLEF)
put(ws, "A2", "Source: SEC EDGAR company facts (CIK 0000796343), pulled 2026-10-20 via pull_data.py. "
              "Adobe's fiscal year ends the Friday nearest Nov 30 (FY2025 ended Nov 28, 2025). "
              "$ in millions; shares in millions. LTM = FY2025 − 9M FY2025 + 9M FY2026 (live formulas).")
header_row(ws, 4, ["", "EDGAR concept", "FY2021", "FY2022", "FY2023", "FY2024",
                   "FY2025", "9M FY2025", "9M FY2026", "LTM to Q3 FY26",
                   "Latest BS — Q3 FY26 (Aug 28, 2026)"],
           [30, 42, 12, 12, 12, 12, 12, 12, 12, 13, 15])
IS_ROWS = [("revenue", "Revenue", "Revenues"),
           ("cost_of_revenue", "Cost of revenue", "CostOfRevenue"),
           ("gross_profit", "Gross profit", "GrossProfit"),
           ("rd", "Research & development", "ResearchAndDevelopmentExpense (software)"),
           ("sm", "Sales & marketing", "SellingAndMarketingExpense"),
           ("ga", "General & administrative", "GeneralAndAdministrativeExpense"),
           ("amort_intangibles", "Amortization of acquired intangibles", "Derived: OperatingExpenses − R&D − S&M − G&A"),
           ("operating_income", "Operating income", "OperatingIncomeLoss"),
           ("nonoperating_net", "Nonoperating income (expense), net", "NonoperatingIncomeExpense"),
           ("pretax_income", "Income before taxes", "IncomeLossFromContinuingOperationsBeforeIncomeTaxes..."),
           ("tax", "Income tax expense", "IncomeTaxExpenseBenefit"),
           ("net_income", "Net income", "NetIncomeLoss")]
R = {}
r = 5
for item, label, concept in IS_ROWS:
    put(ws, f"A{r}", label); put(ws, f"B{r}", concept)
    for j, fy in enumerate(range(2021, 2026)):
        put(ws, f"{get_column_letter(3+j)}{r}", round(H[item][fy]), MONEY)
    put(ws, f"H{r}", round(N25[item]), MONEY)
    put(ws, f"I{r}", round(N26[item]), MONEY)
    put(ws, f"J{r}", f"=G{r}-H{r}+I{r}", MONEY)
    R[item] = r; r += 1
put(ws, f"A{r}", "Diluted weighted-average shares", font=BOLD)
put(ws, f"B{r}", "WeightedAverageNumberOfDilutedSharesOutstanding (9M = mean of quarters)")
for j, fy in enumerate(range(2021, 2026)):
    put(ws, f"{get_column_letter(3+j)}{r}", round(H["shares_diluted"][fy], 1), MONEY1)
put(ws, f"H{r}", round(N25["shares_diluted"], 1), MONEY1)
put(ws, f"I{r}", round(N26["shares_diluted"], 1), MONEY1)
put(ws, f"J{r}", f"=G{r}", MONEY1)
R["shares_diluted"] = r; r += 1
put(ws, f"A{r}", "Diluted EPS (GAAP)", font=BOLD)
for col in "CDEFGHIJ":
    put(ws, f"{col}{r}", f"={col}{R['net_income']}/{col}{R['shares_diluted']}", PX)
R["eps"] = r; r += 2
CF_ROWS = [("da", "Depreciation & amortization (CF stmt)", "DepreciationDepletionAndAmortization"),
           ("sbc", "Stock-based compensation", "ShareBasedCompensation"),
           ("cfo", "Cash flow from operations", "NetCashProvidedByUsedInOperatingActivities"),
           ("capex", "Capital expenditures", "PaymentsToAcquirePropertyPlantAndEquipment"),
           ("buybacks", "Common stock repurchases", "PaymentsForRepurchaseOfCommonStock")]
for item, label, concept in CF_ROWS:
    put(ws, f"A{r}", label); put(ws, f"B{r}", concept)
    for j, fy in enumerate(range(2021, 2026)):
        put(ws, f"{get_column_letter(3+j)}{r}", round(H[item][fy]), MONEY)
    put(ws, f"H{r}", round(N25[item]), MONEY)
    put(ws, f"I{r}", round(N26[item]), MONEY)
    put(ws, f"J{r}", f"=G{r}-H{r}+I{r}", MONEY)
    R[item] = r; r += 1
put(ws, f"A{r}", "Free cash flow (CFO − capex)", font=BOLD)
for col in "CDEFGHIJ":
    put(ws, f"{col}{r}", f"={col}{R['cfo']}-{col}{R['capex']}", MONEY)
R["fcf"] = r; r += 1
put(ws, f"A{r}", "EBITDA (operating income + D&A)", font=BOLD)
for col in "CDEFGHIJ":
    put(ws, f"{col}{r}", f"={col}{R['operating_income']}+{col}{R['da']}", MONEY)
R["ebitda"] = r; r += 2
BS_ROWS = [("cash", "Cash & equivalents", "CashAndCashEquivalentsAtCarryingValue"),
           ("short_term_investments", "Short-term investments", "ShortTermInvestments"),
           ("receivables", "Trade receivables, net", "AccountsReceivableNetCurrent"),
           ("total_assets", "Total assets", "Assets"),
           ("long_term_debt", "Long-term debt", "LongTermDebt"),
           ("current_debt", "Current debt", "DebtCurrent"),
           ("deferred_rev_current", "Deferred revenue (current)", "ContractWithCustomerLiabilityCurrent"),
           ("total_liabilities", "Total liabilities", "Liabilities"),
           ("equity", "Total stockholders' equity", "StockholdersEquity")]
for item, label, concept in BS_ROWS:
    put(ws, f"A{r}", label); put(ws, f"B{r}", concept)
    for j, fy in enumerate(range(2021, 2026)):
        put(ws, f"{get_column_letter(3+j)}{r}", round(BS[item][fy]), MONEY)
    put(ws, f"K{r}", round(BSL[item]), MONEY)
    R[item] = r; r += 1
put(ws, f"A{r}", "Total debt (LT + current)", font=BOLD)
for col in "CDEFGK":
    put(ws, f"{col}{r}", f"={col}{R['long_term_debt']}+{col}{R['current_debt']}", MONEY)
R["total_debt"] = r; r += 1
put(ws, f"A{r}", "Shares outstanding (period end)", font=BOLD)
put(ws, f"B{r}", "dei EntityCommonStockSharesOutstanding — 10-K cover dates for FY23–FY25; Q3 FY26 10-Q cover (Sep 18, 2026) in Latest column")
put(ws, f"E{r}", round(SH_OUT[2023], 1), MONEY1)
put(ws, f"F{r}", round(SH_OUT[2024], 1), MONEY1)
put(ws, f"G{r}", round(SH_OUT[2025], 1), MONEY1)
put(ws, f"K{r}", round(SH_LATEST, 1), MONEY1)
R["shares_out"] = r
ws.freeze_panes = "C5"
HIST = R

# ================================================================ ASSUMPTIONS
ws = wb.create_sheet("Assumptions")
put(ws, "A1", "Assumptions — every yellow cell is an input; the rationale is stated next to it", font=TITLEF)
header_row(ws, 3, ["Assumption", "Value", "Source / rationale"], [46, 14, 110])
A = {}
r = 4
put(ws, f"A{r}", "MARKET (as of Oct 20, 2026)", font=BOLD, fill=SECT); r += 1
def arow(label, value, why, fmt=MONEY1, key=None):
    global r
    put(ws, f"A{r}", label)
    put(ws, f"B{r}", value, fmt, fill=INP)
    put(ws, f"C{r}", why)
    if key: A[key] = f"Assumptions!$B${r}"
    r += 1
arow("Share price — close, Mon Oct 19, 2026 ($)", 247.09,
     "MarketBeat price history and Finnhub quote, retrieved Oct 20, 2026; corroborated by stockanalysis.com (close Oct 19, 2026, 4:00 PM ET). Stooq CSV was JS-blocked.", PX, "price")
arow("Shares outstanding (M) — Sep 18, 2026", round(SH_LATEST, 1),
     "SEC EDGAR dei EntityCommonStockSharesOutstanding, Q3 FY2026 10-Q cover (filed Sep 22, 2026). Market cap at the close: $96.2B (247.09 × 389.2M).", MONEY1, "shares")
arow("Net debt at Q3 FY26 ($M) — debt less cash & ST investments", f"=Historicals!K{HIST['total_debt']}-Historicals!K{HIST['cash']}-Historicals!K{HIST['short_term_investments']}",
     "Formula from Historicals (Q3 FY2026 balance sheet, Aug 28, 2026): total debt $6,363M − cash $4,359M − ST investments $1,280M = $724M net debt — Adobe is essentially net-debt neutral.", MONEY, "net_debt")
r += 1
put(ws, f"A{r}", "Q4 FY2026E BUILD — FY2026E = 9M FY2026 actual + Q4 estimate. Q4 FY2025A is derived (FY2025 − 9M FY2025, formulas).", font=BOLD, fill=SECT); r += 1
put(ws, f"A{r}", "", font=BOLD); put(ws, f"B{r}", "Q4 FY25A", font=BOLD)
put(ws, f"C{r}", "Growth / input", font=BOLD); put(ws, f"D{r}", "Q4 FY26E", font=BOLD)
put(ws, f"E{r}", "Rationale", font=BOLD)
q4h = r; r += 1
Q4 = {}
Q4DEF = [("revenue", "Revenue", "growth", 0.12, "Q1–Q3 FY26 grew +12.0%/+12.7%/+12.9% YoY; Q4 guided confidently (non-GAAP EPS guide $6.30–6.35). I use +12.0% — no further acceleration assumed."),
         ("cost_of_revenue", "Cost of revenue", "growth", 0.13, "AI inference/hosting costs growing slightly faster than revenue, as in 9M FY26 (CoR +13% YoY)."),
         ("rd", "Research & development", "growth", 0.12, "In line with 9M FY26 R&D growth (+13% YoY) and continued Firefly/AI investment."),
         ("sm", "Sales & marketing", "growth", 0.11, "Below revenue growth — modest leverage, consistent with 9M FY26."),
         ("ga", "General & administrative", "growth", 0.05, "G&A ran hot in 9M FY26 (+13% YoY on deal/legal costs); I normalize Q4 growth to +5%."),
         ("amort_intangibles", "Amortization of acquired intangibles", "usd", 40, "Derived IS amortization was $157M in FY2025 and is declining as acquired intangibles roll off (9M FY26: ~$112M); Q4 stub of $40M."),
         ("nonoperating_net", "Nonoperating income, net", "usd", 45, "Interest income on ~$5.6B cash/STI roughly offsets interest on $6.4B of notes; 9M FY26 net was +$7M."),
         ("sbc", "Stock-based compensation", "usd", 520, "Q4 FY25 SBC was $485M (FY25 $1,938M − 9M $1,453M); +7% YoY in line with headcount/comp drift."),
         ("da", "Depreciation & amortization (CF)", "usd", 210, "9M FY26 D&A was $582M (EDGAR YTD); Q4 run-rate ~$200M/quarter."),
         ("capex", "Capital expenditures", "usd", 55, "Capex is minimal at Adobe (FY25: $183M, 0.8% of revenue); 9M FY26 was $180M incl. facilities fit-out."),
         ("buybacks", "Common stock repurchases", "usd", 2600, "9M FY26 repurchases were $6,821M (~$2.3B/qtr); Q4 is seasonally the largest FCF quarter — $2.6B assumed.")]
for item, label, mode, val, why in Q4DEF:
    col_hist = "revenue" if item != "x" else item
    put(ws, f"A{r}", label)
    put(ws, f"B{r}", f"=Historicals!G{HIST[item]}-Historicals!H{HIST[item]}", MONEY)
    if mode == "growth":
        put(ws, f"C{r}", val, PCT, fill=INP)
        put(ws, f"D{r}", f"=B{r}*(1+C{r})", MONEY)
    else:
        put(ws, f"D{r}", val, MONEY, fill=INP)
    put(ws, f"E{r}", why)
    Q4[item] = {"q4a": f"Assumptions!$B${r}", "q4e": f"Assumptions!$D${r}"}
    r += 1
arow("Q4 FY26E effective tax rate", 0.19, "Q4 FY25 ETR was ~28% on discrete items; 9M FY26 ETR was 22.6%. I use 19% for the Q4 stub — near the FY25 full-year rate of 18.3%.", PCT, "q4_etr")
arow("Q4 FY26E diluted shares (M)", 383, "Share count fell from 410.8M (Q1 WA) to 395.5M (Q3 WA) during FY26; Q4 buybacks continue the decline.", MONEY1, "q4_shares")
r += 1
put(ws, f"A{r}", "FORECAST DRIVERS — FY2026E / FY2027E / FY2028E", font=BOLD, fill=SECT); r += 1
put(ws, f"B{r}", "FY26E", font=BOLD); put(ws, f"C{r}", "FY27E", font=BOLD)
put(ws, f"D{r}", "FY28E", font=BOLD); put(ws, f"E{r}", "Rationale", font=BOLD)
r += 1
DRV = {}
DRVDEF = [("ar_pct", "Trade receivables, % of revenue", (0.097, 0.097, 0.097), PCT, "FY25 actual: 9.9% of revenue (receivables are seasonally highest at the Nov year-end). Held flat."),
          ("defrev_g", "Deferred revenue (current) growth", (0.10, 0.10, 0.10), PCT, "Deferred revenue grew +12.7% in FY25 and tracks subscription bookings; I grow it slightly below revenue growth."),
          ("debt_total", "Total debt, year-end ($M)", (6250, None, None), MONEY, "FY26E year-end input: Q3 FY26 actual total debt was $6,363M; assumes ~$113M net repayment in Q4. FY27E/FY28E use the change row below."),
          ("debt_change", "Total debt change ($M)", (None, -250, -250), MONEY, "Maturing notes repaid from cash flow; no new issuance assumed — Adobe has no funding need."),
          ("avg_bb_price", "Average repurchase price ($)", (265, 300, 340), PX, "Blended FY26 average ~$265 (the stock fell through the year: Q3 10-Q cover count implies heavy H2 buying at lower prices); FY27/28 assume prices recover toward — but stay below — my target, which is conservative for share-count decline."),
          ("sbc_issue", "Shares issued for equity comp (M)", (6.0, 6.0, 6.0), MONEY1, "Gross issuance under equity plans before repurchases; roughly offsets ~15% of the buyback share reduction.")]
for key, label, vals, fmt, why in DRVDEF:
    put(ws, f"A{r}", label)
    for j, v in enumerate(vals):
        if v is not None:
            put(ws, f"{get_column_letter(2+j)}{r}", v, fmt, fill=INP)
    put(ws, f"E{r}", why)
    DRV[key] = {2026: f"Assumptions!$B${r}", 2027: f"Assumptions!$C${r}", 2028: f"Assumptions!$D${r}"}
    r += 1
# FY27/28-only drivers in C/D columns
DRV2 = {}
DRV2DEF = [("rev_g", "Revenue growth", (0.11, 0.105), PCT, "FY26E grows +12.4% (9M actual + Q4E). I fade growth slowly — +11.0%, then +10.5% — below the FY26 run-rate; no AI-driven re-acceleration is assumed, which keeps the pitch honest."),
           ("gm", "Gross margin", (0.892, 0.893), PCT, "FY25: 89.3%; 9M FY26: 89.2%. AI inference costs pressure, mix (Document Cloud, enterprise) offsets. Held ~flat."),
           ("rd_pct", "R&D, % of revenue", (0.180, 0.179), PCT, "FY25: 18.0%; 9M FY26: 18.2%. AI product investment stays elevated; slight leverage by FY28."),
           ("sm_pct", "Sales & marketing, % of revenue", (0.270, 0.267), PCT, "FY25: 27.3%; 9M FY26: 27.4%. Modest leverage as self-serve/PLG mix grows."),
           ("ga_pct", "General & administrative, % of revenue", (0.068, 0.067), PCT, "FY25: 6.6%; 9M FY26 ran at 7.6% (deal/legal costs). I normalize toward ~6.7–6.8%."),
           ("amort", "Amortization of acquired intangibles ($M)", (150, 140), MONEY, "Declining schedule as acquired intangibles roll off (derived IS line: FY25 $157M; FY26E ~$152M)."),
           ("int_rate", "Interest rate on beginning debt", (0.041, 0.041), PCT, "Blended coupon: FY25 implied interest was ~$220M on ~$5.4B average debt (nonoperating net +$28M vs other income +$248M)."),
           ("cash_yield", "Interest yield on beginning cash & STI", (0.038, 0.038), PCT, "Money-market/T-bill yields on the cash pile, net of other nonoperating items."),
           ("etr", "Effective tax rate", (0.185, 0.185), PCT, "FY25 ETR: 18.3%; FY23–FY25 range 18–21%. I use 18.5%."),
           ("sbc_pct", "Stock-based compensation, % of revenue", (0.078, 0.076), PCT, "FY25: 8.2% of revenue; 9M FY26: 8.0%. Slow improvement assumed — SBC is the single biggest GAAP/non-GAAP gap and a real cost I do not add back in the DCF."),
           ("da_usd", "Depreciation & amortization, CF stmt ($M)", (800, 790), MONEY, "FY25: $820M; declining as the asset base seasons. Note: D&A here is the cash-flow-statement line (incl. intangible amortization)."),
           ("capex_pct", "Capital expenditures, % of revenue", (0.008, 0.008), PCT, "FY25 actual: 0.8% of revenue. Adobe is asset-light; AI compute is largely opex (cloud) or in cost of revenue."),
           ("buybacks_usd", "Common stock repurchases ($M)", (9500, 10000), MONEY, "FY25 actual: $11,283M (115% of FCF); 9M FY26: $6,821M. I assume ~$9.5–10.0B/yr — below the FY25 pace, funded by FCF while keeping cash >$4B.")]
for key, label, vals, fmt, why in DRV2DEF:
    put(ws, f"A{r}", label)
    for j, v in enumerate(vals):
        put(ws, f"{get_column_letter(3+j)}{r}", v, fmt, fill=INP)
    put(ws, f"E{r}", why)
    DRV2[key] = {2027: f"Assumptions!$C${r}", 2028: f"Assumptions!$D${r}"}
    r += 1
r += 1
put(ws, f"A{r}", "WACC & TERMINAL VALUE", font=BOLD, fill=SECT); r += 1
def wrow(label, value, why, fmt, key):
    global r
    put(ws, f"A{r}", label); put(ws, f"B{r}", value, fmt, fill=INP if isinstance(value, (int, float)) else None)
    put(ws, f"C{r}", why); A[key] = f"Assumptions!$B${r}"; r += 1
wrow("Risk-free rate (10Y UST)", 0.051, "House convention (5.1%), the same rate used across my sector teardown models — deliberately demanding.", PCT, "rf")
wrow("Beta", 1.42, "Finnhub (1.4126) / MarketBeat (1.42), retrieved Oct 20, 2026. Elevated vs Adobe's pre-2024 ~1.0 — the de-rating raised measured beta; I take it as reported rather than flattering the WACC.", PCT, "beta") if False else None
put(ws, f"A{r}", "Beta"); put(ws, f"B{r}", 1.42, "0.00", fill=INP)
put(ws, f"C{r}", "Finnhub (1.4126) / MarketBeat (1.42), retrieved Oct 20, 2026. Elevated vs Adobe's pre-2024 ~1.0 — the de-rating raised measured beta; I take it as reported rather than flattering the WACC.")
A["beta"] = f"Assumptions!$B${r}"; r += 1
wrow("Equity risk premium", 0.045, "House convention (4.5%), consistent with my sector teardown framework.", PCT, "erp")
put(ws, f"A{r}", "Cost of equity (CAPM)")
put(ws, f"B{r}", f"={A['rf']}+{A['beta']}*{A['erp']}", PCT)
put(ws, f"C{r}", "Formula: Rf + beta × ERP.")
A["coe"] = f"Assumptions!$B${r}"; r += 1
put(ws, f"A{r}", "Pre-tax cost of debt"); put(ws, f"B{r}", f"={DRV2['int_rate'][2027]}", PCT)
put(ws, f"C{r}", "Linked to the interest-rate assumption (blended coupon ~4.1%).")
A["cod_pre"] = f"Assumptions!$B${r}"; r += 1
put(ws, f"A{r}", "After-tax cost of debt")
put(ws, f"B{r}", f"={A['cod_pre']}*(1-{DRV2['etr'][2027]})", PCT)
put(ws, f"C{r}", "Formula: pre-tax cost × (1 − ETR).")
A["cod_after"] = f"Assumptions!$B${r}"; r += 1
put(ws, f"A{r}", "Market value of equity ($M)"); put(ws, f"B{r}", f"={A['price']}*{A['shares']}", MONEY)
put(ws, f"C{r}", "Formula: price × shares outstanding.")
A["mve"] = f"Assumptions!$B${r}"; r += 1
put(ws, f"A{r}", "Weight of equity"); put(ws, f"B{r}", f"={A['mve']}/({A['mve']}+Historicals!K{HIST['total_debt']})", PCT)
put(ws, f"C{r}", "Market-value weights using Q3 FY26 gross debt.")
A["we"] = f"Assumptions!$B${r}"; r += 1
put(ws, f"A{r}", "WACC"); put(ws, f"B{r}", f"={A['we']}*{A['coe']}+(1-{A['we']})*{A['cod_after']}", PCT, font=BOLD)
put(ws, f"C{r}", "Formula: wE × CoE + wD × after-tax CoD. ≈11.0% — a high bar that the cash flows still clear.")
A["wacc"] = f"Assumptions!$B${r}"; r += 1
wrow("Terminal growth rate", 0.03, "House convention (3.0%) — roughly long-run nominal growth for a maturing software franchise; below Adobe's forecast growth in every year.", PCT, "tg")
r += 1
put(ws, f"A{r}", "DCF EXTENSION & SCENARIOS", font=BOLD, fill=SECT); r += 1
wrow("FCF growth FY2029E", 0.09, "Fade from FY28 FCF growth as the base enlarges.", PCT, "g29")
wrow("FCF growth FY2030E", 0.08, "Fade continues.", PCT, "g30")
wrow("FCF growth FY2031E", 0.07, "Still ~2.3× the terminal growth rate at exit.", PCT, "g31")
wrow("Bear case: FCF level shock", -0.18, "Bear = base FCF path × (1 − 18%), WACC +50bp, terminal growth 2.5%: growth disappoints and AI spend eats margin, but the franchise stays intact.", PCT, "bear_shock")
wrow("Bear case: WACC add-on", 0.005, "", PCT, "bear_wacc")
wrow("Bear case: terminal growth", 0.025, "", PCT, "bear_g")
wrow("Bull case: FCF level shock", 0.08, "Bull = base FCF path × 1.08, WACC −50bp, terminal growth 3.5%: AI monetization (Firefly credits, Acrobat AI) lands and the multiple debate shifts back to growth.", PCT, "bull_shock")
wrow("Bull case: WACC add-on", -0.005, "", PCT, "bull_wacc")
wrow("Bull case: terminal growth", 0.035, "", PCT, "bull_g")
wrow("Disruption case: WACC", 0.115, "The scenario the market is pricing, in my view: AI-native tools take real share, revenue growth fades to low-single digits and FCF margin compresses toward ~30%. FCF path is a direct input on the Valuation sheet.", PCT, "dis_wacc")
wrow("Disruption case: terminal growth", 0.015, "A melting-ice-cube terminal assumption — deliberately punitive.", PCT, "dis_g")
ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 32; ws.column_dimensions["D"].width = 14
ws.column_dimensions["E"].width = 95

# ================================================================ INCOME STATEMENT
ws = wb.create_sheet("IS")
put(ws, "A1", "Income Statement — actual FY2025 and forecast FY2026E–FY2028E ($M, except per-share)", font=TITLEF)
header_row(ws, 3, ["", "FY2025A", "FY2026E", "FY2027E", "FY2028E"], [44, 13, 13, 13, 13])
I = {}
def irow(row, label, c, d, e, f, fmt=MONEY, bold=False):
    put(ws, f"A{row}", label, font=BOLD if bold else None)
    for col, v in zip("CDEF", (c, d, e, f)):
        if v is not None: put(ws, f"{col}{row}", v, fmt, font=BOLD if bold else None)
    I[label] = row
Hh = lambda item, col="G": f"Historicals!{col}{HIST[item]}"
irow(4, "Revenue", f"={Hh('revenue')}",
     f"=Historicals!I{HIST['revenue']}+{Q4['revenue']['q4e']}",
     f"=D4*(1+{DRV2['rev_g'][2027]})", f"=E4*(1+{DRV2['rev_g'][2028]})", bold=True)
irow(5, "Cost of revenue", f"={Hh('cost_of_revenue')}",
     f"=Historicals!I{HIST['cost_of_revenue']}+{Q4['cost_of_revenue']['q4e']}",
     f"=E4*(1-{DRV2['gm'][2027]})", f"=F4*(1-{DRV2['gm'][2028]})")
irow(6, "Gross profit", "=C4-C5", "=D4-D5", "=E4-E5", "=F4-F5", bold=True)
irow(7, "Research & development", f"={Hh('rd')}",
     f"=Historicals!I{HIST['rd']}+{Q4['rd']['q4e']}",
     f"=E4*{DRV2['rd_pct'][2027]}", f"=F4*{DRV2['rd_pct'][2028]}")
irow(8, "Sales & marketing", f"={Hh('sm')}",
     f"=Historicals!I{HIST['sm']}+{Q4['sm']['q4e']}",
     f"=E4*{DRV2['sm_pct'][2027]}", f"=F4*{DRV2['sm_pct'][2028]}")
irow(9, "General & administrative", f"={Hh('ga')}",
     f"=Historicals!I{HIST['ga']}+{Q4['ga']['q4e']}",
     f"=E4*{DRV2['ga_pct'][2027]}", f"=F4*{DRV2['ga_pct'][2028]}")
irow(10, "Amortization of acquired intangibles", f"={Hh('amort_intangibles')}",
     f"=Historicals!I{HIST['amort_intangibles']}+{Q4['amort_intangibles']['q4e']}",
     f"={DRV2['amort'][2027]}", f"={DRV2['amort'][2028]}")
irow(11, "Operating income", "=C6-C7-C8-C9-C10", "=D6-D7-D8-D9-D10",
     "=E6-E7-E8-E9-E10", "=F6-F7-F8-F9-F10", bold=True)
irow(12, "Interest expense", None, 0,
     f"='CF_BS'!D28*{DRV2['int_rate'][2027]}", f"='CF_BS'!E28*{DRV2['int_rate'][2028]}")
irow(13, "Nonoperating income (expense), net", f"={Hh('nonoperating_net')}",
     f"=Historicals!I{HIST['nonoperating_net']}+{Q4['nonoperating_net']['q4e']}",
     f"=('CF_BS'!D22+'CF_BS'!D23)*{DRV2['cash_yield'][2027]}", f"=('CF_BS'!E22+'CF_BS'!E23)*{DRV2['cash_yield'][2028]}")
irow(14, "Income before taxes", "=C11+C13", "=D11+D13", "=E11-E12+E13", "=F11-F12+F13", bold=True)
irow(15, "Income tax expense", f"={Hh('tax')}",
     f"=Historicals!I{HIST['tax']}+(D14-Historicals!I{HIST['pretax_income']})*{A['q4_etr']}",
     f"=E14*{DRV2['etr'][2027]}", f"=F14*{DRV2['etr'][2028]}")
irow(16, "Net income", "=C14-C15", "=D14-D15", "=E14-E15", "=F14-F15", bold=True)
irow(17, "Diluted weighted-average shares", f"={Hh('shares_diluted', 'G')}",
     f"=(Historicals!I{HIST['shares_diluted']}*9+{A['q4_shares']}*3)/12",
     "='CF_BS'!E47", "='CF_BS'!F47", MONEY1)
irow(18, "Diluted EPS (GAAP)", "=C16/C17", "=D16/D17", "=E16/E17", "=F16/F17", PX)
irow(19, "Memo: non-GAAP EPS proxy (NI + SBC + amort. of acquired intangibles, pre-tax)",
     None, None, None, None, PX)
put(ws, "C19", f"=(C16+Historicals!G{HIST['sbc']}+C10)/C17", PX)
put(ws, "D19", f"=(D16+'CF_BS'!D7+D10)/D17", PX)
put(ws, "E19", f"=(E16+'CF_BS'!E7+E10)/E17", PX)
put(ws, "F19", f"=(F16+'CF_BS'!F7+F10)/F17", PX)
put(ws, "A21", "Notes: (1) FY2025A interest expense is not broken out in EDGAR company facts after FY2023 — it sits inside nonoperating income, net; forecast years split interest expense (on beginning debt) and interest/other income (on beginning cash & ST investments) explicitly. "
               "(2) Adobe guides on a non-GAAP basis (FY2026 guide: $24.45–$24.50, per the Q3 FY26 release). "
               "The proxy above adds back SBC and acquired-intangible amortization pre-tax, so it overstates slightly vs Adobe's tax-effected non-GAAP EPS — "
               "it is a reasonableness check on the forecast, not a replication of Adobe's non-GAAP definition.", font=Font(italic=True, size=9))

# ================================================================ CF + BS
ws = wb.create_sheet("CF_BS")
put(ws, "A1", "Cash Flow & Balance Sheet — linked forecast ($M). The balance sheet balances by construction: every balance-sheet delta appears once in the cash flow, and equity rolls with net income + SBC − buybacks.", font=TITLEF)
header_row(ws, 3, ["", "FY2025A", "FY2026E", "FY2027E", "FY2028E"], [44, 13, 13, 13, 13])
B = {}
def brow(row, label, c, d, e, f, fmt=MONEY, bold=False):
    put(ws, f"A{row}", label, font=BOLD if bold else None)
    for col, v in zip("CDEF", (c, d, e, f)):
        if v is not None: put(ws, f"{col}{row}", v, fmt, font=BOLD if bold else None)
    B[label] = row
brow(5, "Net income", f"=IS!C16", "=IS!D16", "=IS!E16", "=IS!F16")
brow(6, "Plus: depreciation & amortization", f"={Hh('da')}",
     f"=Historicals!I{HIST['da']}+{Q4['da']['q4e']}",
     f"={DRV2['da_usd'][2027]}", f"={DRV2['da_usd'][2028]}")
brow(7, "Plus: stock-based compensation (label)", None, None, None, None)  # placeholder replaced below
B.pop("Plus: stock-based compensation (label)")
put(ws, "A7", "Plus: stock-based compensation")
put(ws, "C7", f"={Hh('sbc')}", MONEY)
put(ws, "D7", f"=Historicals!I{HIST['sbc']}+{Q4['sbc']['q4e']}", MONEY)
put(ws, "E7", f"=IS!E4*{DRV2['sbc_pct'][2027]}", MONEY)
put(ws, "F7", f"=IS!F4*{DRV2['sbc_pct'][2028]}", MONEY)
brow(8, "Less: increase in trade receivables", None,
     "=-(D31-C31)", "=-(E31-D31)", "=-(F31-E31)")
brow(9, "Plus: increase in deferred revenue", None,
     "=D35-C35", "=E35-D35", "=F35-E35")
brow(10, "Plus: increase in other liabilities", None,
     "=D36-C36", "=E36-D36", "=F36-E36")
brow(11, "Cash flow from operations", f"={Hh('cfo')}",
     "=SUM(D5:D7)+D8+D9+D10", "=SUM(E5:E7)+E8+E9+E10", "=SUM(F5:F7)+F8+F9+F10", bold=True)
brow(12, "Capital expenditures", f"={Hh('capex')}",
     f"=Historicals!I{HIST['capex']}+{Q4['capex']['q4e']}",
     f"=IS!E4*{DRV2['capex_pct'][2027]}", f"=IS!F4*{DRV2['capex_pct'][2028]}")
brow(13, "Free cash flow", "=C11-C12", "=D11-D12", "=E11-E12", "=F11-F12", bold=True)
brow(14, "Increase in short-term investments", None, 0, 0, 0)
brow(15, "Common stock repurchases", f"={Hh('buybacks')}",
     f"=Historicals!I{HIST['buybacks']}+{Q4['buybacks']['q4e']}",
     f"={DRV2['buybacks_usd'][2027]}", f"={DRV2['buybacks_usd'][2028]}")
brow(16, "Net debt issued (repaid)", None,
     f"=D34-C34", f"={DRV['debt_change'][2027]}", f"={DRV['debt_change'][2028]}")
brow(17, "Change in cash", None, "=D11-D12-D14-D15+D16", "=E11-E12-E14-E15+E16",
     "=F11-F12-F14-F15+F16", bold=True)
brow(18, "Cash, beginning of year", None, "=C28", "=D28", "=E28")
brow(19, "Cash, end of year", f"={Hh('cash')}", "=D18+D17", "=E18+E17", "=F18+F17", bold=True)
r = 21
put(ws, f"A{r}", "BALANCE SHEET", font=BOLD, fill=SECT); r = 22
brow(22, "Cash & equivalents", "=C19", "=D19", "=E19", "=F19")
B["cash_bs"] = 22
brow(23, "Short-term investments", f"={Hh('short_term_investments')}",
     "=C23+D14", "=D23+E14", "=E23+F14")
brow(24, "Trade receivables, net", f"={Hh('receivables')}",
     f"=IS!D4*{DRV['ar_pct'][2026]}", f"=IS!E4*{DRV['ar_pct'][2027]}",
     f"=IS!F4*{DRV['ar_pct'][2028]}")
brow(25, "Other assets (incl. PP&E, goodwill, intangibles)", 
     f"=Historicals!G{HIST['total_assets']}-C22-C23-C24",
     "=C25+D12-D6", "=D25+E12-E6", "=E25+F12-F6")
brow(26, "Total assets", "=SUM(C22:C25)", "=SUM(D22:D25)", "=SUM(E22:E25)",
     "=SUM(F22:F25)", bold=True)
brow(28, "Total debt", f"={Hh('total_debt')}", f"={DRV['debt_total'][2026]}",
     "=D28+E16", "=E28+F16")
brow(29, "Deferred revenue (current)", f"={Hh('deferred_rev_current')}",
     f"=C29*(1+{DRV['defrev_g'][2026]})", f"=D29*(1+{DRV['defrev_g'][2027]})",
     f"=E29*(1+{DRV['defrev_g'][2028]})")
brow(30, "Other liabilities (incl. deferred revenue, non-current)",
     f"=Historicals!G{HIST['total_liabilities']}-C28-C29",
     "=C30*(IS!D4/IS!C4)", "=D30*(IS!E4/IS!D4)", "=E30*(IS!F4/IS!E4)")
brow(31, "Total liabilities", "=SUM(C28:C30)", "=SUM(D28:D30)", "=SUM(E28:E30)",
     "=SUM(F28:F30)", bold=True)
brow(32, "Total stockholders' equity", f"={Hh('equity')}",
     "=C32+D5+D7-D15", "=D32+E5+E7-E15", "=E32+F5+F7-F15")
brow(33, "Total liabilities & equity", "=C31+C32", "=D31+D32", "=E31+E32",
     "=F31+F32", bold=True)
brow(34, "BALANCE CHECK (assets − L&E; must be 0)", "=C26-C33", "=D26-D33",
     "=E26-E33", "=F26-F33", bold=True)
# re-map: rows used above in CF formulas referenced 28/31/34/35/36/40 — fix references:
# CF rows referenced: D31 (receivables) -> actual row 24; D35 deferred -> 29; D36 other liab -> 30;
# debt in IS referenced CF_BS rows 34 (total debt) -> actual 28; cash+STI row 40 -> rows 22+23.
# We therefore rewrite those formula cells with correct addresses.
put(ws, "D8", "=-(D24-C24)", MONEY); put(ws, "E8", "=-(E24-D24)", MONEY); put(ws, "F8", "=-(F24-E24)", MONEY)
put(ws, "D9", "=D29-C29", MONEY); put(ws, "E9", "=E29-D29", MONEY); put(ws, "F9", "=F29-E29", MONEY)
put(ws, "D10", "=D30-C30", MONEY); put(ws, "E10", "=E30-D30", MONEY); put(ws, "F10", "=F30-E30", MONEY)
put(ws, "D16", "=D28-C28", MONEY)
put(ws, "D18", "=C22", MONEY); put(ws, "E18", "=D22", MONEY); put(ws, "F18", "=E22", MONEY)
put(ws, "D19", "=D18+D17", MONEY); put(ws, "E19", "=E18+E17", MONEY); put(ws, "F19", "=F18+F17", MONEY)
put(ws, "C19", f"={Hh('cash')}", MONEY)
# IS interest/other-income references fix (debt rows 28; cash+STI rows 22+23)
ws_is = wb["IS"]
put(ws_is, "E12", f"='CF_BS'!D28*{DRV2['int_rate'][2027]}", MONEY)
put(ws_is, "F12", f"='CF_BS'!E28*{DRV2['int_rate'][2028]}", MONEY)
put(ws_is, "E13", f"=('CF_BS'!D22+'CF_BS'!D23)*{DRV2['cash_yield'][2027]}", MONEY)
put(ws_is, "F13", f"=('CF_BS'!E22+'CF_BS'!E23)*{DRV2['cash_yield'][2028]}", MONEY)
# share roll
r = 36
put(ws, f"A{r}", "SHARE COUNT ROLL (millions)", font=BOLD, fill=SECT)
brow(37, "Shares outstanding, beginning", None, f"=Historicals!G{HIST['shares_out']}",
     "=D40", "=E40", MONEY1)
brow(38, "Less: shares repurchased", None,
     f"=D15/{DRV['avg_bb_price'][2026]}", f"=E15/{DRV['avg_bb_price'][2027]}",
     f"=F15/{DRV['avg_bb_price'][2028]}", MONEY1)
brow(39, "Plus: shares issued for equity comp", None,
     f"={DRV['sbc_issue'][2026]}", f"={DRV['sbc_issue'][2027]}",
     f"={DRV['sbc_issue'][2028]}", MONEY1)
brow(40, "Shares outstanding, end of year", None, "=D37-D38+D39", "=E37-E38+E39",
     "=F37-F38+F39", MONEY1, bold=True)
brow(41, "Diluted weighted-average shares (BS roll)", None, None,
     "=(E37+E40)/2", "=(F37+F40)/2", MONEY1)
put(ws_is, "E17", "='CF_BS'!E41", MONEY1)
put(ws_is, "F17", "='CF_BS'!F41", MONEY1)
put(ws, "A43", "Cross-check: the roll implies ~381M shares at FY26 year-end vs 389.2M actually outstanding at Sep 18, 2026 with a full quarter of buybacks still to run — consistent. "
               "FY2025A cash-flow column shows reported figures; forecast columns are fully formula-driven.", font=Font(italic=True, size=9))

# ================================================================ VALUATION
ws = wb.create_sheet("Valuation")
put(ws, "A1", "Valuation — DCF (base / bear / bull / disruption), multiples cross-check, sensitivity", font=TITLEF)
V = {}
r = 3
put(ws, f"A{r}", "CURRENT MULTIPLES (at $247.09, Oct 19, 2026 close)", font=BOLD, fill=SECT); r += 1
def vrow(label, value, fmt=MONEY, key=None, bold=False):
    global r
    put(ws, f"A{r}", label, font=BOLD if bold else None)
    if value is not None: put(ws, f"B{r}", value, fmt, font=BOLD if bold else None)
    if key: V[key] = f"Valuation!$B${r}"
    r += 1
vrow("Market capitalization ($M)", f"={A['price']}*{A['shares']}", MONEY, "mktcap")
vrow("Plus: net debt at Q3 FY26 ($M)", f"={A['net_debt']}", MONEY, "nd")
vrow("Enterprise value ($M)", f"=B{r-2}+B{r-1}", MONEY, "ev", bold=True)
vrow("LTM revenue ($M)", f"=Historicals!J{HIST['revenue']}", MONEY)
vrow("LTM EBITDA ($M)", f"=Historicals!J{HIST['ebitda']}", MONEY, "ltm_ebitda")
vrow("LTM free cash flow ($M)", f"=Historicals!J{HIST['fcf']}", MONEY, "ltm_fcf")
vrow("EV / LTM EBITDA", f"={V['ev']}/{V['ltm_ebitda']}", MULT, "ev_ebitda")
vrow("Price / LTM FCF", f"={V['mktcap']}/{V['ltm_fcf']}", MULT, "p_fcf")
vrow("P / E (LTM GAAP)", f"={A['price']}/Historicals!J{HIST['eps']}", MULT, "pe")
r += 1
put(ws, f"A{r}", "OWN-HISTORY MULTIPLES — market cap at calendar year-end (companiesmarketcap.com) vs that fiscal year's fundamentals (EDGAR). FY ends late Nov/early Dec, so the calendar year-end cap is the nearest trading day after the fiscal year-end.", font=BOLD, fill=SECT); r += 1
hdr = r
for i, lab in enumerate(["Fiscal year", "Year-end mkt cap ($M)", "Net debt at FY-end ($M)", "EV ($M)", "EBITDA ($M)", "FCF ($M)", "EV/EBITDA", "P/FCF"]):
    put(ws, f"{get_column_letter(1+i)}{hdr}", lab, font=HDRF, fill=HDR)
r += 1
YECAP = {2021: 269350, 2022: 156450, 2023: 271630, 2024: 196540, 2025: 150080}
hist_first = r
for fy in range(2021, 2026):
    col = get_column_letter(3 + (fy - 2021))
    put(ws, f"A{r}", f"FY{fy}")
    put(ws, f"B{r}", YECAP[fy], MONEY, fill=INP)
    put(ws, f"C{r}", f"=Historicals!{col}{HIST['total_debt']}-Historicals!{col}{HIST['cash']}-Historicals!{col}{HIST['short_term_investments']}", MONEY)
    put(ws, f"D{r}", f"=B{r}+C{r}", MONEY)
    put(ws, f"E{r}", f"=Historicals!{col}{HIST['ebitda']}", MONEY)
    put(ws, f"F{r}", f"=Historicals!{col}{HIST['fcf']}", MONEY)
    put(ws, f"G{r}", f"=D{r}/E{r}", MULT)
    put(ws, f"H{r}", f"=B{r}/F{r}", MULT)
    r += 1
put(ws, f"A{r}", "Median FY21–FY25", font=BOLD)
put(ws, f"G{r}", f"=MEDIAN(G{hist_first}:G{r-1})", MULT, font=BOLD)
put(ws, f"H{r}", f"=MEDIAN(H{hist_first}:H{r-1})", MULT, font=BOLD)
V["med_ev_ebitda"] = f"Valuation!$G${r}"; V["med_p_fcf"] = f"Valuation!$H${r}"
V["fy25_p_fcf"] = f"Valuation!$H${r-1}"
r += 2
put(ws, f"A{r}", "DCF — BASE CASE (FCF FY2027E–FY2031E; FY27/28 from the three-statement model, FY29–31 fade at assumption growth rates)", font=BOLD, fill=SECT); r += 1
V["wacc"] = f"Valuation!$B${r+5}"; V["tg"] = f"Valuation!$B${r+6}"
for j, lab in enumerate(["FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]):
    put(ws, f"{get_column_letter(2+j)}{r}", lab, font=BOLD)
r += 1
fcf_row = r
put(ws, f"A{r}", "Free cash flow ($M)")
put(ws, f"B{r}", "='CF_BS'!E13", MONEY)
put(ws, f"C{r}", "='CF_BS'!F13", MONEY)
put(ws, f"D{r}", f"=C{r}*(1+{A['g29']})", MONEY)
put(ws, f"E{r}", f"=D{r}*(1+{A['g30']})", MONEY)
put(ws, f"F{r}", f"=E{r}*(1+{A['g31']})", MONEY)
r += 1
df_row = r
put(ws, f"A{r}", "Discount factor")
for j in range(5):
    put(ws, f"{get_column_letter(2+j)}{r}", f"=1/(1+{V['wacc']})^{j+1}", "0.000")
r += 1
pv_row = r
put(ws, f"A{r}", "PV of FCF")
for j in range(5):
    cl = get_column_letter(2+j)
    put(ws, f"{cl}{r}", f"={cl}{fcf_row}*{cl}{df_row}", MONEY)
r += 1
def vrow2(label, value, fmt=MONEY, key=None, bold=False):
    global r
    put(ws, f"A{r}", label, font=BOLD if bold else None)
    if value is not None: put(ws, f"B{r}", value, fmt, font=BOLD if bold else None)
    if key: V[key] = f"Valuation!$B${r}"
    r += 1
vrow2("PV of FY27E–31E FCF ($M)", f"=SUM(B{pv_row}:F{pv_row})", MONEY, "pv_fcf")
put(ws, f"A{r}", "WACC"); put(ws, f"B{r}", f"={A['wacc']}", PCT); r += 1
put(ws, f"A{r}", "Terminal growth"); put(ws, f"B{r}", f"={A['tg']}", PCT); r += 1
vrow2("Terminal value ($M)", f"=F{fcf_row}*(1+{V['tg']})/({V['wacc']}-{V['tg']})", MONEY, "tv")
vrow2("PV of terminal value ($M)", f"=B{r-1}*F{df_row}", MONEY, "pv_tv")
vrow2("Enterprise value ($M)", f"={V['pv_fcf']}+{V['pv_tv']}", MONEY, "dcf_ev", bold=True)
vrow2("Less: net debt ($M)", f"={A['net_debt']}", MONEY)
vrow2("Equity value ($M)", f"=B{r-2}-B{r-1}", MONEY, "eq_val")
vrow2("Diluted shares out (M)", f"={A['shares']}", MONEY1)
vrow2("TARGET PRICE", f"=B{r-2}/B{r-1}", PX, "target", bold=True)
vrow2("Current price", f"={A['price']}", PX)
vrow2("Implied upside", f"=B{r-2}/B{r-1}-1", PCT, "upside", bold=True)
r += 1
put(ws, f"A{r}", "SCENARIOS (same FCF base unless shocked; disruption uses its own FCF path)", font=BOLD, fill=SECT); r += 1
hdr = r
for i, lab in enumerate(["Scenario", "FCF FY27E", "FCF FY28E", "FCF FY29E", "FCF FY30E", "FCF FY31E", "WACC", "Term. g", "Equity value ($M)", "Per share", "Upside"]):
    put(ws, f"{get_column_letter(1+i)}{hdr}", lab, font=HDRF, fill=HDR)
r += 1
def scen(label, fcf_formulas, wacc_f, g_f):
    global r
    put(ws, f"A{r}", label)
    for j, f in enumerate(fcf_formulas):
        put(ws, f"{get_column_letter(2+j)}{r}", f, MONEY)
    put(ws, f"G{r}", wacc_f, PCT); put(ws, f"H{r}", g_f, PCT)
    pv = "+".join(f"{get_column_letter(2+j)}{r}/(1+$G{r})^{j+1}" for j in range(5))
    tv = f"F{r}*(1+H{r})/(G{r}-H{r})/(1+G{r})^5"
    put(ws, f"I{r}", f"={pv}+{tv}-{A['net_debt']}", MONEY)
    put(ws, f"J{r}", f"=I{r}/{A['shares']}", PX)
    put(ws, f"K{r}", f"=J{r}/{A['price']}-1", PCT)
    rr = r; r += 1
    return rr
base_f = [f"=B{fcf_row}", f"=C{fcf_row}", f"=D{fcf_row}", f"=E{fcf_row}", f"=F{fcf_row}"]
scen("Base", base_f, f"={V['wacc']}", f"={V['tg']}")
scen("Bear (FCF −18%, WACC +50bp, g 2.5%)",
     [f"=B{fcf_row}*(1+{A['bear_shock']})", f"=C{fcf_row}*(1+{A['bear_shock']})",
      f"=D{fcf_row}*(1+{A['bear_shock']})", f"=E{fcf_row}*(1+{A['bear_shock']})",
      f"=F{fcf_row}*(1+{A['bear_shock']})"],
     f"={V['wacc']}+{A['bear_wacc']}", f"={A['bear_g']}")
scen("Bull (FCF +8%, WACC −50bp, g 3.5%)",
     [f"=B{fcf_row}*(1+{A['bull_shock']})", f"=C{fcf_row}*(1+{A['bull_shock']})",
      f"=D{fcf_row}*(1+{A['bull_shock']})", f"=E{fcf_row}*(1+{A['bull_shock']})",
      f"=F{fcf_row}*(1+{A['bull_shock']})"],
     f"={V['wacc']}+{A['bull_wacc']}", f"={A['bull_g']}")
scen("AI disruption (market-implied bear)", [10800, 10600, 10200, 9900, 9700],
     f"={A['dis_wacc']}", f"={A['dis_g']}")
put(ws, f"A{r}", "Disruption FCF path is a direct input: revenue growth fades to low-single digits, FCF margin compresses toward ~30% as AI-native rivals take share. It values the stock near today's price — i.e., the market is already paying for this scenario.", font=Font(italic=True, size=9)); r += 2
put(ws, f"A{r}", "SENSITIVITY — target price at WACC (rows) × terminal growth (columns), base FCF path", font=BOLD, fill=SECT); r += 1
put(ws, f"B{r}", 0.025, PCT, font=BOLD); put(ws, f"C{r}", 0.03, PCT, font=BOLD); put(ws, f"D{r}", 0.035, PCT, font=BOLD)
sens_hdr = r; r += 1
for off in (-0.01, -0.005, 0.0, 0.005, 0.01):
    put(ws, f"A{r}", f"={V['wacc']}+{off}", PCT)
    for j, gc in enumerate("BCD"):
        pv = "+".join(f"{get_column_letter(2+k)}{fcf_row}/(1+$A{r})^{k+1}" for k in range(5))
        tv = f"F${fcf_row}*(1+{gc}${sens_hdr})/($A{r}-{gc}${sens_hdr})/(1+$A{r})^5"
        put(ws, f"{gc}{r}", f"=({pv}+{tv}-{A['net_debt']})/{A['shares']}", PX)
    r += 1
r += 1
put(ws, f"A{r}", "MULTIPLES CROSS-CHECKS", font=BOLD, fill=SECT); r += 1
vrow2("FY27E FCF per share", f"='CF_BS'!E13/IS!E17", PX, "fcfps27")
vrow2("Implied price at FY25's own (already de-rated) P/FCF", f"={V['fy25_p_fcf']}*{V['fcfps27']}", PX, "xcheck1")
vrow2("FY27E GAAP EPS", "=IS!E18", PX)
vrow2("Implied price at 20.0× FY27E GAAP EPS", f"=B{r-1}*20", PX, "xcheck2")
vrow2("EV / FY27E EBITDA at today's price", f"={V['ev']}/(IS!E11+'CF_BS'!E6)", MULT)
put(ws, f"A{r}", "Notes: valuation struck at Oct 20, 2026 using the latest reported balance sheet (Q3 FY26) and share count (Sep 18, 2026). "
                 "The stub Q4 FY26E free cash flow (~$3.8B, the seasonally largest quarter) is excluded from the DCF — a deliberate conservatism worth roughly $10/share. "
                 "SBC is treated as a real cost throughout (it is inside FCF); nothing is added back.", font=Font(italic=True, size=9))
for col, w in zip("ABCDEFGHIJK", [46, 14, 14, 14, 14, 14, 12, 12, 15, 11, 10]):
    ws.column_dimensions[col].width = w

# ================================================================ COVER
ws = wb.create_sheet("Cover", 0)
put(ws, "A1", "Adobe Inc. (NASDAQ: ADBE) — Equity Pitch", font=Font(bold=True, size=16, color=NAVY))
put(ws, "A2", "LONG — the market is pricing an AI disruption that the filings do not show", font=Font(bold=True, size=12))
put(ws, "A3", "Archit Chawla · October 20, 2026 · All figures from SEC EDGAR filings unless noted; $ in millions except per-share")
put(ws, "A5", "Recommendation", font=SUBF)
put(ws, "A6", "LONG / BUY", font=Font(bold=True, size=13))
put(ws, "A7", "Target price"); put(ws, "B7", f"={V['target']}", PX, font=BOLD)
put(ws, "A8", "Price (close, Oct 19, 2026)"); put(ws, "B8", f"={A['price']}", PX)
put(ws, "A9", "Implied upside"); put(ws, "B9", f"={V['upside']}", PCT, font=BOLD)
put(ws, "A10", "Street consensus"); put(ws, "B10", "Hold, avg target $272–284 (stockanalysis.com / MarketBeat, Oct 20, 2026)")
put(ws, "A12", "Snapshot (LTM to Q3 FY2026)", font=SUBF)
put(ws, "A13", "Revenue (LTM)"); put(ws, "B13", f"=Historicals!J{HIST['revenue']}", MONEY)
put(ws, "A14", "Free cash flow (LTM)"); put(ws, "B14", f"=Historicals!J{HIST['fcf']}", MONEY)
put(ws, "A15", "Buybacks (FY2025)"); put(ws, "B15", f"=Historicals!G{HIST['buybacks']}", MONEY)
put(ws, "A16", "Shares outstanding (Sep 18, 2026)"); put(ws, "B16", f"={A['shares']}", MONEY1)
put(ws, "A17", "EV / LTM EBITDA"); put(ws, "B17", f"={V['ev_ebitda']}", MULT)
put(ws, "A18", "P / LTM FCF"); put(ws, "B18", f"={V['p_fcf']}", MULT)
put(ws, "A20", "Contents: Historicals (EDGAR inputs) · Assumptions (every driver, with rationale) · IS · CF_BS (linked three-statement forecast, FY2026E–FY2028E) · Valuation (DCF, scenarios, sensitivity, multiples history).")
put(ws, "A21", "Method in one line: FY2026E is built as 9M FY2026 actuals + a Q4 stub grown off Q4 FY2025; FY2027E–28E are driver-based; the balance sheet balances by construction and cash ties to the cash-flow statement in every year.")
put(ws, "A22", "Disclaimer: independent student research project prepared for educational purposes. Not investment advice. Verify all figures against the underlying SEC filings before acting on them.")
ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 60

out = os.path.join(HERE, "ADBE_3statement_model.xlsx")
wb.save(out)
print("saved", out)
