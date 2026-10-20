# Adobe Equity Pitch — LONG, target $500

An end-to-end equity pitch built from primary sources: every historical figure comes from Adobe's SEC EDGAR filings (company facts API), and the valuation runs off a linked three-statement Excel model — not a screenshot of someone else's estimates.

**The call (Oct 20, 2026):** LONG Adobe (NASDAQ: ADBE) at $247.09 (close, Oct 19, 2026), target **$500** (+102%). The market is pricing an AI-disruption scenario (my disruption case values the stock at $242); the filings show revenue re-accelerating to +12.9% YoY, a 40.8% LTM free-cash-flow margin, and a buyback that retired 13.9% of the share count between January 2024 and September 2026.

## Files

| File | What it is |
|---|---|
| `ADBE_pitch_memo.md` / `.pdf` | The pitch: variant perception, thesis pillars, valuation, catalysts, risks, verification notes |
| `ADBE_3statement_model.xlsx` | Linked three-statement model (live formulas): Historicals · Assumptions · IS · CF_BS · Valuation · Cover |
| `pull_data.py` | Rerunnable EDGAR pull → `data/ADBE_edgar_series_<date>.csv` + raw JSON |
| `build_model.py` | Rerunnable model builder (reads the CSV, writes the workbook) |
| `data/` | Raw EDGAR company facts, the tidy series CSV, and `market_data_2026-10-20.json` (price/beta/market-cap sources) |

## Headline numbers

- LTM to Q3 FY2026: revenue $25,970M · EBITDA $10,037M · free cash flow $10,592M (40.8% margin) · net income $7,284M
- Today: 9.7× EV/EBITDA, 9.1× P/FCF — vs Adobe's own FY2021–25 year-end medians of 25.6× and 25.0×
- Forecast: revenue $26.71B (FY26E) → $29.65B (FY27E) → $32.77B (FY28E), an 11.3% FY25–28E CAGR with growth *faded* from the FY26 run-rate; FCF $11.24B → $12.91B → $14.37B
- DCF: WACC 11.0% (Rf 5.1%, beta 1.42, ERP 4.5%, terminal growth 3.0%) → base $500 · bear $369 · bull $609 · disruption $242

## How to rerun

```bash
python3 pull_data.py     # refreshes data/ from SEC EDGAR (no API key needed)
python3 build_model.py   # rebuilds ADBE_3statement_model.xlsx from the CSV
```

## Verification (done at build, Oct 20, 2026)

The workbook was recalculated in LibreOffice headless and checked: FY2025 revenue ($23,769M), net income ($7,130M) and operating cash flow ($10,031M) tie exactly to EDGAR; the balance-sheet check row is 0 in every forecast year; balance-sheet cash equals the cash-flow roll; the DCF target equals the center cell of the WACC × terminal-growth sensitivity grid; the model's FY2026E non-GAAP EPS proxy ($24.64) sits just above Adobe's $24.45–24.50 guidance, as expected for a pre-tax add-back proxy. Two data traps worth knowing: EDGAR's `fy` field is the *filing's* fiscal year (periods here are keyed on end dates), and the income-statement amortization line must be derived as OperatingExpenses − R&D − S&M − G&A ($157M in FY2025) — the identically named tag is a cash-flow concept and breaks the tie-out by ~$150M. The share price is cited from MarketBeat/Finnhub/stockanalysis.com (Oct 19, 2026 close) because Stooq's CSV endpoint was blocked by a JavaScript browser check; sources are recorded in `data/market_data_2026-10-20.json`.

## Limitations

Beta is used as reported (1.42, elevated by the de-rating itself) — conservative for the target. The multiples history uses calendar year-end market caps against late-November fiscal year-ends (approximate by a few weeks of drift). FY2026E embeds one estimated quarter (Q4; Adobe reports Dec 9, 2026). SBC is charged through share-count dilution, not as a cash cost in FCF — stated openly in the memo.

*Independent student research project by Archit Chawla, prepared for educational purposes. Not investment advice.*
