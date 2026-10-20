# Adobe Inc. (NASDAQ: ADBE) — Equity Pitch

**Recommendation: LONG** · **Target price: $500** · Price: $247.09 (close, Oct 19, 2026) · **Implied upside: +102%**
Archit Chawla
Market cap $96.2B · EV $96.9B · Shares out 389.2M (Sep 18, 2026) · Street: Hold, average target $272–284

## Variant perception

The market prices Adobe as an AI loser whose creative franchise melts away: at $247 the stock is worth almost exactly my disruption scenario ($242), in which revenue growth fades to low-single digits and the FCF margin compresses toward 30%. The filings show the opposite trajectory — growth re-accelerating toward 13%, a ~41% FCF margin, and a buyback retiring roughly 6% of the share count every year.

## Business snapshot

Adobe is the incumbent in digital content creation and experience: Digital Media (Creative Cloud, Document Cloud/Acrobat) sells subscriptions to creators, knowledge workers and enterprises; Digital Experience sells the marketing/analytics stack to the same enterprise base. The model is ~90% gross margin, subscription-led, with $7.09B of current deferred revenue on the Q3 FY2026 balance sheet — revenue already contracted, not yet recognized. No dividend (verified: no dividend concept in Adobe's EDGAR filings); all capital return is buybacks. Essentially no net debt ($0.72B at Q3 FY26).

## Thesis pillars

1. **Growth is re-accelerating, not rolling over.** Revenue grew +12.0%, +12.7% and +12.9% YoY in Q1–Q3 FY2026, up from +10.5% in FY2025; LTM revenue is $26.0B. Deferred revenue grew +11.1% YoY to $7.09B (Q3 FY26) — the forward book is compounding at double digits, which is not what a disrupted franchise looks like.
2. **The cash machine is intact.** LTM free cash flow is $10.59B, a 40.8% margin on revenue; FY2025 operating cash flow was $10.03B on capex of just 0.8% of revenue. AI investment runs through opex and cost of revenue — it is already inside these numbers, not a capex wall still to come.
3. **The buyback is a second return engine.** Adobe repurchased $11.28B in FY2025 — 114% of FCF — and $35.7B over five years. Shares outstanding fell from 452.0M (Jan 2024) to 389.2M (Sep 2026), −13.9%. In my model, diluted weighted shares fall from 427M (FY25) to 344M (FY28E), so GAAP EPS compounds from $18.97 (FY26E) to $29.67 (FY28E) even though net income grows a more modest 6–19% a year. (SBC of ~$2.1B a year is charged against this through the share-count roll — it is a real cost, not added back.)
4. **The multiple prices a funeral.** Today: 9.7× EV/LTM EBITDA and 9.1× price/LTM FCF. Adobe's own fiscal-year-end multiples over FY2021–FY2025 (EDGAR fundamentals, calendar year-end market caps) were 15.7–40.6× EV/EBITDA and 15.2–39.1× P/FCF, medians 25.6× and 25.0×. Even applying FY2025's already de-rated 15.2× P/FCF to FY2027E FCF per share gives $534.

## Valuation

DCF on FY2027E–FY2031E free cash flow from the linked three-statement model (FY26E revenue $26.71B, FY27E $29.65B, FY28E $32.77B — an 11.3% FY25–28E revenue CAGR with growth faded, not accelerated; FCF $11.24B / $12.91B / $14.37B). WACC 11.0%: risk-free 5.1% (house convention), beta 1.42 (Finnhub/MarketBeat, Oct 20, 2026), ERP 4.5% → cost of equity 11.5%; after-tax cost of debt 3.3%; terminal growth 3.0%.

| Scenario | Assumptions | Value/share | vs $247.09 |
|---|---|---|---|
| Bull | FCF +8%, WACC 10.5%, g 3.5% | $609 | +147% |
| **Base (target)** | Model FCF, WACC 11.0%, g 3.0% | **$500** | **+102%** |
| Bear | FCF −18%, WACC 11.5%, g 2.5% | $369 | +49% |
| AI disruption | Growth fades to LSD, FCF margin → ~30%, WACC 11.5%, g 1.5% | $242 | −2% |

Cross-checks cluster with the DCF: 20.0× FY2027E GAAP EPS ($24.36) = $487; FY2025's own P/FCF on FY2027E FCF/share = $534. Sensitivity (WACC ±100bp, g 2.5–3.5%): $425–609. The gap to the Street's $272–284 is the variant perception in one line: consensus sits between my disruption and bear cases — it is paying for partial disruption. The Q4 FY2026 stub FCF (~$3.8B, ≈$10/share) is excluded from the DCF, a deliberate conservatism.

## Catalysts

- **Dec 9, 2026 — Q4 FY2026 results and FY2027 guidance.** Adobe guided Q4 non-GAAP EPS to $6.30–6.35; my model's FY2026 non-GAAP EPS proxy is $24.64 vs the $24.45–24.50 full-year guide — the print either validates the "filings over fears" read or breaks it.
- **Buyback pace in the Q4 print** (~$2.6B assumed): at $247 every $1B retires ~1% of the company.
- **AI monetization disclosures** (Firefly generation volumes, Acrobat AI attach) across FY2027 prints — the specific evidence the market says it is waiting for.

## Risks — and what would change my call

- **AI-native competition** (Canva, Figma, generative image/video tools) taking Creative seats rather than expanding the market. This is the disruption case; it is a real risk, and it is also already the price.
- **Seat-growth slowdown.** Tripwire: deferred revenue growth below +5% YoY for two consecutive quarters, or revenue growth below +9% — pillar 1 is broken and I would exit to the disruption valuation.
- **Guidance credibility.** A cut to FY2027 guidance on Dec 9 breaks the thesis's hinge; the stock's de-rating began with guide disappointment, and one more resets the debate.
- **SBC creep.** Stock comp is 8.2% of revenue (FY2025). If it rises toward 10%, the buyback has to run harder just to stand still.
- **Multiple stays sentiment-driven.** Beta is 1.42; even a right thesis will be volatile. Position sizing should assume the $190 June 2026 low can be revisited.

## Methodology & verification

- All historicals are SEC EDGAR company facts (CIK 0000796343), pulled Oct 20, 2026 (`pull_data.py`); periods are keyed on period end dates because EDGAR's `fy` field is the filing's year, not the period's. Adobe's fiscal year ends the Friday nearest Nov 30 (FY2025 ended Nov 28, 2025).
- The Excel model (`ADBE_3statement_model.xlsx`, live formulas) was recalculated in LibreOffice and verified: FY2025 revenue ($23,769M), net income ($7,130M) and operating cash flow ($10,031M) tie exactly to EDGAR; the balance sheet balances (check = 0) in every forecast year and balance-sheet cash equals the cash-flow roll throughout; the DCF target equals the center cell of the sensitivity grid ($499.91).
- FY2026E is built as 9M FY2026 actuals + a Q4 stub grown off derived Q4 FY2025 (FY minus 9M) — every stub growth rate is an assumption cell with a stated rationale.
- One derivation to know: the income-statement amortization line is derived as OperatingExpenses − R&D − S&M − G&A ($157M in FY2025) because the identically named EDGAR tag is a cash-flow concept that does not tie to the income statement (using it breaks the tie-out by ~$150M).

## Limitations

Beta is taken as reported (1.42 — itself inflated by the de-rating), which raises the WACC and lowers the target; a normalized beta would raise the target materially. The multiples history uses calendar year-end market caps (companiesmarketcap.com) against late-November fiscal year-ends, so it is approximate by a few weeks of price drift, and uses year-end snapshots rather than averages. FY2026E still embeds one estimated quarter. The model is consolidated — EDGAR company facts do not support a segment-level build. SBC is charged through dilution (share count) rather than as a cash cost inside FCF; that is the standard treatment and it is stated, not hidden.

*Independent student research project, prepared for educational purposes. Not investment advice. Figures should be verified against the underlying SEC filings before any investment decision.*
