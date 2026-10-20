#!/usr/bin/env python3
"""Pull Adobe Inc. (CIK 0000796343) financials from SEC EDGAR company facts.

Rerunnable. Saves the raw company-facts JSON plus a tidy CSV keyed on PERIOD
END DATE into data/ with the pull date stamped in the filename.

Source: https://data.sec.gov/api/xbrl/companyfacts/CIK0000796343.json

Fiscal-year convention: Adobe's fiscal year ends on the Friday nearest
Nov 30 and is labelled by the calendar year in which it ends (FY2025 ended
Nov 28, 2025). NOTE: the `fy` field in EDGAR company facts is the FILING's
fiscal year (a 10-K carries comparatives for three years), so periods here
are identified by end date, never by the `fy` field.
"""
import csv
import datetime as dt
import json
import os
import sys
import urllib.request

CIK = "0000796343"
UA = {"User-Agent": "ArchitChawla Research archit.26.chawla@gmail.com"}
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")

DURATION_CONCEPTS = {
    "revenue": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"],
    "cost_of_revenue": ["CostOfRevenue"],
    "gross_profit": ["GrossProfit"],
    "rd": ["ResearchAndDevelopmentExpenseSoftwareExcludingAcquiredInProcessCost",
           "ResearchAndDevelopmentExpense"],
    "sm": ["SellingAndMarketingExpense"],
    "ga": ["GeneralAndAdministrativeExpense"],
    "operating_expenses": ["OperatingExpenses"],
    "operating_income": ["OperatingIncomeLoss"],
    "interest_expense": ["InterestExpense"],
    "other_nonoperating": ["OtherNonoperatingIncomeExpense"],
    "nonoperating_net": ["NonoperatingIncomeExpense"],
    "pretax_income": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
    "tax": ["IncomeTaxExpenseBenefit"],
    "net_income": ["NetIncomeLoss"],
    "da": ["DepreciationDepletionAndAmortization",
           "DepreciationAmortizationAndAccretionNet"],
    "sbc": ["ShareBasedCompensation"],
    "cfo": ["NetCashProvidedByUsedInOperatingActivities"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "buybacks": ["PaymentsForRepurchaseOfCommonStock"],
    "dividends": ["PaymentsOfDividendsCommonStock", "PaymentsOfDividends"],
}
INSTANT_CONCEPTS = {
    "cash": ["CashAndCashEquivalentsAtCarryingValue"],
    "short_term_investments": ["ShortTermInvestments"],
    "receivables": ["AccountsReceivableNetCurrent"],
    "total_assets": ["Assets"],
    "total_liabilities": ["Liabilities"],
    "equity": ["StockholdersEquity"],
    "long_term_debt": ["LongTermDebt"],
    "current_debt": ["DebtCurrent"],
    "deferred_rev_current": ["ContractWithCustomerLiabilityCurrent",
                             "DeferredRevenueCurrent"],
}
SHARE_CONCEPTS = {
    "shares_diluted": ["WeightedAverageNumberOfDilutedSharesOutstanding"],
    "shares_basic": ["WeightedAverageNumberOfSharesOutstandingBasic"],
}


def fetch(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def pick(facts, taxonomy, names, unit_pref=("USD", "shares")):
    tax = facts.get("facts", {}).get(taxonomy, {})
    for n in names:
        if n in tax:
            units = tax[n].get("units", {})
            for u in unit_pref:
                if u in units:
                    return n, units[u]
            if units:
                return n, units[next(iter(units))]
    return None, []


def fy_label(end):
    """Adobe FY label from a period end date: Nov/Dec ends -> that year's FY;
    Feb/May/Aug ends belong to the FY ending the following Nov/Dec."""
    d = dt.date.fromisoformat(end)
    if d.month >= 11:
        return d.year
    return d.year if d.month <= 2 else d.year + 1  # quarters sit inside FY of Nov end


def duration_rows(item, concept, obs, kind):
    """kind: 'annual' (350-380d, 10-K), 'quarter' (80-100d, 10-Q),
    'ytd9m' (250-290d, 10-Q) for cash-flow LTM builds."""
    out = {}
    for o in obs:
        if not o.get("start"):
            continue
        form = o.get("form", "")
        days = (dt.date.fromisoformat(o["end"]) - dt.date.fromisoformat(o["start"])).days
        if kind == "annual" and not (form.startswith("10-K") and 350 <= days <= 380):
            continue
        if kind == "quarter" and not (form.startswith("10-Q") and 80 <= days <= 100):
            continue
        if kind == "ytd9m" and not (form.startswith("10-Q") and 250 <= days <= 290):
            continue
        key = o["end"]
        if key not in out or o.get("filed", "") > out[key].get("filed", ""):
            out[key] = o
    rows = []
    for end, o in sorted(out.items()):
        rows.append({"item": item, "concept": concept, "kind": kind,
                     "fy": fy_label(end), "end": end, "start": o.get("start"),
                     "val": o.get("val"), "form": o.get("form"),
                     "filed": o.get("filed")})
    return rows


def instant_rows(item, concept, obs, ends):
    by_end = {}
    for o in obs:
        if o.get("form", "")[:4] not in ("10-K", "10-Q"):
            continue
        if o["end"] not in by_end or o.get("filed", "") > by_end[o["end"]].get("filed", ""):
            by_end[o["end"]] = o
    rows = []
    for end in ends:
        if end in by_end:
            o = by_end[end]
            rows.append({"item": item, "concept": concept, "kind": "instant",
                         "fy": fy_label(end), "end": end, "start": "",
                         "val": o.get("val"), "form": o.get("form"),
                         "filed": o.get("filed")})
    return rows


def main():
    os.makedirs(DATA, exist_ok=True)
    today = dt.date.today().isoformat()
    facts = fetch(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json")
    raw_path = os.path.join(DATA, f"ADBE_companyfacts_{today}.json")
    with open(raw_path, "w") as f:
        json.dump(facts, f)
    print(f"raw facts -> {raw_path} ({os.path.getsize(raw_path)/1e6:.1f} MB)")
    print("entity:", facts.get("entityName"))

    # fiscal year-end dates from the revenue series
    _, rev_obs = pick(facts, "us-gaap", DURATION_CONCEPTS["revenue"])
    fy_ends = sorted({r["end"] for r in duration_rows("revenue", "Revenues", rev_obs, "annual")})
    q_ends = sorted({r["end"] for r in duration_rows("revenue", "Revenues", rev_obs, "quarter")})
    snapshot_ends = sorted(set(fy_ends[-6:] + q_ends[-4:]))
    print("FY ends:", fy_ends[-6:], "| latest quarters:", q_ends[-4:])

    rows = []
    for item, names in DURATION_CONCEPTS.items():
        concept, obs = pick(facts, "us-gaap", names)
        if concept is None:
            print(f"  MISSING {item}")
            continue
        rows += duration_rows(item, concept, obs, "annual")
        rows += duration_rows(item, concept, obs, "quarter")
        rows += duration_rows(item, concept, obs, "ytd9m")
    for item, names in SHARE_CONCEPTS.items():
        concept, obs = pick(facts, "us-gaap", names, ("shares",))
        if concept is None:
            continue
        rows += duration_rows(item, concept, obs, "annual")
        rows += duration_rows(item, concept, obs, "quarter")
    for item, names in INSTANT_CONCEPTS.items():
        concept, obs = pick(facts, "us-gaap", names)
        if concept is None:
            print(f"  MISSING {item}")
            continue
        rows += instant_rows(item, concept, obs, snapshot_ends)
    # shares outstanding (dei, instant)
    concept, obs = pick(facts, "dei", ["EntityCommonStockSharesOutstanding"], ("shares",))
    if concept:
        by_end = {}
        for o in obs:
            if o["end"] not in by_end or o.get("filed", "") > by_end[o["end"]].get("filed", ""):
                by_end[o["end"]] = o
        for end, o in sorted(by_end.items()):
            if end >= "2024-01-01":
                rows.append({"item": "shares_outstanding", "concept": concept,
                             "kind": "instant", "fy": fy_label(end), "end": end,
                             "start": "", "val": o.get("val"),
                             "form": o.get("form"), "filed": o.get("filed")})

    # Derived: IS amortization of acquired intangibles = total operating
    # expenses - R&D - S&M - G&A (the AmortizationOfIntangibleAssets tag is a
    # cash-flow concept and does not tie to the income-statement line).
    from collections import defaultdict
    vals = defaultdict(dict)
    for r in rows:
        if r["item"] in ("operating_expenses", "rd", "sm", "ga") and \
                r["kind"] in ("annual", "quarter"):
            vals[(r["kind"], r["end"])][r["item"]] = r
    for (kind, end), d in sorted(vals.items()):
        if len(d) == 4:
            base = d["operating_expenses"]
            rows.append({"item": "amort_intangibles",
                         "concept": "Derived: OperatingExpenses - R&D - S&M - G&A",
                         "kind": kind, "fy": base["fy"], "end": end,
                         "start": base["start"],
                         "val": d["operating_expenses"]["val"] - d["rd"]["val"]
                                - d["sm"]["val"] - d["ga"]["val"],
                         "form": base["form"], "filed": base["filed"]})
    csv_path = os.path.join(DATA, f"ADBE_edgar_series_{today}.csv")
    with open(csv_path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["item", "concept", "kind", "fy",
                                          "end", "start", "val", "form", "filed"])
        w.writeheader()
        w.writerows(rows)
    print(f"series -> {csv_path} ({len(rows)} rows)")


if __name__ == "__main__":
    sys.exit(main())
