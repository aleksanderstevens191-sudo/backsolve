"""Stage 4: turn the raw public downloads (data/stage4/raw/, from scripts/stage4_fetch.mjs) into one
clean, auditable file: data/stage4/companies.json.

Every number in that file keeps its source: for SEC numbers the XBRL concept name, the filing's
accession number, form and filing date; for prices the date and the source file; for rates the
date and the series. The workbook generator (scripts/make_sec_dcf.py) only reads this file.

How SEC numbers are picked (plain English)
-------------------------------------------
SEC's "company facts" API lists every number a company has tagged in its filings, under a concept
name like `Revenues`. Companies don't all use the same names (Apple tags revenue as
`RevenueFromContractWithCustomerExcludingAssessedTax`, others as `Revenues`), and names changed
over the years, so for each line we try a short list of concepts.

* Full-year numbers (revenue, operating income, ...): only facts from annual reports (10-K or
  10-K/A) covering 350-380 days. When several filings report the same year (a later 10-K repeats
  the two previous years), we take the most recently filed one, so restated figures win.
* Year-end balances (cash, debt): facts dated exactly at the fiscal year end, from 10-Ks.
* Revenue: when a year is tagged under more than one revenue concept, we take the largest, since
  the others are usually parts of the total. The concept used is recorded for each year.

Usage: python3 scripts/stage4_extract.py
"""
import csv, datetime as dt, gzip, io, json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
RAW = os.path.join(ROOT, "data", "stage4", "raw")
OUT = os.path.join(ROOT, "data", "stage4", "companies.json")

COMPANIES = [  # same list as scripts/stage4_fetch.mjs
    ("AAPL", 320193, "Apple"), ("MSFT", 789019, "Microsoft"), ("NVDA", 1045810, "NVIDIA"),
    ("AMZN", 1018724, "Amazon"), ("GOOGL", 1652044, "Alphabet"), ("META", 1326801, "Meta Platforms"),
    ("WMT", 104169, "Walmart"), ("JNJ", 200406, "Johnson & Johnson"), ("PG", 80424, "Procter & Gamble"),
    ("KO", 21344, "Coca-Cola"),
]
YEARS = 10

DURATION = {
    "revenue": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
                "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet", "SalesRevenueGoodsNet"],
    "operatingIncome": ["OperatingIncomeLoss"],
    "pretaxIncome": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                     "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "incomeTax": ["IncomeTaxExpenseBenefit"],
    "netIncome": ["NetIncomeLoss"],
    "interestExpense": ["InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt", "InterestAndDebtExpense"],
    "depreciationAmortization": ["DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet",
                                 "DepreciationAndAmortization", "DepreciationAmortizationAndOther"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "operatingCashFlow": ["NetCashProvidedByUsedInOperatingActivities",
                          "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
}
INSTANT = {
    "cash": ["CashAndCashEquivalentsAtCarryingValue"],
    "securitiesCurrent": ["MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent", "ShortTermInvestments"],
    "securitiesNoncurrent": ["MarketableSecuritiesNoncurrent", "AvailableForSaleSecuritiesDebtSecuritiesNoncurrent"],
    "longTermDebtTotal": ["LongTermDebt", "LongTermDebtAndCapitalLeaseObligationsIncludingCurrentMaturities"],
    "longTermDebtNoncurrent": ["LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations"],
    "longTermDebtCurrent": ["LongTermDebtCurrent", "LongTermDebtAndCapitalLeaseObligationsCurrent"],
    "shortTermBorrowings": ["ShortTermBorrowings", "CommercialPaper"],
}


def load_gz(rel):
    with gzip.open(os.path.join(RAW, rel + ".gz")) as f:
        return f.read()


def days(a, b):
    return (dt.date.fromisoformat(b) - dt.date.fromisoformat(a)).days


def src(f, concept):
    return {"concept": concept, "accn": f["accn"], "form": f["form"], "filed": f["filed"],
            **({"start": f["start"]} if "start" in f else {}), "end": f["end"]}


def annual(gaap, concept, instant):
    """end date -> latest-filed 10-K fact for this concept."""
    out = {}
    units = gaap.get(concept, {}).get("units", {})
    for f in units.get("USD", []):
        if not f.get("form", "").startswith("10-K"):
            continue
        if instant != ("start" not in f):
            continue
        if not instant and not (350 <= days(f["start"], f["end"]) <= 380):
            continue
        prev = out.get(f["end"])
        if prev is None or f["filed"] > prev["filed"]:
            out[f["end"]] = f
    return out


def pick(gaap, concepts, end, instant, largest=False):
    found = []
    for c in concepts:
        f = annual(gaap, c, instant).get(end)
        if f is not None:
            found.append((f, c))
            if not largest:
                break
    if not found:
        return None
    f, c = max(found, key=lambda x: x[0]["val"]) if largest else found[0]
    return {"value": f["val"], **src(f, c)}


def company(ticker, cik, name, price_date):
    facts = json.loads(load_gz(f"sec/CIK{cik:010d}.json"))
    gaap = facts["facts"]["us-gaap"]
    dei = facts["facts"].get("dei", {})
    # Fiscal year ends = the year ends that have revenue in a 10-K, newest 10 (+1 for growth).
    ends = set()
    for c in DURATION["revenue"]:
        ends |= set(annual(gaap, c, False))
    ends = sorted(e for e in ends if e <= price_date)[-(YEARS):]
    for a, b in zip(ends, ends[1:]):
        if not 350 <= days(a, b) <= 380:
            raise SystemExit(f"{ticker}: fiscal year ends {a} and {b} are not a year apart")
    years = []
    for e in ends:
        y = {"end": e}
        for k, cs in DURATION.items():
            y[k] = pick(gaap, cs, e, False, largest=(k == "revenue"))
        years.append(y)
    last = ends[-1]
    bal = {"end": last}
    for k, cs in INSTANT.items():
        bal[k] = pick(gaap, cs, last, True)
    bal["previousEnd"] = ends[-2]
    prev = {k: pick(gaap, cs, ends[-2], True) for k, cs in INSTANT.items() if k.startswith(("longTerm", "shortTerm"))}
    bal["previousDebt"] = prev

    # Shares outstanding: the cover-page number (dei) from the most recently filed report on or
    # before the price date. Companies with several share classes (Alphabet) report one number per
    # class in the same filing; those are added up.
    sh = [f for f in dei.get("EntityCommonStockSharesOutstanding", {}).get("units", {}).get("shares", []) if f["filed"] <= price_date]
    latest = max(f["filed"] for f in sh)
    same = [f for f in sh if f["filed"] == latest]
    accn = same[0]["accn"]
    same = [f for f in same if f["accn"] == accn]
    shares = {"value": sum(f["val"] for f in same), "concept": "dei:EntityCommonStockSharesOutstanding", "accn": accn,
              "form": same[0]["form"], "filed": latest, "end": same[0]["end"], "classes": len(same)}
    dil = [f for f in gaap.get("WeightedAverageNumberOfDilutedSharesOutstanding", {}).get("units", {}).get("shares", []) if f["filed"] <= price_date]
    if dil:
        d = max(dil, key=lambda f: (f["end"], f["filed"]))
        shares["dilutedWeightedCheck"] = {"value": d["val"], **src(d, "WeightedAverageNumberOfDilutedSharesOutstanding")}
    return {"ticker": ticker, "cik": cik, "name": name, "secEntityName": facts.get("entityName"),
            "source": f"data/stage4/raw/sec/CIK{cik:010d}.json.gz (https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json)",
            "years": years, "balance": bal, "shares": shares}


def yahoo_series(rel):
    j = json.loads(load_gz(rel))["chart"]["result"][0]
    q = j["indicators"]["quote"][0]
    adj = j["indicators"].get("adjclose", [{}])[0].get("adjclose") or q["close"]
    tz = j["meta"].get("gmtoffset", -18000)
    rows = []
    for t, c, a in zip(j["timestamp"], q["close"], adj):
        if c is None or a is None:
            continue
        d = dt.datetime.fromtimestamp(t + tz, dt.timezone.utc).date().isoformat()  # exchange-local date
        rows.append((d, c, a))
    return rows, j["meta"]


def main():
    man = json.load(open(os.path.join(RAW, "manifest.json")))
    by_file = {f["file"]: f for f in man["files"] if not f.get("failed")}
    fetched = dt.datetime.fromisoformat(man["startedAt"].replace("Z", "+00:00"))
    fetched_ny = (fetched - dt.timedelta(hours=5)).date().isoformat()  # New York date (EST; close enough to pick "before today")

    # Prices: the last full trading day before the download day (a bar for the download day itself
    # may be a half-finished day).
    prices = {}
    for t, _, _ in COMPANIES + [("GSPC", 0, "S&P 500")]:
        rel = f"prices/{t}.json"
        rows, meta = yahoo_series(rel)
        rows = [r for r in rows if r[0] < fetched_ny]
        prices[t] = {"rows": rows, "meta": {k: meta.get(k) for k in ("symbol", "currency", "exchangeName", "instrumentType")}, "file": f"data/stage4/raw/{rel}.gz",
                     "url": by_file[f"data/stage4/raw/{rel}.gz"]["url"].split("?")[0], "fetchedAt": by_file[f"data/stage4/raw/{rel}.gz"]["fetchedAt"]}
    price_date = min(p["rows"][-1][0] for p in prices.values())

    # Month-end prices for beta: 61 month ends -> 60 monthly returns, ending with the last full
    # month before the price date.
    def month_ends(rows):
        m = {}
        for d, c, a in rows:
            m[d[:7]] = (d, a)
        return m
    last_full_month = (dt.date.fromisoformat(price_date).replace(day=1) - dt.timedelta(days=1)).isoformat()[:7]
    idx_me = month_ends(prices["GSPC"]["rows"])
    months = sorted(k for k in idx_me if k <= last_full_month)[-61:]

    # 10-year Treasury yield (FRED DGS10, percent): the last value on or before the price date.
    rf = None
    for row in csv.reader(io.StringIO(load_gz("rates/DGS10.csv").decode())):
        if row and row[0][:2] in ("19", "20") and row[1] not in (".", "") and row[0] <= price_date:
            rf = (row[0], float(row[1]) / 100)
    # Damodaran's implied equity risk premium (histimpl.xls): the latest year's FCFE-based value.
    erp = damodaran_erp()

    out = {
        "generatedBy": "scripts/stage4_extract.py", "rawManifest": "data/stage4/raw/manifest.json", "downloadedAt": man["startedAt"],
        "priceDate": price_date,
        "riskFreeRate": {"value": rf[1], "date": rf[0], "series": "DGS10, 10-year US Treasury constant-maturity yield",
                         "source": "data/stage4/raw/rates/DGS10.csv.gz (https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10)"},
        "equityRiskPremium": erp,
        "betaMonths": months,
        "index": {"ticker": "^GSPC", "name": "S&P 500 (price index)", "file": prices["GSPC"]["file"], "monthEnds": [idx_me[m] for m in months]},
        "companies": [],
    }
    for t, cik, name in COMPANIES:
        c = company(t, cik, name, price_date)
        p = prices[t]
        last = [r for r in p["rows"] if r[0] <= price_date][-1]
        me = month_ends(p["rows"])
        c["price"] = {"date": last[0], "close": last[1], "source": p["file"], "url": p["url"], "fetchedAt": p["fetchedAt"],
                      "note": "Yahoo Finance daily close (split-adjusted, not dividend-adjusted)", **p["meta"]}
        c["beta"] = {"monthEnds": [me[m] for m in months], "note": "month-end close adjusted for splits and dividends (Yahoo adjclose)"}
        c["sources"] = c.pop("source")
        out["companies"].append(c)
        miss = [f"{y['end']} {k}" for y in c["years"] for k in DURATION if y[k] is None]
        print(f"{t}: FY {c['years'][0]['end']}..{c['years'][-1]['end']}, price {last[1]:.2f} on {last[0]}, shares {c['shares']['value']:,} ({c['shares']['classes']} class(es)); missing: {', '.join(miss) or 'none'}")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, "w"), indent=1)
    print(f"price date {price_date}; 10y Treasury {rf[1]:.4f} on {rf[0]}; ERP {erp['value']:.4f} ({erp['asOf']})\nwrote {os.path.relpath(OUT, ROOT)}")


def damodaran_erp():
    import xlrd
    book = xlrd.open_workbook(file_contents=load_gz("rates/histimpl.xls"))
    for sh in book.sheets():
        for r in range(min(sh.nrows, 40)):
            row = [str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)]
            if row and row[0].lower().startswith("year"):
                cols = [c for c, h in enumerate(row) if "implied" in h.lower() and "fcfe" in h.lower()]
                if not cols:
                    cols = [c for c, h in enumerate(row) if "implied" in h.lower() and "premium" in h.lower()]
                if not cols:
                    continue
                col = cols[0]
                last = None
                for rr in range(r + 1, sh.nrows):
                    y, v = sh.cell_value(rr, 0), sh.cell_value(rr, col)
                    if isinstance(y, float) and isinstance(v, float) and 1900 < y < 2100:
                        last = (int(y), v)
                return {"value": last[1], "asOf": f"start of {last[0] + 1} (row for year-end {last[0]})", "column": row[col], "sheet": sh.name,
                        "source": "data/stage4/raw/rates/histimpl.xls.gz (Aswath Damodaran, NYU Stern, https://pages.stern.nyu.edu/~adamodar/pc/datasets/histimpl.xls)"}
    raise SystemExit("could not find the implied premium column in histimpl.xls")


if __name__ == "__main__":
    main()
