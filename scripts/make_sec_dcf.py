"""Stage 4: build one real Excel valuation workbook per company from data/stage4/companies.json.

Each workbook (workbooks/sec/<TICKER>.xlsx) has live Excel formulas all the way from the SEC
numbers to the value per share and the LBO's IRR:

  Data     10 fiscal years from SEC filings (hard-coded, exactly as reported, in US dollars), with
           growth, margins and the actual 5-year revenue growth as formulas.
  Balance  cash, marketable securities, debt and shares outstanding at the last fiscal year end.
  Beta     60 monthly returns of the stock and the S&P 500; beta computed by formula.
  WACC     every input with its source next to it.
  DCF      10-year free-cash-flow DCF + Gordon-growth terminal value -> value per share.
  LBO      a 5-year leveraged buyout at a premium to the share price -> IRR.
  Sources  one row per hard-coded number: XBRL concept, accession number, form, filing date.

openpyxl writes formulas without computed values. The files are then recalculated and saved by a
spreadsheet program so they carry cached values: LibreOffice here (scripts/lo_recalc.py), and
Microsoft Excel on Alek's Mac (scripts/excel_recalc_sec.sh), which is the ground truth.

Also writes workbooks/sec/cells.json: where the key cells are, for scripts/stage4.ts.

Usage: python3 scripts/make_sec_dcf.py [--no-recalc]
"""
import datetime as dt, json, os, subprocess, sys, tempfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
OUT = os.path.join(ROOT, "workbooks", "sec")
DATA = os.environ.get("STAGE4_DATA") or os.path.join(ROOT, "data", "stage4", "companies.json")

INPUT = Font(color="0000CC")          # blue = hard-coded input (banker convention)
BOLD = Font(bold=True)
OUTFILL = PatternFill("solid", fgColor="FFF2CC")
YEARS = 10
COLS = [get_column_letter(2 + i) for i in range(YEARS)]  # B..K, oldest -> newest
LAST = COLS[-1]

# Assumptions that are judgement calls, not data. Each one is written next to its cell.
TERMINAL_GROWTH = 0.025
BLUME = (0.67, 0.33)
NWC_PCT = 0.0
LBO = {"premium": 0.30, "fees": 0.02, "leverage": 5.0, "spread": 0.03, "years": 5, "target_irr": 0.20}


def serial(iso):
    return (dt.date.fromisoformat(iso) - dt.date(1899, 12, 30)).days


def put(ws, addr, v, fmt=None, inp=False, bold=False):
    ws[addr] = v
    if fmt:
        ws[addr].number_format = fmt
    if inp:
        ws[addr].font = INPUT
    if bold:
        ws[addr].font = BOLD


def build(c, m, path):
    wb = Workbook()
    srcs = []  # (sheet!cell, what, source dict or text)

    # ---------------- Data ----------------
    d = wb.active; d.title = "Data"
    put(d, "A1", f"{c['name']} ({c['ticker']}): 10 fiscal years from SEC EDGAR company facts (US$)", bold=True)
    put(d, "A2", f"Source: {c['sources']}. Every hard-coded number is listed with its XBRL concept and filing on the Sources sheet.")
    put(d, "A4", "Fiscal year ended", bold=True)
    lines = [("revenue", "Revenue"), ("operatingIncome", "Operating income (EBIT)"), ("depreciationAmortization", "Depreciation & amortization"),
             ("capex", "Capital expenditures"), ("operatingCashFlow", "Operating cash flow"), ("netIncome", "Net income"),
             ("pretaxIncome", "Pre-tax income"), ("incomeTax", "Income tax expense"), ("interestExpense", "Interest expense")]
    row = {k: 5 + i for i, (k, _) in enumerate(lines)}
    years = c["years"]
    # EBIT fallback: a company that doesn't report an operating income line (or not for the last 3
    # years) gets EBIT = pre-tax income + interest expense, by formula, for every year.
    ebit_proxy = any(y["operatingIncome"] is None for y in years[-3:])
    missing = []
    for k, label in lines:
        put(d, f"A{row[k]}", label + (" = pre-tax income + interest expense (no operating income line in the filings)" if k == "operatingIncome" and ebit_proxy else ""))
    for col, y in zip(COLS, years):
        put(d, f"{col}4", serial(y["end"]), "yyyy-mm-dd", inp=True)
        srcs.append((f"Data!{col}4", "fiscal year end", f"period end of the revenue fact ({y['revenue']['accn']})"))
        for k, label in lines:
            a = f"{col}{row[k]}"
            if k == "operatingIncome" and ebit_proxy:
                d[a] = f"={col}{row['pretaxIncome']}+{col}{row['interestExpense']}"
                d[a].number_format = "#,##0"
                continue
            f = y[k]
            if f is None:
                missing.append(f"{label} FY{y['end'][:4]}")
                continue
            put(d, a, f["value"], "#,##0", inp=True)
            srcs.append((f"Data!{a}", f"{label}, FY ended {y['end']}", f))
    r = 15
    put(d, f"A{r}", "Ratios (formulas)", bold=True)
    ratio = {}
    def ratio_row(key, label, fmt, f, need, start=0):
        nonlocal r
        r += 1; ratio[key] = r
        put(d, f"A{r}", label)
        for i, col in enumerate(COLS):
            if i < start:
                continue
            ok = all(years[i - (1 if n.endswith("_prev") else 0)][n.replace("_prev", "")] is not None or (n.startswith("operatingIncome") and ebit_proxy) for n in need)
            if ok:
                d[f"{col}{r}"] = "=" + f(col, COLS[i - 1] if i else None)
                d[f"{col}{r}"].number_format = fmt
    R = row
    ratio_row("growth", "Revenue growth", "0.0%", lambda c_, p: f"{c_}{R['revenue']}/{p}{R['revenue']}-1", ["revenue", "revenue_prev"], start=1)
    ratio_row("margin", "Operating margin", "0.0%", lambda c_, p: f"{c_}{R['operatingIncome']}/{c_}{R['revenue']}", ["operatingIncome", "revenue"])
    ratio_row("da", "D&A % revenue", "0.0%", lambda c_, p: f"{c_}{R['depreciationAmortization']}/{c_}{R['revenue']}", ["depreciationAmortization", "revenue"])
    ratio_row("capex", "Capex % revenue", "0.0%", lambda c_, p: f"{c_}{R['capex']}/{c_}{R['revenue']}", ["capex", "revenue"])
    ratio_row("tax", "Effective tax rate", "0.0%", lambda c_, p: f"{c_}{R['incomeTax']}/{c_}{R['pretaxIncome']}", ["incomeTax", "pretaxIncome"])
    ratio_row("fcf", "Free cash flow (operating cash flow − capex)", "#,##0", lambda c_, p: f"{c_}{R['operatingCashFlow']}-{c_}{R['capex']}", ["operatingCashFlow", "capex"])
    ratio_row("fcfm", "Free cash flow margin", "0.0%", lambda c_, p: f"{c_}{ratio['fcf']}/{c_}{R['revenue']}", ["operatingCashFlow", "capex", "revenue"])
    r += 2
    first5 = COLS[-6]
    put(d, f"A{r}", "Actual 5-year revenue growth (compound annual)", bold=True)
    d[f"B{r}"] = f"=({LAST}{R['revenue']}/{first5}{R['revenue']})^(1/5)-1"; d[f"B{r}"].number_format = "0.00%"; d[f"B{r}"].fill = OUTFILL
    put(d, f"C{r}", f"FY{years[-6]['end'][:4]} to FY{years[-1]['end'][:4]} ({years[-6]['end']} to {years[-1]['end']})")
    cagr5 = f"B{r}"; r += 1
    put(d, f"A{r}", "Actual 9-year revenue growth (compound annual)")
    d[f"B{r}"] = f"=({LAST}{R['revenue']}/{COLS[0]}{R['revenue']})^(1/9)-1"; d[f"B{r}"].number_format = "0.00%"
    cagr9 = f"B{r}"
    d.column_dimensions["A"].width = 52
    for col in COLS:
        d.column_dimensions[col].width = 17

    def avg3(key):
        return f"AVERAGE(Data!{COLS[-3]}{ratio[key]}:Data!{LAST}{ratio[key]})"

    # ---------------- Balance ----------------
    b = wb.create_sheet("Balance")
    bal = c["balance"]
    put(b, "A1", f"Balance sheet at fiscal year end {bal['end']} (latest 10-K) and shares outstanding", bold=True)
    items = [("cash", "Cash and cash equivalents"), ("securitiesCurrent", "Marketable securities, current"),
             ("securitiesNoncurrent", "Marketable securities, non-current"),
             ("longTermDebtTotal", "Long-term debt incl. current portion"), ("longTermDebtNoncurrent", "Long-term debt, non-current"),
             ("longTermDebtCurrent", "Long-term debt, current portion"), ("shortTermBorrowings", "Short-term borrowings / commercial paper")]
    brow = {}
    put(b, "B3", "This year", bold=True); put(b, "C3", "Previous year", bold=True)
    put(b, "B2", serial(bal["end"]), "yyyy-mm-dd", inp=True); put(b, "C2", serial(bal["previousEnd"]), "yyyy-mm-dd", inp=True)
    put(b, "A2", "Date")
    for i, (k, label) in enumerate(items):
        rr = 4 + i; brow[k] = rr
        put(b, f"A{rr}", label)
        f = bal.get(k)
        put(b, f"B{rr}", f["value"] if f else 0, "#,##0", inp=True)
        srcs.append((f"Balance!B{rr}", label, f if f else "not tagged at this date in the 10-K: 0"))
        if k.startswith(("longTerm", "shortTerm")):
            fp = bal["previousDebt"].get(k)
            put(b, f"C{rr}", fp["value"] if fp else 0, "#,##0", inp=True)
            srcs.append((f"Balance!C{rr}", label + " (previous year)", fp if fp else "not tagged at this date in the 10-K: 0"))
    # Total debt: the "incl. current portion" figure when tagged, else non-current + current portion.
    rr = 12
    put(b, f"A{rr}", "Cash and marketable securities", bold=True)
    b[f"B{rr}"] = f"=B{brow['cash']}+B{brow['securitiesCurrent']}+B{brow['securitiesNoncurrent']}"; b[f"B{rr}"].number_format = "#,##0"
    cash_cell = f"Balance!B{rr}"
    rr += 1
    put(b, f"A{rr}", "Total debt", bold=True)
    for col in "BC":
        lt = "longTermDebtTotal"
        has_total = (bal.get(lt) if col == "B" else bal["previousDebt"].get(lt)) is not None
        b[f"{col}{rr}"] = (f"={col}{brow[lt]}+{col}{brow['shortTermBorrowings']}" if has_total else
                           f"={col}{brow['longTermDebtNoncurrent']}+{col}{brow['longTermDebtCurrent']}+{col}{brow['shortTermBorrowings']}")
        b[f"{col}{rr}"].number_format = "#,##0"
    put(b, f"D{rr}", "Long-term debt incl. current portion (or non-current + current portion when the total isn't tagged) + short-term borrowings. Leases not included.")
    debt_cell, debt_prev = f"Balance!B{rr}", f"Balance!C{rr}"
    rr += 1
    put(b, f"A{rr}", "Net debt (debt − cash and securities)")
    b[f"B{rr}"] = f"=B{rr - 1}-B{rr - 2}"; b[f"B{rr}"].number_format = "#,##0"
    rr += 2
    sh = c["shares"]
    put(b, f"A{rr}", "Shares outstanding", bold=True)
    put(b, f"B{rr}", sh["value"], "#,##0", inp=True)
    put(b, f"C{rr}", f"cover page of the {sh['form']} filed {sh['filed']} (as of {sh['end']})" + (f", {sh['classes']} share classes added" if sh["classes"] > 1 else ""))
    srcs.append((f"Balance!B{rr}", "Shares outstanding", sh))
    shares_cell = f"Balance!B{rr}"
    b.column_dimensions["A"].width = 44
    for col in "BC":
        b.column_dimensions[col].width = 20

    # ---------------- Beta ----------------
    bt = wb.create_sheet("Beta")
    put(bt, "A1", f"Beta: {c['ticker']} vs the S&P 500, 60 monthly returns", bold=True)
    put(bt, "A2", "Month-end prices from Yahoo Finance (stock: adjusted for splits and dividends; index: S&P 500 price index ^GSPC).")
    for col, h in zip("ABCDE", ["Month end", c["ticker"], "S&P 500", f"{c['ticker']} return", "S&P 500 return"]):
        put(bt, f"{col}4", h, bold=True)
    n = len(m["betaMonths"])
    for i in range(n):
        rr = 5 + i
        (ds, ps), (_, pi) = c["beta"]["monthEnds"][i], m["index"]["monthEnds"][i]
        put(bt, f"A{rr}", serial(ds), "yyyy-mm-dd", inp=True)
        put(bt, f"B{rr}", ps, "0.0000", inp=True)
        put(bt, f"C{rr}", pi, "0.00", inp=True)
        if i:
            bt[f"D{rr}"] = f"=B{rr}/B{rr - 1}-1"; bt[f"E{rr}"] = f"=C{rr}/C{rr - 1}-1"
            bt[f"D{rr}"].number_format = bt[f"E{rr}"].number_format = "0.00%"
    srcs.append((f"Beta!B5:B{4 + n}", f"{c['ticker']} month-end adjusted close", c["price"]["source"]))
    srcs.append((f"Beta!C5:C{4 + n}", "S&P 500 month-end close", m["index"]["file"]))
    lo, hi = 6, 4 + n
    D, E = f"D{lo}:D{hi}", f"E{lo}:E{hi}"
    rr = hi + 2
    put(bt, f"A{rr}", "Number of returns"); bt[f"B{rr}"] = f"=COUNT({D})"; nret = f"B{rr}"
    put(bt, f"A{rr + 1}", "Covariance (stock, index)"); bt[f"B{rr + 1}"] = f"=SUMPRODUCT({D},{E})/{nret}-AVERAGE({D})*AVERAGE({E})"
    put(bt, f"A{rr + 2}", "Variance (index)"); bt[f"B{rr + 2}"] = f"=SUMPRODUCT({E},{E})/{nret}-AVERAGE({E})^2"
    put(bt, f"A{rr + 3}", "Raw beta = covariance / variance", bold=True); bt[f"B{rr + 3}"] = f"=B{rr + 1}/B{rr + 2}"
    bt[f"B{rr + 3}"].number_format = "0.000"
    raw_beta = f"Beta!B{rr + 3}"
    bt.column_dimensions["A"].width = 34

    # ---------------- WACC ----------------
    w = wb.create_sheet("WACC")
    put(w, "A1", "Discount rate (WACC): every input and its source", bold=True)
    rf, erp = m["riskFreeRate"], m["equityRiskPremium"]
    wrows = [
        ("rf", "Risk-free rate (10-year US Treasury yield)", rf["value"], "0.00%", True, f"FRED DGS10 on {rf['date']}: {rf['source']}"),
        ("erp", "Equity risk premium", erp["value"], "0.00%", True, f"Damodaran implied ERP, {erp['asOf']}, column '{erp['column']}': {erp['source']}"),
        ("raw", "Raw beta", f"={raw_beta}", "0.000", False, "Beta sheet: 60 monthly returns vs the S&P 500"),
        ("bw", "Blume weight on raw beta", BLUME[0], "0.00", True, "Assumption (Blume 1971): adjusted beta = 0.67 × raw + 0.33 × 1, pulls betas toward 1"),
        ("beta", "Adjusted beta", None, "0.000", False, "formula"),
        ("ke", "Cost of equity = risk-free + beta × ERP", None, "0.00%", False, "CAPM"),
        ("int", "Interest expense, last fiscal year", f"=Data!{LAST}{R['interestExpense']}", "#,##0", False, "Data sheet (SEC)"),
        ("avgdebt", "Average debt, last two year ends", f"=({debt_cell}+{debt_prev})/2", "#,##0", False, "Balance sheet"),
        ("kd", "Pre-tax cost of debt", None, "0.00%", False, "interest expense ÷ average debt; never below the risk-free rate"),
        ("tax", "Tax rate (3-year effective)", None, "0.00%", False, "income tax ÷ pre-tax income, last 3 fiscal years added up (Data sheet)"),
        ("price", "Share price", c["price"]["close"], "0.00", True, f"Yahoo Finance close on {c['price']['date']}: {c['price']['source']}"),
        ("shares", "Shares outstanding", f"={shares_cell}", "#,##0", False, "Balance sheet (SEC cover page)"),
        ("E", "Equity market value", None, "#,##0", False, "price × shares"),
        ("D", "Debt (book value)", f"={debt_cell}", "#,##0", False, "Balance sheet (SEC)"),
        ("adj", "WACC adjustment (normally 0)", 0, "0.00%", True, "For sensitivity tests only: set to +1% or −1% to move the WACC by 1 percentage point"),
        ("wacc", "WACC", None, "0.00%", False, "E/(D+E) × cost of equity + D/(D+E) × cost of debt × (1 − tax) + adjustment"),
    ]
    W = {k: f"B{4 + i}" for i, (k, *_) in enumerate(wrows)}
    formulas = {
        "beta": f"={W['bw']}*{W['raw']}+(1-{W['bw']})",
        "ke": f"={W['rf']}+{W['beta']}*{W['erp']}",
        "kd": f"=IF({W['avgdebt']}>0,MAX({W['rf']},{W['int']}/{W['avgdebt']}),{W['rf']})",
        "tax": f"=SUM(Data!{COLS[-3]}{R['incomeTax']}:Data!{LAST}{R['incomeTax']})/SUM(Data!{COLS[-3]}{R['pretaxIncome']}:Data!{LAST}{R['pretaxIncome']})",
        "E": f"={W['price']}*{W['shares']}",
        "wacc": f"={W['E']}/({W['D']}+{W['E']})*{W['ke']}+{W['D']}/({W['D']}+{W['E']})*{W['kd']}*(1-{W['tax']})+{W['adj']}",
    }
    if years[-1]["interestExpense"] is None:
        formulas["kd"] = f"={W['rf']}"
        wrows[[k for k, *_ in wrows].index("kd")] = ("kd", "Pre-tax cost of debt", None, "0.00%", False, "interest expense not tagged in the latest 10-K, so the risk-free rate is used (assumption)")
        wrows[[k for k, *_ in wrows].index("int")] = ("int", "Interest expense, last fiscal year", 0, "#,##0", True, "not tagged in the latest 10-K")
    for i, (k, label, v, fmt, inp, note) in enumerate(wrows):
        rr = 4 + i
        put(w, f"A{rr}", label, bold=(k == "wacc"))
        put(w, f"B{rr}", formulas.get(k, v), fmt, inp=inp)
        put(w, f"C{rr}", note)
        if inp:
            srcs.append((f"WACC!B{rr}", label, note))
    w[W["wacc"]].fill = OUTFILL
    w.column_dimensions["A"].width = 42; w.column_dimensions["B"].width = 18

    # ---------------- DCF ----------------
    f = wb.create_sheet("DCF")
    put(f, "A1", f"{c['name']}: 10-year discounted cash flow, valued at fiscal year end {years[-1]['end']} and rolled forward to {c['price']['date']}", bold=True)
    hist5 = (years[-1]["revenue"]["value"] / years[-6]["revenue"]["value"]) ** 0.2 - 1
    arows = [
        ("g", "Revenue growth, each of the next 10 years", round(hist5, 6), "0.00%", True, f"INPUT. Base case = actual 5-year growth (Data!{cagr5}), rounded. Backsolve solves for this."),
        ("tg", "Terminal growth (after year 10)", TERMINAL_GROWTH, "0.00%", True, "Assumption: long-run nominal growth, below the risk-free rate"),
        ("margin", "Operating margin", f"={avg3('margin')}", "0.00%", False, "average of the last 3 fiscal years (Data)"),
        ("tax", "Tax rate", f"=WACC!{W['tax']}", "0.00%", False, "WACC sheet"),
        ("da", "D&A % revenue", f"={avg3('da')}", "0.00%", False, "average of the last 3 fiscal years (Data)"),
        ("capex", "Capex % revenue", f"={avg3('capex')}", "0.00%", False, "average of the last 3 fiscal years (Data)"),
        ("nwc", "Working capital investment, % of revenue increase", NWC_PCT, "0.00%", True, "Assumption: not modelled (0). SEC working-capital tags differ by company."),
        ("wacc", "WACC", f"=WACC!{W['wacc']}", "0.00%", False, "WACC sheet"),
        ("ke", "Cost of equity", f"=WACC!{W['ke']}", "0.00%", False, "WACC sheet (used to roll the value forward to the price date)"),
        ("price", "Share price", f"=WACC!{W['price']}", "0.00", False, f"Yahoo Finance close on {c['price']['date']}"),
        ("pdate", "Price date", serial(c["price"]["date"]), "yyyy-mm-dd", True, "date of that close"),
        ("fyend", "Last fiscal year end", f"=Data!{LAST}4", "yyyy-mm-dd", False, "Data sheet"),
    ]
    A = {k: f"B{3 + i}" for i, (k, *_) in enumerate(arows)}
    for i, (k, label, v, fmt, inp, note) in enumerate(arows):
        rr = 3 + i
        put(f, f"A{rr}", label); put(f, f"B{rr}", v, fmt, inp=inp); put(f, f"C{rr}", note)
    f[A["g"]].fill = OUTFILL
    srcs.append((f"DCF!{A['pdate']}", "Price date", c["price"]["source"]))
    P = [get_column_letter(2 + i) for i in range(YEARS + 1)]  # B = last actual, C..L = years 1..10
    hdr = 17
    labels = ["Year", "Revenue", "EBIT", "NOPAT = EBIT × (1 − tax)", "D&A", "Capex", "Working capital investment",
              "Free cash flow = NOPAT + D&A − capex − working capital", "Discount factor", "Present value"]
    Y = {lab: hdr + i for i, lab in enumerate(labels)}
    for lab, rr in Y.items():
        put(f, f"A{rr}", lab)
    for i, col in enumerate(P):
        put(f, f"{col}{Y['Year']}", i)
        if i == 0:
            f[f"{col}{Y['Revenue']}"] = f"=Data!{LAST}{R['revenue']}"
            continue
        p = P[i - 1]
        f[f"{col}{Y['Revenue']}"] = f"={p}{Y['Revenue']}*(1+{A['g'].replace('B', '$B$')})"
        f[f"{col}{Y['EBIT']}"] = f"={col}{Y['Revenue']}*$B${A['margin'][1:]}"
        f[f"{col}{Y['NOPAT = EBIT × (1 − tax)']}"] = f"={col}{Y['EBIT']}*(1-$B${A['tax'][1:]})"
        f[f"{col}{Y['D&A']}"] = f"={col}{Y['Revenue']}*$B${A['da'][1:]}"
        f[f"{col}{Y['Capex']}"] = f"={col}{Y['Revenue']}*$B${A['capex'][1:]}"
        f[f"{col}{Y['Working capital investment']}"] = f"=({col}{Y['Revenue']}-{p}{Y['Revenue']})*$B${A['nwc'][1:]}"
        fcf = Y["Free cash flow = NOPAT + D&A − capex − working capital"]
        f[f"{col}{fcf}"] = f"={col}{Y['NOPAT = EBIT × (1 − tax)']}+{col}{Y['D&A']}-{col}{Y['Capex']}-{col}{Y['Working capital investment']}"
        f[f"{col}{Y['Discount factor']}"] = f"=1/(1+$B${A['wacc'][1:]})^{col}{Y['Year']}"
        f[f"{col}{Y['Present value']}"] = f"={col}{fcf}*{col}{Y['Discount factor']}"
    for lab, rr in Y.items():
        for col in P:
            f[f"{col}{rr}"].number_format = "0" if lab == "Year" else ("0.0000" if lab == "Discount factor" else "#,##0")
    fcf = Y["Free cash flow = NOPAT + D&A − capex − working capital"]
    v0 = Y["Present value"] + 2
    vrows = [
        ("pv", "Sum of present values, years 1–10", f"=SUM(C{Y['Present value']}:L{Y['Present value']})", "#,##0"),
        ("tv", "Terminal value at year 10", f"=L{fcf}*(1+{A['tg']})/({A['wacc']}-{A['tg']})", "#,##0"),
        ("pvtv", "Present value of terminal value", f"=B{v0 + 1}*L{Y['Discount factor']}", "#,##0"),
        ("ev", "Enterprise value at fiscal year end", f"=B{v0}+B{v0 + 2}", "#,##0"),
        ("cash", "+ cash and marketable securities", f"={cash_cell}", "#,##0"),
        ("debt", "− debt", f"={debt_cell}", "#,##0"),
        ("eq", "Equity value at fiscal year end", f"=B{v0 + 3}+B{v0 + 4}-B{v0 + 5}", "#,##0"),
        ("sh", "Shares outstanding", f"={shares_cell}", "#,##0"),
        ("vfy", "Value per share at fiscal year end", f"=B{v0 + 6}/B{v0 + 7}", "0.00"),
        ("t", "Years from fiscal year end to price date", f"=({A['pdate']}-{A['fyend']})/365.25", "0.000"),
        ("vps", "Value per share at the price date", f"=B{v0 + 8}*(1+{A['ke']})^B{v0 + 9}", "0.00"),
        ("mkt", "Market share price", f"={A['price']}", "0.00"),
        ("gap", "Model value ÷ market price − 1", f"=B{v0 + 10}/B{v0 + 11}-1", "0.0%"),
    ]
    V = {}
    for i, (k, label, fx, fmt) in enumerate(vrows):
        rr = v0 + i; V[k] = f"B{rr}"
        put(f, f"A{rr}", label, bold=(k == "vps")); put(f, f"B{rr}", fx, fmt)
    f[V["vps"]].fill = OUTFILL
    put(f, f"C{V['t'][1:]}", "The equity value is computed at the last fiscal year end (when the balance sheet is dated) and grown at the cost of equity to the price date.")
    f.column_dimensions["A"].width = 58
    for col in P:
        f.column_dimensions[col].width = 16

    # ---------------- LBO ----------------
    L = wb.create_sheet("LBO")
    put(L, "A1", f"{c['name']}: hypothetical 5-year leveraged buyout at a premium to the {c['price']['date']} share price", bold=True)
    put(L, "A2", "Illustrative only: no buyout of a company this size is being suggested. Debt is paid down with all free cash flow; interest is charged on the opening balance each year.")
    lrows = [
        ("prem", "Offer premium over the market price", LBO["premium"], "0.0%", True, "Assumption: typical take-private premium"),
        ("offer", "Offer price per share", f"=WACC!{W['price']}*(1+B4)", "0.00", False, "market price × (1 + premium)"),
        ("eqp", "Equity purchase price", f"=B5*{shares_cell}", "#,##0", False, "offer price × shares"),
        ("refi", "+ existing debt refinanced", f"={debt_cell}", "#,##0", False, "Balance sheet"),
        ("cashu", "− existing cash and securities used", f"={cash_cell}", "#,##0", False, "Balance sheet"),
        ("ev", "Purchase enterprise value", "=B6+B7-B8", "#,##0", False, ""),
        ("ebitda0", "EBITDA, last fiscal year", f"=Data!{LAST}{R['operatingIncome']}+Data!{LAST}{R['depreciationAmortization']}", "#,##0", False, "EBIT + D&A (Data)"),
        ("mult", "Entry EV / EBITDA", "=B9/B10", "0.0x", False, ""),
        ("fees", "Transaction fees, % of EV", LBO["fees"], "0.0%", True, "Assumption"),
        ("lev", "New debt, × EBITDA", LBO["leverage"], "0.0x", True, "Assumption"),
        ("debt0", "New debt", "=B13*B10", "#,##0", False, ""),
        ("eq0", "Sponsor equity", "=B9*(1+B12)-B14", "#,##0", False, "EV + fees − new debt"),
        ("spread", "Spread over the risk-free rate on new debt", LBO["spread"], "0.00%", True, "Assumption"),
        ("rate", "Interest rate on new debt", f"=WACC!{W['rf']}+B16", "0.00%", False, "10-year Treasury + spread"),
        ("g", "Revenue growth, each year", round(hist5, 6), "0.00%", True, f"INPUT. Base case = actual 5-year growth (Data!{cagr5}), rounded"),
        ("m0", "EBITDA margin, last fiscal year", f"=B10/Data!{LAST}{R['revenue']}", "0.00%", False, ""),
        ("dm", "Margin improvement by year 5 (percentage points of EBITDA margin, phased in evenly)", 0, "0.00%", True, "INPUT. Base case 0. Backsolve solves for this."),
        ("dmult", "Exit multiple minus entry multiple", 0, "0.0x", True, "Assumption: exit at the entry multiple"),
        ("target", "Target IRR", LBO["target_irr"], "0.0%", True, "Goal for the backsolve"),
    ]
    for i, (k, label, v, fmt, inp, note) in enumerate(lrows):
        rr = 4 + i
        put(L, f"A{rr}", label); put(L, f"B{rr}", v, fmt, inp=inp); put(L, f"C{rr}", note)
    LB = {k: f"B{4 + i}" for i, (k, *_) in enumerate(lrows)}
    L[LB["dm"]].fill = OUTFILL
    Q = [get_column_letter(2 + i) for i in range(LBO["years"] + 1)]  # B = entry, C..G = years 1..5
    h = 25
    ll = ["Year", "Revenue", "EBITDA margin", "EBITDA", "D&A", "EBIT", "Interest (on opening debt)", "Taxes",
          "Capex", "Free cash flow after interest", "Debt (closing)", "Cash (closing)", "Cash flow to sponsor"]
    LL = {lab: h + i for i, lab in enumerate(ll)}
    e0 = h + len(ll) + 1  # exit block starts here; exit equity value is at row e0 + 3
    for lab, rr in LL.items():
        put(L, f"A{rr}", lab)
    for i, col in enumerate(Q):
        put(L, f"{col}{LL['Year']}", i)
        if i == 0:
            L[f"{col}{LL['Revenue']}"] = f"=Data!{LAST}{R['revenue']}"
            L[f"{col}{LL['Debt (closing)']}"] = f"={LB['debt0']}"
            L[f"{col}{LL['Cash (closing)']}"] = 0
            L[f"{col}{LL['Cash flow to sponsor']}"] = f"=-{LB['eq0']}"
            continue
        p = Q[i - 1]
        L[f"{col}{LL['Revenue']}"] = f"={p}{LL['Revenue']}*(1+$B${LB['g'][1:]})"
        L[f"{col}{LL['EBITDA margin']}"] = f"=$B${LB['m0'][1:]}+$B${LB['dm'][1:]}*{col}{LL['Year']}/{LBO['years']}"
        L[f"{col}{LL['EBITDA']}"] = f"={col}{LL['Revenue']}*{col}{LL['EBITDA margin']}"
        L[f"{col}{LL['D&A']}"] = f"={col}{LL['Revenue']}*DCF!$B${A['da'][1:]}"
        L[f"{col}{LL['EBIT']}"] = f"={col}{LL['EBITDA']}-{col}{LL['D&A']}"
        L[f"{col}{LL['Interest (on opening debt)']}"] = f"={p}{LL['Debt (closing)']}*$B${LB['rate'][1:]}"
        L[f"{col}{LL['Taxes']}"] = f"=MAX(0,{col}{LL['EBIT']}-{col}{LL['Interest (on opening debt)']})*DCF!$B${A['tax'][1:]}"
        L[f"{col}{LL['Capex']}"] = f"={col}{LL['Revenue']}*DCF!$B${A['capex'][1:]}"
        fc = LL["Free cash flow after interest"]
        L[f"{col}{fc}"] = f"={col}{LL['EBITDA']}-{col}{LL['Interest (on opening debt)']}-{col}{LL['Taxes']}-{col}{LL['Capex']}"
        L[f"{col}{LL['Debt (closing)']}"] = f"=MAX(0,{p}{LL['Debt (closing)']}-{col}{fc})"
        L[f"{col}{LL['Cash (closing)']}"] = f"={p}{LL['Cash (closing)']}+MAX(0,{col}{fc}-{p}{LL['Debt (closing)']})"
        L[f"{col}{LL['Cash flow to sponsor']}"] = 0 if i < LBO["years"] else f"=B{e0 + 3}"
    for lab, rr in LL.items():
        for col in Q:
            L[f"{col}{rr}"].number_format = "0" if lab == "Year" else ("0.00%" if "margin" in lab else "#,##0")
    G = Q[-1]
    erows = [
        ("xebitda", "Exit EBITDA (year 5)", f"={G}{LL['EBITDA']}", "#,##0"),
        ("xmult", "Exit multiple", f"={LB['mult']}+{LB['dmult']}", "0.0x"),
        ("xev", "Exit enterprise value", f"=B{e0}*B{e0 + 1}", "#,##0"),
        ("xeq", "Exit equity value", f"=B{e0 + 2}-{G}{LL['Debt (closing)']}+{G}{LL['Cash (closing)']}", "#,##0"),
        ("_", "", None, None),
        ("moic", "Money multiple (MOIC)", f"=B{e0 + 3}/{LB['eq0']}", "0.00x"),
        ("irr", "IRR", f"=IRR(B{LL['Cash flow to sponsor']}:{G}{LL['Cash flow to sponsor']})", "0.00%"),
        ("gap", "IRR minus target", f"=B{e0 + 6}-{LB['target']}", "0.00%"),
    ]
    E_ = {}
    for i, (k, label, fx, fmt) in enumerate(erows):
        rr = e0 + i; E_[k] = f"B{rr}"
        if fx is None:
            continue
        put(L, f"A{rr}", label, bold=(k == "irr")); put(L, f"B{rr}", fx, fmt)
    L[E_["irr"]].fill = OUTFILL
    L.column_dimensions["A"].width = 60
    for col in Q:
        L.column_dimensions[col].width = 16

    # ---------------- Sources ----------------
    s = wb.create_sheet("Sources")
    for col, hd in zip("ABCDEFGH", ["Cell", "What", "XBRL concept / source", "Accession number", "Form", "Filed", "Period start", "Period end"]):
        put(s, f"{col}1", hd, bold=True)
    for i, (cell, what, so) in enumerate(srcs):
        rr = 2 + i
        s[f"A{rr}"] = cell; s[f"B{rr}"] = what
        if isinstance(so, dict):
            s[f"C{rr}"] = so.get("concept"); s[f"D{rr}"] = so.get("accn"); s[f"E{rr}"] = so.get("form")
            s[f"F{rr}"] = so.get("filed"); s[f"G{rr}"] = so.get("start"); s[f"H{rr}"] = so.get("end")
        else:
            s[f"C{rr}"] = so
    s.column_dimensions["A"].width = 16; s.column_dimensions["B"].width = 44; s.column_dimensions["C"].width = 60

    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(path)
    return {
        "file": f"workbooks/sec/{c['ticker']}.xlsx", "ticker": c["ticker"], "name": c["name"],
        "valuePerShare": f"DCF!{V['vps']}", "growth": f"DCF!{A['g']}", "price": f"WACC!{W['price']}", "priceValue": c["price"]["close"],
        "priceDate": c["price"]["date"], "wacc": f"WACC!{W['wacc']}", "waccAdjustment": f"WACC!{W['adj']}", "riskFree": f"WACC!{W['rf']}", "erp": f"WACC!{W['erp']}",
        "beta": f"WACC!{W['beta']}", "costOfEquity": f"WACC!{W['ke']}", "costOfDebt": f"WACC!{W['kd']}", "taxRate": f"WACC!{W['tax']}",
        "cagr5": f"Data!{cagr5}", "cagr9": f"Data!{cagr9}", "fiscalYears": [y["end"] for y in years],
        "lbo": {"irr": f"LBO!{E_['irr']}", "marginImprovement": f"LBO!{LB['dm']}", "premium": f"LBO!{LB['prem']}", "target": LBO["target_irr"],
                "entryMultiple": f"LBO!{LB['mult']}", "ebitdaMargin0": f"LBO!{LB['m0']}", "moic": f"LBO!{E_['moic']}", "growth": f"LBO!{LB['g']}",
                "sponsorEquity": f"LBO!{LB['eq0']}"},
        "ebitProxy": ebit_proxy, "missing": missing, "hardcodedInputsWithSources": len(srcs),
    }


def main():
    data = json.load(open(DATA))
    os.makedirs(OUT, exist_ok=True)
    cells = {"generatedBy": "scripts/make_sec_dcf.py", "from": "data/stage4/companies.json", "priceDate": data["priceDate"],
             "assumptions": {"terminalGrowth": TERMINAL_GROWTH, "blume": BLUME, "workingCapitalPctOfRevenueIncrease": NWC_PCT, "lbo": LBO},
             "companies": []}
    tmp = tempfile.mkdtemp()
    pairs = []
    for c in data["companies"]:
        raw = os.path.join(tmp, f"{c['ticker']}.xlsx")
        info = build(c, data, raw)
        cells["companies"].append(info)
        pairs += [raw, os.path.join(OUT, f"{c['ticker']}.xlsx")]
        print(f"{c['ticker']}: built ({info['hardcodedInputsWithSources']} sourced inputs){'; EBIT = pre-tax + interest' if info['ebitProxy'] else ''}{'; missing ' + ', '.join(info['missing']) if info['missing'] else ''}")
    json.dump(cells, open(os.path.join(OUT, "cells.json"), "w"), indent=1)
    if "--no-recalc" in sys.argv:
        for a, b in zip(pairs[::2], pairs[1::2]):
            os.replace(a, b)
        print("saved WITHOUT cached values (--no-recalc)")
        return
    # Fill in cached values with LibreOffice so the files open with numbers anywhere. Excel's own
    # values replace these when scripts/excel_recalc_sec.sh runs on a Mac with Excel.
    subprocess.run(["/usr/bin/python3", os.path.join(HERE, "lo_recalc.py"), *pairs], check=True)
    for out in pairs[1::2]:
        set_calc(out)


# Excel's Goal Seek stops when the target is within "Maximum Change" (default 0.001). Set it to
# 1e-9, the same setting as the Stage 3 Goal Seek files, so Excel's answers are precise enough to
# compare with ours at 1e-6. (Iteration is switched on only because Excel stores Maximum Change with
# it; these models have no circular references.)
CALC = '<calcPr calcId="191029" iterate="1" iterateCount="1000" iterateDelta="1.0000000000000001E-9"/>'


def set_calc(path):
    import re, zipfile
    with zipfile.ZipFile(path) as z:
        names = z.namelist(); files = {n: z.read(n) for n in names}
    wbx = files["xl/workbook.xml"].decode()
    wbx = re.sub(r"<calcPr[^>]*/>", CALC, wbx) if "<calcPr" in wbx else wbx.replace("</workbook>", CALC + "</workbook>")
    assert CALC in wbx
    files["xl/workbook.xml"] = wbx.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, files[n])


if __name__ == "__main__":
    main()
