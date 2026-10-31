"""Build the Stage 1 test workbooks with live formulas.

openpyxl writes formulas but no computed values, so afterwards we open each file in
LibreOffice (headless) and save it again: that fills in the cached values that our
engine is checked against. These cached values come from LibreOffice, NOT Microsoft
Excel. Opening the files in Excel and saving them would replace them with Excel's own.

Usage: python3 scripts/make_workbooks.py
"""
import os, subprocess, tempfile
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "workbooks")
COLS = "BCDEFG"  # B = base/actual year, C..G = five projection years


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def three_statement(path):
    wb = Workbook()
    a = wb.active; a.title = "Assumptions"
    a["A1"] = "Three-statement model: assumptions"
    a["A4"] = "Year"
    for i, c in enumerate(COLS):
        a[f"{c}4"] = 2025 + i
    rows = {
        5: ("Revenue growth", [None, .08, .07, .06, .05, .04]),
        6: ("Gross margin", [None, .40, .41, .42, .42, .43]),
        7: ("SG&A % revenue", [None, .18, .175, .17, .17, .165]),
        8: ("D&A % revenue", [None, .04, .04, .04, .04, .04]),
        9: ("Capex % revenue", [None, .05, .05, .045, .045, .04]),
        10: ("DSO (days)", [None, 45, 44, 43, 42, 42]),
        11: ("DIO (days)", [None, 50, 50, 49, 48, 48]),
        12: ("DPO (days)", [None, 38, 38, 39, 40, 40]),
    }
    for r, (label, vals) in rows.items():
        a[f"A{r}"] = label
        for c, v in zip(COLS, vals):
            if v is not None:
                a[f"{c}{r}"] = v
    singles = {3: ("Base-year revenue", 1000, "BaseRevenue"), 13: ("Tax rate", .25, "TaxRate"),
               14: ("Interest rate on debt", .07, "InterestRate"), 15: ("Mandatory debt repayment", 40, "DebtRepay"),
               16: ("Minimum cash", 50, "MinCash"), 17: ("Dividend payout", .30, "Payout")}
    for r, (label, v, nm) in singles.items():
        a[f"A{r}"] = label; a[f"B{r}"] = v
        name(wb, nm, f"Assumptions!$B${r}")

    IS = wb.create_sheet("IS"); BS = wb.create_sheet("Balance Sheet"); CF = wb.create_sheet("Cash Flow"); V = wb.create_sheet("Valuation")
    for ws in (IS, BS, CF):
        ws["A4"] = "Year"
        for c in COLS:
            ws[f"{c}4"] = f"=Assumptions!{c}4"

    labels = {5: "Revenue", 6: "COGS", 7: "Gross profit", 8: "SG&A", 9: "EBITDA", 10: "D&A", 11: "EBIT",
              12: "Interest expense", 13: "Pre-tax income", 14: "Taxes", 15: "Net income", 16: "EBITDA margin", 17: "Status"}
    for r, l in labels.items():
        IS[f"A{r}"] = l
    IS["B5"] = "=BaseRevenue"; IS["B6"] = -600; IS["B8"] = -170; IS["B10"] = -40; IS["B12"] = -26
    for r, f in {7: "=SUM(B5:B6)", 9: "=B7+B8", 11: "=B9+B10", 13: "=B11+B12", 14: "=-MAX(0,B13)*TaxRate", 15: "=B13+B14", 16: "=B9/B5"}.items():
        IS[f"B{r}"] = f
    for i, c in enumerate(COLS[1:]):
        p = COLS[i]  # previous column
        IS[f"{c}5"] = f"={p}5*(1+Assumptions!{c}5)"
        IS[f"{c}6"] = f"=-{c}5*(1-Assumptions!{c}6)"
        IS[f"{c}7"] = f"=SUM({c}5:{c}6)"
        IS[f"{c}8"] = f"=-{c}5*Assumptions!{c}7"
        IS[f"{c}9"] = f"={c}7+{c}8"
        IS[f"{c}10"] = f"=-{c}5*Assumptions!{c}8"
        IS[f"{c}11"] = f"={c}9+{c}10"
        IS[f"{c}12"] = f"=-InterestRate*('Balance Sheet'!{p}12+'Balance Sheet'!{p}13)"  # on opening balances: no circularity
        IS[f"{c}13"] = f"={c}11+{c}12"
        IS[f"{c}14"] = f"=-MAX(0,{c}13)*Assumptions!$B$13"
        IS[f"{c}15"] = f"={c}13+{c}14"
        IS[f"{c}16"] = f"={c}9/{c}5"
    for c in COLS:
        IS[f"{c}17"] = f'=IF(AND({c}9>0,{c}15>0),"Profitable",IF(OR({c}9>0,{c}15>0),"Mixed","Loss"))'

    for r, l in {5: "Cash", 6: "Accounts receivable", 7: "Inventory", 8: "PP&E", 9: "Total assets", 11: "Accounts payable",
                 12: "Term debt", 13: "Revolver", 14: "Total liabilities", 15: "Equity", 16: "Total liabilities + equity",
                 17: "Assets - (L+E)", 18: "Balance check"}.items():
        BS[f"A{r}"] = l
    for r, v in {5: 100, 6: 120, 7: 80, 8: 500, 11: 60, 12: 400, 13: 0}.items():
        BS[f"B{r}"] = v
    BS["B9"] = "=SUM(B5:B8)"; BS["B14"] = "=SUM(B11:B13)"; BS["B15"] = "=B9-B14"
    for i, c in enumerate(COLS[1:]):
        p = COLS[i]
        BS[f"{c}5"] = f"='Cash Flow'!{c}16"
        BS[f"{c}6"] = f"=IS!{c}5*Assumptions!{c}10/365"
        BS[f"{c}7"] = f"=-IS!{c}6*Assumptions!{c}11/365"
        BS[f"{c}8"] = f"={p}8-'Cash Flow'!{c}10+IS!{c}10"
        BS[f"{c}9"] = f"=SUM({c}5:{c}8)"
        BS[f"{c}11"] = f"=-IS!{c}6*Assumptions!{c}12/365"
        BS[f"{c}12"] = f"={p}12+'Cash Flow'!{c}13"
        BS[f"{c}13"] = f"={p}13+'Cash Flow'!{c}14"
        BS[f"{c}14"] = f"=SUM({c}11:{c}13)"
        BS[f"{c}15"] = f"={p}15+IS!{c}15+'Cash Flow'!{c}15"
    for c in COLS:
        BS[f"{c}16"] = f"={c}14+{c}15"
        BS[f"{c}17"] = f"={c}9-{c}16"
        BS[f"{c}18"] = f'=IF(ABS({c}17)<0.001,"OK","CHECK")'

    for r, l in {5: "Net income", 6: "D&A", 7: "(Increase) in AR", 8: "(Increase) in inventory", 9: "Increase in AP",
                 10: "Capex", 11: "Operating + investing cash flow", 12: "Cash before financing", 13: "Term debt repayment",
                 14: "Revolver draw / (repayment)", 15: "Dividends", 16: "Ending cash", 17: "Unlevered free cash flow"}.items():
        CF[f"A{r}"] = l
    for i, c in enumerate(COLS[1:]):
        p = COLS[i]
        CF[f"{c}5"] = f"=IS!{c}15"
        CF[f"{c}6"] = f"=-IS!{c}10"
        CF[f"{c}7"] = f"='Balance Sheet'!{p}6-'Balance Sheet'!{c}6"
        CF[f"{c}8"] = f"='Balance Sheet'!{p}7-'Balance Sheet'!{c}7"
        CF[f"{c}9"] = f"='Balance Sheet'!{c}11-'Balance Sheet'!{p}11"
        CF[f"{c}10"] = f"=-IS!{c}5*Assumptions!{c}9"
        CF[f"{c}11"] = f"=SUM({c}5:{c}10)"
        CF[f"{c}12"] = f"='Balance Sheet'!{p}5+{c}11"
        CF[f"{c}13"] = f"=-MIN(DebtRepay,'Balance Sheet'!{p}12)"
        CF[f"{c}14"] = f"=MAX(MinCash-({c}12+{c}13+{c}15),-'Balance Sheet'!{p}13)"
        CF[f"{c}15"] = f"=-MAX(0,IS!{c}15)*Payout"
        CF[f"{c}16"] = f"={c}12+{c}13+{c}14+{c}15"
        CF[f"{c}17"] = f"=IS!{c}11*(1-TaxRate)+SUM({c}6:{c}10)"

    V["A1"] = "DCF valuation"
    vals = [(3, "WACC", 0.09), (4, "Terminal growth", 0.025), (12, "Shares outstanding", 100)]
    for r, l, v in vals:
        V[f"A{r}"] = l; V[f"B{r}"] = v
    name(wb, "WACC", "Valuation!$B$3"); name(wb, "TGrowth", "Valuation!$B$4")
    name(wb, "UFCF", "'Cash Flow'!$C$17:$G$17")
    for r, l, f in [
        (6, "PV of forecast FCF", "=NPV(WACC,UFCF)"),
        (7, "Terminal value", "='Cash Flow'!G17*(1+TGrowth)/(WACC-TGrowth)"),
        (8, "PV of terminal value", "=B7/(1+WACC)^5"),
        (9, "Enterprise value", "=B6+B8"),
        (10, "Net debt", "='Balance Sheet'!B12+'Balance Sheet'!B13-'Balance Sheet'!B5"),
        (11, "Equity value", "=B9-B10"),
        (13, "Value per share", "=B11/B12"),
        (14, "Average EBITDA margin", "=AVERAGE(IS!C16:G16)"),
        (15, "Revenue-weighted gross margin", "=SUMPRODUCT(IS!C5:G5,Assumptions!C6:G6)/SUM(IS!C5:G5)"),
        (16, "Lowest EBITDA", "=MIN(IS!C9:G9)"),
        (17, "Highest EBITDA", "=MAX(IS!C9:G9)"),
        (18, "Terminal value share of EV", "=B8/B9"),
        (19, "Sanity", '=IF(OR(B18>0.85,WACC<=TGrowth),"Check terminal value","OK")'),
        (20, "Revenue CAGR", "=(IS!G5/IS!B5)^(1/5)-1"),
    ]:
        V[f"A{r}"] = l; V[f"B{r}"] = f
    name(wb, "SharePrice", "Valuation!$B$13")
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(path)


def lbo(path):
    wb = Workbook()
    I = wb.active; I.title = "Inputs"
    I["A1"] = "LBO with circular interest (interest on average balances)"
    inputs = [(3, "Entry multiple (x EBITDA)", 10, "EntryMultiple"), (4, "Debt % of entry EV", 0.6, None),
              (5, "Interest rate on debt", 0.08, "IntRate"), (6, "Interest rate on cash", 0.03, "CashRate"),
              (7, "Base revenue", 500, None), (8, "Revenue growth", 0.05, None), (9, "Entry EBITDA margin", 0.2, None),
              (10, "Margin improvement per year", 0.005, "MarginStep"), (11, "Capex % revenue", 0.03, None),
              (12, "Tax rate", 0.25, "TaxRateL"), (13, "Exit multiple (x EBITDA)", 10, "ExitMultiple"),
              (14, "Minimum cash", 20, None), (15, "Fees % of EV", 0.02, None), (16, "Close date", 45657, None),
              (17, "Target IRR", 0.2, "TargetIRR")]
    for r, l, v, nm in inputs:
        I[f"A{r}"] = l; I[f"B{r}"] = v
        if nm:
            name(wb, nm, f"Inputs!$B${r}")
    I["B16"].number_format = "yyyy-mm-dd"

    M = wb.create_sheet("Model")
    for r, l in {4: "Year", 5: "Date", 6: "Revenue", 7: "EBITDA margin", 8: "EBITDA", 9: "Capex", 10: "Interest expense",
                 11: "Interest income", 12: "Taxes", 13: "Ending debt", 14: "Ending cash", 15: "Debt paydown (cash sweep)",
                 16: "Entry enterprise value", 17: "Free cash flow before sweep", 18: "Sponsor equity"}.items():
        M[f"A{r}"] = l
    M["B4"] = 0; M["B5"] = "=Inputs!B16"; M["B6"] = "=Inputs!$B$7"; M["B7"] = "=Inputs!$B$9"; M["B8"] = "=B6*B7"
    M["B16"] = "=B8*EntryMultiple"; M["B13"] = "=B16*Inputs!$B$4"; M["B14"] = "=Inputs!$B$14"
    M["B18"] = "=B16*(1+Inputs!$B$15)-B13+B14"
    for i, c in enumerate(COLS[1:]):
        p = COLS[i]
        M[f"{c}4"] = f"={p}4+1"
        M[f"{c}5"] = f"={p}5+365"
        M[f"{c}6"] = f"={p}6*(1+Inputs!$B$8)"
        M[f"{c}7"] = f"={p}7+MarginStep"
        M[f"{c}8"] = f"={c}6*{c}7"
        M[f"{c}9"] = f"=-{c}6*Inputs!$B$11"
        M[f"{c}10"] = f"=-IntRate*AVERAGE({p}13,{c}13)"     # circular: depends on this year's ending debt
        M[f"{c}11"] = f"=CashRate*AVERAGE({p}14,{c}14)"     # circular: depends on this year's ending cash
        M[f"{c}12"] = f"=-MAX(0,{c}8+{c}10+{c}11)*TaxRateL"
        M[f"{c}17"] = f"=SUM({c}8:{c}12)"
        M[f"{c}15"] = f"=MAX(0,MIN({p}14+{c}17-Inputs!$B$14,{p}13))"
        M[f"{c}13"] = f"={p}13-{c}15"
        M[f"{c}14"] = f"={p}14+{c}17-{c}15"

    R = wb.create_sheet("Returns")
    for r, l, f in [(3, "Exit enterprise value", "=Model!G8*ExitMultiple"), (4, "Exit equity value", "=B3-Model!G13+Model!G14"),
                    (5, "MOIC", "=B4/Model!B18"), (8, "IRR", "=IRR(B7:G7)"), (9, "XNPV at target IRR", "=XNPV(TargetIRR,B7:G7,Model!B5:G5)"),
                    (10, "NPV at target IRR", "=B7+NPV(TargetIRR,C7:G7)"), (11, "Meets target?", '=IF(B8>=TargetIRR,"Yes","No")'),
                    (12, "Strong deal?", "=OR(B5>2.5,B8>0.25)"), (13, "Fully levered and profitable?", "=AND(Model!G13>0,Model!G8>Model!B8)"),
                    (14, "Total interest paid", "=-SUM(Model!C10:G10)"), (15, "Peak debt / EBITDA", "=MAX(Model!B13/Model!B8,Model!C13/Model!C8)"),
                    (16, "Average sweep", "=AVERAGE(Model!C15:G15)"), (17, "Debt x rate check", "=SUMPRODUCT(Model!C13:G13,Model!C7:G7)")]:
        R[f"A{r}"] = l; R[f"B{r}"] = f
    R["A7"] = "Equity cash flows"
    R["B7"] = "=-Model!B18"
    for c in COLS[1:]:
        R[f"{c}7"] = f"=IF(Model!{c}4=5,$B$4,0)"
    # Turn on iterative calculation with a tight tolerance so every engine lands on the same fixed point.
    wb.calculation = CalcProperties(fullCalcOnLoad=True, iterate=True, iterateCount=100, iterateDelta=1e-12)
    wb.save(path)


def coverage(path):
    """Every supported function plus Excel's awkward rules (text in ranges, TRUE as 1, precedence)."""
    from openpyxl.worksheet.formula import ArrayFormula
    wb = Workbook()
    ws = wb.active; ws.title = "Funcs"
    two = wb.create_sheet("Sheet Two"); other = wb.create_sheet("Other")
    for r, v in enumerate([1, 2, "text", True, None, 5.5], start=1):
        if v is not None:
            ws[f"A{r}"] = v
    for r, (cf, d) in enumerate(zip([-1000, 300, 400, 350, 200], [45292, 45474, 45658, 46023, 46387]), start=1):
        ws[f"C{r}"] = cf; ws[f"D{r}"] = d
    two["B2"] = 7; other["A1"] = 3
    for r, v in enumerate([10, 20, 30], start=1):
        two[f"A{r}"] = v
    name(wb, "Data", "'Sheet Two'!$A$1:$A$3")
    name(wb, "Rate", "Other!$A$1")
    cases = [
        "=SUM(A1:A6)", '=SUM(1,"2",TRUE)', "=AVERAGE(A1:A6)", "=MIN(A1:A6)", "=MAX(A1:A6)", "=MIN(A5)",
        "=-2^2", "=2^3^2", "=10%*50", "=1/0", '=IF(A1>0,"pos")', '=IF(A1<0,"neg")', "=AND(A1:A4)", "=OR(A1>5,A2>5)",
        "=SUMPRODUCT((C1:C5>0)*1,C1:C5)", "=NPV(0.1,A1:A6)", "=IRR({-100,30,40,50})", "=IRR(C1:C5,0.2)", "=XNPV(0.08,C1:C5,D1:D5)",
        '="Value: "&A6', "=A1=A2", '="abc"="ABC"', "=ROUND(2.675,2)", "=ABS(-3)", '=IFERROR(1/0,"x")', "=NOT(A4)", "=COUNT(A1:A6)",
        "=SUM(Data)", "=Other!A1+'Sheet Two'!B2", "=IF(AND(A1>0,OR(A2>5,A6>5)),SUM(A1:A2)*MAX(A6,1),MIN(A1,A2))",
        "=SUMPRODUCT(C1:C5,D1:D5)/SUM(D1:D5)", "=C1*$A$1+C$2*$A2", "=SUM(C:C)", "=MAX(-1,-2,TRUE)", '=5-"2"', '="a"+1',
        "=NPV(Rate/10,C2:C5)+C1", "=AVERAGE(Data,A1,4)", "=IF(A1,A2,A6)", "=OR(FALSE,A5)", "=1+2*3-4/2^2", "=(1+2)*3", "=-A6%",
        '=IF(A3="TEXT",1,0)', "=A2>A1", '="b">"a"', "=AND(TRUE,1,2)", "=SUMPRODUCT(A1:A2,A1:A2)", "=MAX(C1:C5)-MIN(C1:C5)",
        "=IRR(C1:C5)*100", "=XNPV(Rate/10,C1:C5,D1:D5)", "=ROUND(-1.5,0)", "=ROUND(1234.5678,-2)", "=A5+1", '=A5&"x"',
        "=SUM(A1,A2,A6)/COUNT(A1:A6)", "=NPV(0.05,1,2,3)", "=IF(1,,2)", "=AVERAGE(1,2,3,4)", "=MIN(3,1,2)",
    ]
    for i, f in enumerate(cases, start=1):
        ws[f"F{i}"] = f
        ws[f"E{i}"] = f[1:]  # label: the formula text
    ws["G1"] = ArrayFormula("G1", "=SUM(A1:A2*C1:C2)")
    ws["G2"] = ArrayFormula("G2:G3", "=C1:C2*2")
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(path)


def recalc_with_libreoffice(pairs):
    # LibreOffice's Python (with the `uno` module) is the system one.
    args = [x for pair in pairs for x in pair]
    subprocess.run(["/usr/bin/python3", os.path.join(HERE, "lo_recalc.py"), *args], check=True)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    raw = tempfile.mkdtemp()
    pairs = []
    for fn, builder in [("three_statement.xlsx", three_statement), ("lbo_circular.xlsx", lbo), ("functions_coverage.xlsx", coverage)]:
        src = os.path.join(raw, fn)
        builder(src)
        pairs.append((src, os.path.join(OUT, fn)))
    recalc_with_libreoffice(pairs)
