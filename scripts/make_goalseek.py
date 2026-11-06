"""Build the Stage 3 Goal Seek case files (workbooks/goalseek/gs*.xlsx) for Excel.

Each case is one of the Excel-saved test workbooks plus, at most, two extra cells on the sheet that
holds the changing cell: a text line with the Goal Seek instructions and a helper cell that repeats
the target cell (e.g. =Valuation!B13). That way the Set cell, the To value and the By changing cell
are all on the sheet the file opens on, which is the simplest thing to type into Excel's dialog.

The files also turn on iterative calculation with Maximum Change 0.000000001: Excel's Goal Seek
stops when the target cell is within Maximum Change of the goal, and the default (0.001) is far too
loose for a 1e-6 comparison. (The LBO needs iterative calculation for its loops anyway.)

openpyxl keeps formulas and names but drops cached values, so Excel recalculates on open.
The cases themselves are listed in workbooks/goalseek/cases.json.

Usage: python3 scripts/make_goalseek.py
"""
import json, os, re
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "..", "workbooks")
GS = os.path.join(WB, "goalseek")


def split(ref):
    sheet, addr = ref.rsplit("!", 1)
    return sheet.strip("'"), addr


def main():
    cases = json.load(open(os.path.join(GS, "cases.json")))["cases"]
    for c in cases:
        wb = load_workbook(os.path.join(WB, c["source"]))
        set_sheet, set_addr = split(c["setCell"])
        ch_sheet, ch_addr = split(c["changingCell"])
        ws = wb[ch_sheet]
        if c["helper"]:
            h_sheet, h_addr = split(c["helper"])
            assert h_sheet == ch_sheet, "helper must be on the changing cell's sheet"
            row = int(re.sub(r"[A-Z]", "", h_addr))
            assert ws[h_addr].value is None and ws[f"A{row}"].value is None, f"{c['helper']} is not empty"
            ws[f"A{row}"] = f"Goal Seek target cell (repeats {c['setCell']})"
            ws[h_addr] = f"={c['setCell']}"
            dialog_set = h_addr
            note_row = row + 1
        else:
            assert set_sheet == ch_sheet
            dialog_set = set_addr
            note_row = ws.max_row + 2
        assert ws[f"A{note_row}"].value is None
        ws[f"A{note_row}"] = (f"Stage 3 check {c['id']}: Data > What-If Analysis > Goal Seek. Set cell {dialog_set}, "
                              f"To value {c['toValue']}, By changing cell {ch_addr}. Then OK, OK, and File > Save As.")
        # Open on the changing cell's sheet, with only that sheet selected (several selected sheets
        # would put Excel in "group" mode, where Goal Seek is not available).
        idx = wb.sheetnames.index(ch_sheet)
        for i, s in enumerate(wb.worksheets):
            s.sheet_view.tabSelected = i == idx
        wb.active = idx
        wb.calculation = CalcProperties(fullCalcOnLoad=True, iterate=True, iterateCount=1000, iterateDelta=1e-9)
        out = os.path.join(GS, c["file"])
        wb.save(out)
        c["dialog"] = {"setCell": dialog_set, "toValue": c["toValue"], "byChangingCell": ch_addr, "sheet": ch_sheet}
        print(f"{c['id']}: {out}\n    open on sheet '{ch_sheet}'; Goal Seek: Set cell {dialog_set}, To value {c['toValue']}, By changing cell {ch_addr}")


if __name__ == "__main__":
    main()
