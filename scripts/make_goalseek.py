"""Build the Stage 3 Goal Seek case files (workbooks/goalseek/gs*.xlsx) for Excel.

Each case is one of the Excel-saved test workbooks plus, at most, two extra cells on the sheet that
holds the changing cell: a text line with the Goal Seek instructions and a helper cell that repeats
the target cell (e.g. =Valuation!B13). That way the Set cell, the To value and the By changing cell
are all on the sheet the file opens on, which is the simplest thing to type into Excel's dialog.

Why this edits the .xlsx XML directly instead of using openpyxl: openpyxl throws away every cached
value (the numbers Excel saved). For the circular LBO that is fatal: Excel has to solve the loops
from nothing when the file opens, an error gets into the interest loop and feeds itself, and the IRR
shows #VALUE!. Patching the XML keeps every one of Excel's saved numbers, so the file opens exactly as
Excel left it. The helper cell gets the target cell's own cached value, so it is consistent too.

The files keep iterative calculation on with Maximum Change 0.000000001: Excel's Goal Seek stops when
the target cell is within Maximum Change of the goal, and the default (0.001) is far too loose for a
1e-6 comparison. (The LBO needs iterative calculation for its loops anyway.)

The cases themselves are listed in workbooks/goalseek/cases.json.

Usage: python3 scripts/make_goalseek.py
"""
import json, os, re, zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, "..", "workbooks")
GS = os.path.join(WB, "goalseek")
CALC = '<calcPr calcId="191029" iterate="1" iterateCount="1000" iterateDelta="1.0000000000000001E-9"/>'


def split(ref):
    sheet, addr = ref.rsplit("!", 1)
    return sheet.strip("'"), addr


def rc(addr):
    m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


def col_name(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def sheet_paths(files):
    """sheet name -> path of its XML part, in workbook order."""
    wbx = files["xl/workbook.xml"].decode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    targets = {}
    for rel in re.findall(r"<Relationship [^>]*>", rels):
        rid = re.search(r'Id="([^"]+)"', rel).group(1)
        targets[rid] = re.search(r'Target="([^"]+)"', rel).group(1)
    out = {}
    for tag in re.findall(r"<sheet [^>]*>", wbx):
        name = re.search(r'name="([^"]+)"', tag).group(1).replace("&amp;", "&")
        t = targets[re.search(r'r:id="([^"]+)"', tag).group(1)]
        out[name] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    return out


def cached(xml, addr):
    m = re.search(rf'<c r="{addr}"[^>]*?(?:/>|>(.*?)</c>)', xml, re.S)
    v = m and m.group(1) and re.search(r"<v>(.*?)</v>", m.group(1))
    assert v, f"{addr} has no cached value"
    return v.group(1)


def put_cell(xml, addr, cell_xml):
    """Insert a cell into <sheetData>, keeping rows and cells in order. The cell must be empty."""
    row, col = rc(addr)
    assert not re.search(rf'<c r="{addr}"[ >/]', xml), f"{addr} is not empty"
    for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S):
        r = int(m.group(1))
        if r == row:
            body = m.group(2) or ""
            cells = list(re.finditer(r'<c r="([A-Z]+\d+)"', body))
            pos = next((x.start() for x in cells if rc(x.group(1))[1] > col), len(body))
            open_tag = re.sub(r' spans="[^"]*"', "", re.match(r"<row [^>]*?(?=/?>)", m.group(0)).group(0))
            return xml[:m.start()] + f"{open_tag}>{body[:pos]}{cell_xml}{body[pos:]}</row>" + xml[m.end():]
        if r > row:
            return xml[:m.start()] + f'<row r="{row}">{cell_xml}</row>' + xml[m.start():]
    if "</sheetData>" in xml:
        i = xml.index("</sheetData>")
        return xml[:i] + f'<row r="{row}">{cell_xml}</row>' + xml[i:]
    return xml.replace("<sheetData/>", f'<sheetData><row r="{row}">{cell_xml}</row></sheetData>')


def text_cell(addr, s):
    return f'<c r="{addr}" t="inlineStr"><is><t>{escape(s)}</t></is></c>'


def grow_dimension(xml, addr):
    row, col = rc(addr)
    m = re.search(r'<dimension ref="([A-Z]+\d+)(?::([A-Z]+\d+))?"/>', xml)
    if not m:
        return xml
    r2, c2 = rc(m.group(2) or m.group(1))
    return xml[:m.start()] + f'<dimension ref="{m.group(1)}:{col_name(max(c2, col))}{max(r2, row)}"/>' + xml[m.end():]


def build(c):
    with zipfile.ZipFile(os.path.join(WB, c["source"])) as z:
        order = [i.filename for i in z.infolist()]
        files = {n: z.read(n) for n in order}
    paths = sheet_paths(files)
    set_sheet, set_addr = split(c["setCell"])
    ch_sheet, ch_addr = split(c["changingCell"])
    ch_path = paths[ch_sheet]
    xml = files[ch_path].decode()
    if c["helper"]:
        h_sheet, h_addr = split(c["helper"])
        assert h_sheet == ch_sheet, "helper must be on the changing cell's sheet"
        row = rc(h_addr)[0]
        v = cached(files[paths[set_sheet]].decode(), set_addr)  # Excel's own number for the target cell
        target = c["setCell"] if " " not in set_sheet else f"'{set_sheet}'!{set_addr}"
        xml = put_cell(xml, f"A{row}", text_cell(f"A{row}", f"Goal Seek target cell (repeats {c['setCell']})"))
        xml = put_cell(xml, h_addr, f'<c r="{h_addr}"><f>{escape(target)}</f><v>{v}</v></c>')
        xml = grow_dimension(xml, h_addr)
        dialog_set, note_row = h_addr, row + 1
    else:
        assert set_sheet == ch_sheet
        dialog_set = set_addr
        note_row = max(int(r) for r in re.findall(r'<row r="(\d+)"', xml)) + 2
    note = (f"Stage 3 check {c['id']}: Data > What-If Analysis > Goal Seek. Set cell {dialog_set}, "
            f"To value {c['toValue']}, By changing cell {ch_addr}. Then OK, OK, and File > Save As.")
    xml = put_cell(xml, f"A{note_row}", text_cell(f"A{note_row}", note))
    files[ch_path] = grow_dimension(xml, f"A{note_row}").encode()

    # Open on the changing cell's sheet, with only that sheet selected (several selected sheets would
    # put Excel in "group" mode, where Goal Seek is not available).
    idx = list(paths).index(ch_sheet)
    for name, p in paths.items():
        s = re.sub(r' tabSelected="[01]"', "", files[p].decode())
        if name == ch_sheet:
            s = s.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
        files[p] = s.encode()
    wbx = re.sub(r' activeTab="\d+"', "", files["xl/workbook.xml"].decode())
    if idx:
        wbx = wbx.replace("<workbookView ", f'<workbookView activeTab="{idx}" ', 1)
    wbx = re.sub(r"<calcPr[^>]*/>", CALC, wbx)
    assert CALC in wbx
    files["xl/workbook.xml"] = wbx.encode()

    out = os.path.join(GS, c["file"])
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(n, files[n])
    print(f"{c['id']}: {out}\n    open on sheet '{ch_sheet}'; Goal Seek: Set cell {dialog_set}, To value {c['toValue']}, By changing cell {ch_addr}")


def main():
    for c in json.load(open(os.path.join(GS, "cases.json")))["cases"]:
        build(c)


if __name__ == "__main__":
    main()
