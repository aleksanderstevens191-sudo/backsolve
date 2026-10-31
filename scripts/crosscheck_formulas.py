"""Cross-check: recalculate each test workbook with the Python `formulas` library (an independent
Excel engine) and compare it cell by cell with (a) the file's cached values and (b) our engine.

Run `npx tsx scripts/check.ts` first: it writes results/engine_values/*.json.
Usage: python3 scripts/crosscheck_formulas.py [extra.xlsx ...]

Note: `formulas` handles a circular reference by cutting the loop (the cut input is treated
as an error/zero), not by iterating like Excel, so cells in or downstream of a loop are
reported separately rather than counted as passes or failures.
"""
import glob, json, os, platform, re, sys, time
import numpy as np
import formulas
import datetime, zipfile
import openpyxl
from openpyxl.utils.datetime import to_excel
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
TOL = 1e-9


def norm_key(k):
    # "'[file.xlsx]SHEET NAME'!B5" -> ("SHEET NAME", "B5")
    m = re.match(r"^'?\[[^\]]+\](.*?)'?!([A-Z]+[0-9]+)$", k)
    return (m.group(1).replace("''", "'").upper(), m.group(2)) if m else None


def our_key(k):
    sheet, addr = k.rsplit("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet.upper(), addr


def scalar(v):
    if hasattr(v, "value"):
        v = v.value
    if isinstance(v, np.ndarray):
        v = v.ravel()[0] if v.size else None
    if isinstance(v, np.generic):
        v = v.item()
    if type(v).__name__ == "XlError":
        return {"error": str(v)}
    if type(v).__name__ in ("_Empty", "Empty") or str(v) == "empty":
        return None
    return v


def same(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= TOL * max(1.0, abs(b))
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, dict) and isinstance(b, dict):
        return a.get("error") == b.get("error")
    return a == b


def main(paths):
    report = {"generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
              "machine": {"os": platform.platform(), "python": platform.python_version(), "formulas": formulas.__version__},
              "tolerance": f"|a - b| <= {TOL} * max(1, |b|)", "workbooks": []}
    ok_all = True
    for path in paths:
        name = os.path.basename(path)
        ours_path = os.path.join(ROOT, "results", "engine_values", name.replace(".xlsx", ".json"))
        dump = json.load(open(ours_path))
        ours = {our_key(k): v for k, v in dump["values"].items()}
        tainted = {our_key(k) for k in dump["inOrDownstreamOfLoops"]}

        cached_wb = openpyxl.load_workbook(path, data_only=True)
        with zipfile.ZipFile(path) as z:
            libre = b"LibreOffice" in z.read("docProps/app.xml")
        known = json.load(open(os.path.join(ROOT, "workbooks", "known_differences.json"))).get(name, {})
        cached = {}
        for ws in cached_wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, (datetime.datetime, datetime.date)):
                        v = to_excel(v)  # openpyxl turns date-formatted numbers into dates; compare the serial number
                    if cell.data_type == "e":
                        v = {"error": v}
                    cached[(ws.title.upper(), cell.coordinate)] = v

        t = time.time()
        model = formulas.ExcelModel().loads(path).finish(circular=True)
        sol = model.calculate()
        secs = time.time() - t
        fv = {}
        for k, v in sol.items():
            nk = norm_key(k)
            if nk:
                fv[nk] = scalar(v)
            else:  # a range result such as an array formula's spill block 'G2:G3'
                m = re.match(r"^'?\[[^\]]+\](.*?)'?!([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$", k)
                arr = getattr(v, "value", None)
                if m and isinstance(arr, np.ndarray):
                    sheet = m.group(1).replace("''", "'").upper()
                    c1, r1 = column_index_from_string(m.group(2)), int(m.group(3))
                    for i in range(arr.shape[0]):
                        for j in range(arr.shape[1]):
                            fv.setdefault((sheet, f"{get_column_letter(c1 + j)}{r1 + i}"), scalar(arr[i, j]))

        rows = {"vs_cached": [0, 0], "vs_ours": [0, 0]}
        diffs, loop_diffs, missing = [], [], []
        for key in ours:
            if key not in fv:
                missing.append("!".join(key)); continue
            f = fv[key]
            ck = "!".join(key)
            exp = next((e["excel"] for c, e in known.items() if our_key(c) == key), None) if libre else None
            for label, other in (("vs_cached", exp if exp is not None else cached.get(key)), ("vs_ours", ours[key])):
                rows[label][0] += 1
                if same(f, other):
                    rows[label][1] += 1
                else:
                    entry = {"cell": "!".join(key), "formulas": f, "other": other, "against": label}
                    (loop_diffs if key in tainted else diffs).append(entry)
        failed = len(diffs) > 0 or len(missing) > 0
        ok_all &= not failed
        w = {"file": name, "formulasSeconds": round(secs, 2),
             "compared_vs_cached": rows["vs_cached"][0], "agree_vs_cached": rows["vs_cached"][1],
             "compared_vs_ours": rows["vs_ours"][0], "agree_vs_ours": rows["vs_ours"][1],
             "cellsInOrDownstreamOfLoops": len(tainted),
             "loopCellsWhereFormulasDisagrees": len({d["cell"] for d in loop_diffs}),
             "disagreementsOutsideLoops": diffs[:50], "missingFromFormulas": missing[:50], "ok": not failed}
        report["workbooks"].append(w)
        print(f"\n{name}: formulas library took {secs:.2f}s")
        print(f"  formulas vs cached: {w['agree_vs_cached']}/{w['compared_vs_cached']} agree")
        print(f"  formulas vs our engine: {w['agree_vs_ours']}/{w['compared_vs_ours']} agree")
        if tainted:
            print(f"  {len(tainted)} cells are in or downstream of a circular loop; formulas disagrees on {w['loopCellsWhereFormulasDisagrees']} of them"
                  " (it cuts loops instead of iterating, so these are not counted as failures)")
        for d in diffs[:10]:
            print("  DISAGREES OUTSIDE LOOPS", d)
        if missing:
            print(f"  {len(missing)} cells missing from formulas output, e.g. {missing[:5]}")
        print("  " + ("PASS" if not failed else "FAIL"))
    out = os.path.join(ROOT, "results", "stage1_crosscheck_formulas.json")
    json.dump(report, open(out, "w"), indent=2, default=str)
    print(f"\nwrote {os.path.relpath(out, ROOT)} — overall {'PASS' if ok_all else 'FAIL'}")
    return 0 if ok_all else 1


if __name__ == "__main__":
    files = sorted(glob.glob(os.path.join(ROOT, "workbooks", "*.xlsx"))) + sys.argv[1:]
    sys.exit(main(files))
