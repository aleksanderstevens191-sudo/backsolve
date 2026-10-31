"""Recalculate .xlsx files in LibreOffice (headless, via its UNO API) and save the computed
values into the files as their cached values.

Why not a plain `soffice --convert-to`? For workbooks with circular references, LibreOffice's
single load-time calculation stops before the loop has converged (we measured fixed-point
residuals of about 1.0 in lbo_circular.xlsx). Calling calculateAll() repeatedly lets its
iterative solver finish, so we recalc until no cell changes, then save.

These cached values come from LibreOffice, not Microsoft Excel.

Usage: /usr/bin/python3 scripts/lo_recalc.py in.xlsx out.xlsx [in2.xlsx out2.xlsx ...]
"""
import os, subprocess, sys, tempfile, time
import uno
from com.sun.star.beans import PropertyValue


def pv(n, v):
    x = PropertyValue(); x.Name = n; x.Value = v; return x


def snapshot(doc):
    vals = []
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        cur = sh.createCursor(); cur.gotoEndOfUsedArea(False)
        rng = sh.getCellRangeByPosition(0, 0, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
        for row in rng.getDataArray():
            vals.extend(x for x in row if isinstance(x, float))
    return vals


def main(pairs):
    profile = tempfile.mkdtemp()
    port = 2002 + os.getpid() % 1000
    proc = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--invisible",
                             "--norestore", f"--accept=socket,host=localhost,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"); break
            except Exception:
                time.sleep(0.5)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        for src, dst in pairs:
            doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)), "_blank", 0, (pv("Hidden", True),))
            prev, passes = None, 0
            for passes in range(1, 501):
                doc.calculateAll()
                cur = snapshot(doc)
                if prev is not None and len(cur) == len(prev) and max((abs(a - b) for a, b in zip(cur, prev)), default=0) == 0:
                    break
                prev = cur
            doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(dst)), (pv("FilterName", "Calc MS Excel 2007 XML"),))
            doc.close(True)
            print(f"{dst}: {passes} LibreOffice recalculation passes until no value changed")
    finally:
        proc.terminate()


if __name__ == "__main__":
    a = sys.argv[1:]
    main(list(zip(a[::2], a[1::2])))
