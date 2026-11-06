"""Run LibreOffice's own Goal Seek on the Stage 3 case files (an independent engine, like the
`formulas` cross-check in Stage 1). This is NOT Excel: Excel's Goal Seek is the ground truth and
runs on Alek's Mac; this is a second opinion that can run in the cloud.

For each case in workbooks/goalseek/cases.json: open the file, recalculate until no value changes
(LibreOffice doesn't finish solving loops on load), call LibreOffice's XGoalSeek.seekGoal, put the
answer into the changing cell, recalculate again until stable, and record what LibreOffice got.

Usage: /usr/bin/python3 scripts/lo_goalseek.py   -> results/stage3_goalseek_libreoffice.json
"""
import json, os, platform, subprocess, sys, tempfile, time
from datetime import datetime, timezone
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellAddress

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
GS = os.path.join(ROOT, "workbooks", "goalseek")


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


def settle(doc):
    prev = None
    for passes in range(1, 501):
        doc.calculateAll()
        cur = snapshot(doc)
        if prev is not None and len(cur) == len(prev) and max((abs(a - b) for a, b in zip(cur, prev)), default=0) == 0:
            return passes
        prev = cur
    return passes


def cell_of(doc, ref):
    sheet, addr = ref.rsplit("!", 1)
    sh = doc.Sheets.getByName(sheet.strip("'"))
    c = sh.getCellRangeByName(addr)
    a = c.getCellAddress()
    return c, CellAddress(a.Sheet, a.Column, a.Row)


def main():
    cases = json.load(open(os.path.join(GS, "cases.json")))["cases"]
    profile = tempfile.mkdtemp()
    port = 2002 + os.getpid() % 1000
    proc = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--invisible",
                             "--norestore", f"--accept=socket,host=localhost,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out = []
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"); break
            except Exception:
                time.sleep(0.5)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        prov = smgr.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
        about = prov.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", (pv("nodepath", "/org.openoffice.Setup/Product"),))
        version = f"{about.getByName('ooName')} {about.getByName('ooSetupVersionAboutBox')}"
        for c in cases:
            doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(os.path.join(GS, c["file"]))), "_blank", 0, (pv("Hidden", True),))
            settle(doc)
            target, taddr = cell_of(doc, c["setCell"])
            var, vaddr = cell_of(doc, c["changingCell"])
            start = var.getValue()
            t0 = time.perf_counter()
            res = doc.seekGoal(taddr, vaddr, repr(float(c["toValue"])))
            ms = (time.perf_counter() - t0) * 1000
            var.setValue(res.Result)
            passes = settle(doc)
            got = target.getValue()
            row = {"id": c["id"], "file": c["file"], "setCell": c["setCell"], "toValue": c["toValue"], "changingCell": c["changingCell"],
                   "start": start, "libreofficeAnswer": res.Result, "libreofficeDivergence": res.Divergence,
                   "setCellAfterwards": got, "missBy": got - c["toValue"], "recalcPassesAfterwards": passes, "ms": ms}
            out.append(row)
            print(f"{c['id']}: LibreOffice Goal Seek -> {c['changingCell']} = {res.Result!r} (divergence {res.Divergence:.3g}); {c['setCell']} then = {got!r}")
            doc.close(True)
    finally:
        proc.terminate()
    os.makedirs(os.path.join(ROOT, "results"), exist_ok=True)
    with open(os.path.join(ROOT, "results", "stage3_goalseek_libreoffice.json"), "w") as f:
        json.dump({"generatedAt": datetime.now(timezone.utc).isoformat(), "engine": version,
                   "machine": f"{platform.system()} {platform.release()}, Python {sys.version.split()[0]}",
                   "note": "LibreOffice Goal Seek (XGoalSeek.seekGoal), not Excel. Second opinion only.", "cases": out}, f, indent=2)
        f.write("\n")


if __name__ == "__main__":
    main()
