#!/bin/bash
# Stage 4, on a Mac with Microsoft Excel: make Excel the ground truth for the ten SEC workbooks.
#
# For each company in workbooks/sec/cells.json:
#   1. Open workbooks/sec/<T>.xlsx in Excel, recalculate everything, save. The file's cached values
#      are then Excel's own (until now they were LibreOffice's).
#   2. Copy it to workbooks/sec/excel/<T>_growth_excel.xlsx and run Excel's Goal Seek: set the value
#      per share to the market price by changing the revenue growth input. Save.
#   3. Copy it to workbooks/sec/excel/<T>_lbo_excel.xlsx and run Goal Seek: set the LBO IRR to 20%
#      by changing the margin improvement input. Save.
# If Excel reports that Goal Seek found no solution, that copy is deleted and only the record says so.
# Writes results/stage4_excel_run.json. Then run `npm run stage4` to compare everything.
#
# Usage: bash scripts/excel_sec.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SEC="$ROOT/workbooks/sec"
OUTDIR="$SEC/excel"
mkdir -p "$OUTDIR" "$ROOT/results"
TMP="$(mktemp -d)"
rm -f "${OUTDIR:?}"/*_excel.xlsx

# ticker|setSheet|setCell|goal|chSheet|chCell for each Goal Seek (from cells.json)
node -e '
const c = require(process.argv[1]);
for (const x of c.companies) {
  const [vs, vc] = x.valuePerShare.split("!"), [gs, gc] = x.growth.split("!");
  const [is, ic] = x.lbo.irr.split("!"), [ms, mc] = x.lbo.marginImprovement.split("!");
  console.log([x.ticker, vs, vc, x.priceValue, gs, gc, is, ic, x.lbo.target, ms, mc].join("|"));
}' "$SEC/cells.json" > "$TMP/cases.txt"

recalc() { # file
  osascript <<EOF
tell application "Microsoft Excel"
  set wb to open workbook workbook file name (POSIX file "$1" as text)
  calculate full
  save wb
  close wb saving no
end tell
EOF
}

seek() { # tag file setSheet setCell goal chSheet chCell
  local tag=$1 work=$2 ss=$3 sc=$4 goal=$5 cs=$6 cc=$7
  osascript <<EOF > "$TMP/$tag.txt"
tell application "Microsoft Excel"
  set wb to open workbook workbook file name (POSIX file "$work" as text)
  set ws to worksheet "$cs" of wb
  activate object ws
  calculate full
  set startVal to string value of range "$sc" of worksheet "$ss" of wb
  set ok to goal seek (range "$sc" of worksheet "$ss" of wb) goal $goal changing cell (range "$cc" of ws)
  set a to value of range "$cc" of ws
  set o to string value of range "$sc" of worksheet "$ss" of wb
  if ok then save wb
  close wb saving no
  return (ok as text) & "|" & (a as text) & "|" & o & "|" & startVal
end tell
EOF
  echo "$tag: $(cat "$TMP/$tag.txt")"
  case "$(cat "$TMP/$tag.txt")" in true*) ;; *) rm -f "${work:?}";; esac
}

while IFS='|' read -r T vs vc price gs gc is ic target ms mc; do
  recalc "$SEC/$T.xlsx"
  echo "$T: recalculated and saved by Excel"
  cp "$SEC/$T.xlsx" "$OUTDIR/${T}_growth_excel.xlsx"
  seek "${T}_growth" "$OUTDIR/${T}_growth_excel.xlsx" "$vs" "$vc" "$price" "$gs" "$gc"
  cp "$SEC/$T.xlsx" "$OUTDIR/${T}_lbo_excel.xlsx"
  seek "${T}_lbo" "$OUTDIR/${T}_lbo_excel.xlsx" "$is" "$ic" "$target" "$ms" "$mc"
done < "$TMP/cases.txt"

EXCEL_VER=$(defaults read "/Applications/Microsoft Excel.app/Contents/Info.plist" CFBundleShortVersionString)
EXCEL_BUILD=$(defaults read "/Applications/Microsoft Excel.app/Contents/Info.plist" CFBundleVersion)
python3 - "$TMP" "$ROOT/results/stage4_excel_run.json" "$EXCEL_VER (build $EXCEL_BUILD)" <<'PY'
import json, sys, platform, subprocess, datetime
tmp, out, ver = sys.argv[1:4]
cases = []
for line in open(f"{tmp}/cases.txt"):
    t = line.split("|")[0]
    for kind in ("growth", "lbo"):
        ok, a, o, start = open(f"{tmp}/{t}_{kind}.txt").read().rstrip("\n").split("|")
        cases.append({"ticker": t, "kind": kind, "excelReportedSuccess": ok == "true", "changingCellAfter": a,
                      "setCellShownBefore": start, "setCellShownAfter": o, "savedFile": ok == "true"})
json.dump({
    "generatedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    "machine": {"what": "Alek's Mac", "cpu": subprocess.run(["sysctl", "-n", "machdep.cpu.brand_string"], capture_output=True, text=True).stdout.strip(),
                "os": "macOS " + platform.mac_ver()[0], "excel": "Microsoft Excel for Mac " + ver},
    "method": "AppleScript: open, `calculate full`, save (workbooks/sec/<T>.xlsx); then `goal seek` (Range.GoalSeek) on copies in workbooks/sec/excel/",
    "note": "Numbers here are AppleScript's display. Full-precision answers are read from the saved files by `npm run stage4` (results/stage4_sec.json companies[].excelGoalSeek).",
    "cases": cases,
}, open(out, "w"), indent=1)
print("wrote", out)
PY
rm -rf "${TMP:?}"
