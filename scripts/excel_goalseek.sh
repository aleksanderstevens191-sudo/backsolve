#!/bin/bash
# Run Microsoft Excel's own Goal Seek on the four Stage 3 cases, on a Mac with Excel installed.
#
# What it does, case by case (cases listed in workbooks/goalseek/cases.json): copy
# workbooks/goalseek/gsN_*.xlsx to workbooks/goalseek/excel/gsN_*_excel.xlsx, open it in Excel, run
# Goal Seek (AppleScript `goal seek`, the same engine as Data > What-If Analysis > Goal Seek) and save.
# `npm run stage3` then reads the saved files. If Excel reports that it found no solution, the copy is
# deleted and only this script's record says what happened.
#
# The case files come from scripts/make_goalseek.py, which keeps Excel's cached values. (An earlier
# openpyxl version dropped them, and the circular LBO then opened with #VALUE! in Excel.)
#
# Writes results/stage3_goalseek_excel_run.json with what Excel returned for each case.
# Excel's displayed numbers here are AppleScript's ~12-digit rendering; the full-precision answers
# are read from the saved files by `npm run stage3`.
#
# Usage: bash scripts/excel_goalseek.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
GS="$ROOT/workbooks/goalseek"
OUTDIR="$GS/excel"
mkdir -p "$OUTDIR" "$ROOT/results"
TMP="$(mktemp -d)"  # only for this script's own notes; Excel never opens files here

seek() { # id file setSheet setCell goal chSheet chCell
  local id=$1 f=$2 ss=$3 sc=$4 goal=$5 cs=$6 cc=$7
  local work="${OUTDIR:?}/${f%.xlsx}_excel.xlsx"   # work in place: Excel's sandbox asks for file access on new folders, which blocks a script
  cp "$GS/$f" "$work"
  osascript <<EOF > "$TMP/$id.txt"
tell application "Microsoft Excel"
  set wb to open workbook workbook file name (POSIX file "$work" as text)
  set ws to worksheet "$cs" of wb
  activate object ws
  set onOpen to string value of range "$sc" of worksheet "$ss" of wb
  calculate full
  set startVal to string value of range "$sc" of worksheet "$ss" of wb
  set ok to goal seek (range "$sc" of worksheet "$ss" of wb) goal $goal changing cell (range "$cc" of ws)
  set a to value of range "$cc" of ws
  set o to string value of range "$sc" of worksheet "$ss" of wb
  if ok then save wb
  close wb saving no
  return (ok as text) & "|" & (a as text) & "|" & o & "|" & startVal & "|" & onOpen
end tell
EOF
  echo "$id: $(cat "$TMP/$id.txt")"
  case "$(cat "$TMP/$id.txt")" in true*) ;; *) rm -f "${work:?}";; esac  # Excel found no solution: save nothing
}

rm -f "${OUTDIR:?}"/*_excel.xlsx
# The Set cell is the helper cell when the target is on another sheet (Excel wants them on one sheet).
seek gs1 gs1_three_statement_growth.xlsx Assumptions B19 30 Assumptions C5
seek gs2 gs2_three_statement_terminal_growth.xlsx Valuation B13 30 Valuation B4
seek gs3 gs3_three_statement_wacc.xlsx Valuation B13 30 Valuation B3
seek gs4 gs4_lbo_margin_step.xlsx Inputs B19 0.2 Inputs B10

EXCEL_VER=$(defaults read "/Applications/Microsoft Excel.app/Contents/Info.plist" CFBundleVersion)
python3 - "$TMP" "$ROOT/results/stage3_goalseek_excel_run.json" "$EXCEL_VER" <<'PY'
import json, sys, platform, subprocess, datetime
tmp, out, ver = sys.argv[1:4]
cases = []
for cid in ["gs1", "gs2", "gs3", "gs4"]:
    ok, a, o, start, on_open = open(f"{tmp}/{cid}.txt").read().rstrip("\n").split("|")
    cases.append({"id": cid, "excelReportedSuccess": ok == "true", "changingCellAfter": a,
                  "setCellShownOnOpen": on_open, "setCellShownBefore": start, "setCellShownAfter": o, "savedFile": ok == "true"})
json.dump({
    "generatedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    "machine": {"what": "Alek's Mac", "cpu": subprocess.run(["sysctl", "-n", "machdep.cpu.brand_string"], capture_output=True, text=True).stdout.strip(),
                "os": "macOS " + platform.mac_ver()[0], "excel": "Microsoft Excel for Mac " + ver},
    "method": "AppleScript `goal seek` (Range.GoalSeek), iterative calc on, Maximum Change 1e-9, 1000 iterations",
    "note": "Numbers here are AppleScript's ~12-digit display. Full-precision answers are read from the saved files by `npm run stage3` (results/stage3_backsolve.json goalSeek[].excel).",
    "cases": cases,
}, open(out, "w"), indent=1)
print("wrote", out)
PY
rm -rf "${TMP:?}"
