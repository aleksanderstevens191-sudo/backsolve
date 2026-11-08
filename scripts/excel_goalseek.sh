#!/bin/bash
# Run Microsoft Excel's own Goal Seek on the four Stage 3 cases, on a Mac with Excel installed.
#
# What it does, case by case (cases listed in workbooks/goalseek/cases.json):
#   gs1-gs3: copy workbooks/goalseek/gsN_*.xlsx to workbooks/goalseek/excel/gsN_*_excel.xlsx, open it
#            in Excel, run Goal Seek (AppleScript `goal seek`, the same engine as Data > What-If
#            Analysis > Goal Seek), and save. `npm run stage3` then reads the saved files.
#   gs4:     the gs4 file was written by openpyxl, which drops cached values, so Excel recalculates
#            the circular LBO from scratch on open. When it does, a #VALUE! gets into the interest loop
#            and stays there (an error inside a circular loop feeds itself), so the IRR shows #VALUE!.
#            We therefore start from the Excel-saved workbooks/lbo_circular.xlsx (converged cached
#            values), add the same helper cell (Inputs!B19 = Returns!B8) and note, and Goal Seek that.
#            If Excel reports success the file is saved as gs4_..._excel.xlsx; if not, it is deleted.
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

seek() { # id src dst setSheet setCell goal chSheet chCell readSheet readCell prep(0/1)
  local id=$1 src=$2 dst=$3 ss=$4 sc=$5 goal=$6 cs=$7 cc=$8 rs=$9 rc=${10} prep=${11}
  local work="$dst"   # work in place: Excel's sandbox asks for file access on new folders, which blocks a script
  cp "$src" "$work"
  local note=""
  if [ "$prep" = 1 ]; then
    note=$(python3 -c "from openpyxl import load_workbook;print(load_workbook('$GS/gs4_lbo_margin_step.xlsx')['Inputs']['A20'].value)")
  fi
  osascript <<EOF > "$TMP/$id.txt"
tell application "Microsoft Excel"
  set wb to open workbook workbook file name (POSIX file "$work" as text)
  set ws to worksheet "$cs" of wb
  activate object ws
  if $prep = 1 then
    set value of range "A19" of ws to "Goal Seek target cell (repeats Returns!B8)"
    set formula of range "B19" of ws to "=Returns!B8"
    set value of range "A20" of ws to "$note"
  end if
  calculate full
  set startVal to string value of range "$rc" of worksheet "$rs" of wb
  set ok to goal seek (range "$sc" of worksheet "$ss" of wb) goal $goal changing cell (range "$cc" of ws)
  set a to value of range "$cc" of ws
  set o to string value of range "$rc" of worksheet "$rs" of wb
  if ok then save wb
  close wb saving no
  return (ok as text) & "|" & (a as text) & "|" & o & "|" & startVal
end tell
EOF
  echo "$id: $(cat "$TMP/$id.txt")"
  case "$(cat "$TMP/$id.txt")" in true*) ;; *) rm -f "$dst";; esac  # Excel found no solution: save nothing
}

rm -f "$OUTDIR"/*_excel.xlsx
seek gs1 "$GS/gs1_three_statement_growth.xlsx" "$OUTDIR/gs1_three_statement_growth_excel.xlsx" Valuation B13 30 Assumptions C5 Valuation B13 0
seek gs2 "$GS/gs2_three_statement_terminal_growth.xlsx" "$OUTDIR/gs2_three_statement_terminal_growth_excel.xlsx" Valuation B13 30 Valuation B4 Valuation B13 0
seek gs3 "$GS/gs3_three_statement_wacc.xlsx" "$OUTDIR/gs3_three_statement_wacc_excel.xlsx" Valuation B13 30 Valuation B3 Valuation B13 0
seek gs4 "$ROOT/workbooks/lbo_circular.xlsx" "$OUTDIR/gs4_lbo_margin_step_excel.xlsx" Inputs B19 0.2 Inputs B10 Returns B8 1

# gs4 as shipped (openpyxl file, no cached values): what Excel shows on open
RAW=$(osascript <<EOF
tell application "Microsoft Excel"
  set wb to open workbook workbook file name (POSIX file "$GS/gs4_lbo_margin_step.xlsx" as text)
  set r to (string value of range "B8" of worksheet "Returns" of wb) & "|" & (string value of range "C10" of worksheet "Model" of wb)
  close wb saving no
  return r
end tell
EOF
)
echo "gs4 file as shipped, on open: IRR, Model!C10 = $RAW"

EXCEL_VER=$(defaults read "/Applications/Microsoft Excel.app/Contents/Info.plist" CFBundleVersion)
python3 - "$TMP" "$ROOT/results/stage3_goalseek_excel_run.json" "$EXCEL_VER" "$RAW" <<'PY'
import json, sys, platform, subprocess, datetime
tmp, out, ver, raw = sys.argv[1:5]
cases = []
for cid in ["gs1", "gs2", "gs3", "gs4"]:
    ok, a, o, start = open(f"{tmp}/{cid}.txt").read().rstrip("\n").split("|")
    cases.append({"id": cid, "excelReportedSuccess": ok == "true", "changingCellAfter": a,
                  "setCellShownAfter": o, "setCellShownBefore": start, "savedFile": ok == "true"})
irr, c10 = raw.split("|")
json.dump({
    "generatedAt": datetime.datetime.now(datetime.timezone.utc).isoformat(),
    "machine": {"what": "Alek's Mac", "cpu": subprocess.run(["sysctl", "-n", "machdep.cpu.brand_string"], capture_output=True, text=True).stdout.strip(),
                "os": "macOS " + platform.mac_ver()[0], "excel": "Microsoft Excel for Mac " + ver},
    "method": "AppleScript `goal seek` (Range.GoalSeek), iterative calc on, Maximum Change 1e-9, 1000 iterations",
    "note": "Numbers here are AppleScript's ~12-digit display. Full-precision answers are read from the saved files by `npm run stage3` (results/stage3_backsolve.json goalSeek[].excel).",
    "cases": cases,
    "gs4FileAsShipped": {"irrOnOpen": irr, "modelC10OnOpen": c10,
                          "why": "openpyxl drops cached values; Excel recalculates the circular LBO from scratch and an error gets stuck in the interest loop"},
}, open(out, "w"), indent=1)
print("wrote", out)
PY
rm -rf "$TMP"
