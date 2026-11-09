# Backsolve

Turn an Excel valuation model into a differentiable graph and solve it backwards from a price to
the assumptions that price requires. Runs in the browser; workbooks never leave your computer.

**Status: Stage 3 of 5 (solving backwards).** See [`docs/stage1.md`](docs/stage1.md) (formula engine), [`docs/stage2.md`](docs/stage2.md) (reverse-mode gradients, IRR and loops via the implicit function theorem, kink report), [`docs/stage3.md`](docs/stage3.md) (from a target price to the closest assumptions that hit it exactly, the full curve of answers, Goal Seek, WACC ±1 pp) and [`LAUNCH_FACTS.md`](LAUNCH_FACTS.md).

## Quick start

```bash
npm install
npm test                         # unit tests + every test workbook vs. its cached values
npx tsx scripts/check.ts         # same check, with a full report -> results/stage1_forward.json
npm run stage2                   # every gradient vs finite differences + benchmark -> results/stage2_gradients.json
npm run stage3                   # backsolve checks (Goal Seek, closest point, curves, WACC ±1 pp) -> results/stage3_backsolve.json
npm run stage3:chart             # redraw docs/stage3_curve.svg from that file
pip install formulas openpyxl
python3 scripts/crosscheck_formulas.py   # independent engine -> results/stage1_crosscheck_formulas.json
python3 scripts/make_workbooks.py        # rebuild test workbooks (needs LibreOffice)
python3 scripts/make_goalseek.py         # rebuild the Goal Seek case files for Excel (workbooks/goalseek/)
/usr/bin/python3 scripts/lo_goalseek.py  # LibreOffice's Goal Seek on those cases -> results/stage3_goalseek_libreoffice.json
```

Check any other workbook: `npx tsx scripts/check.ts path/to/model.xlsx`.

## Layout

```
src/engine/   parser, values, functions, workbook (graph + recalculation), compare,
              tape + ad + grad (reverse-mode gradients, kink report)
              solve (backsolver: Goal Seek, closest point, curve tracing)
src/xlsx/     .xlsx reader (zip + XML; works in browser and Node)
scripts/      test-workbook builder, LibreOffice recalc, checks, formulas cross-check
workbooks/    test workbooks (+ known_differences.json); goalseek/ holds the Goal Seek case files
results/      machine-readable outputs every published number comes from
tests/        vitest suites
```
