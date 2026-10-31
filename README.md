# Backsolve

Turn an Excel valuation model into a differentiable graph and solve it backwards from a price to
the assumptions that price requires. Runs in the browser; workbooks never leave your computer.

**Status: Stage 1 of 5 (formula engine).** See [`docs/stage1.md`](docs/stage1.md) and [`LAUNCH_FACTS.md`](LAUNCH_FACTS.md).

## Quick start

```bash
npm install
npm test                         # unit tests + every test workbook vs. its cached values
npx tsx scripts/check.ts         # same check, with a full report -> results/stage1_forward.json
pip install formulas openpyxl
python3 scripts/crosscheck_formulas.py   # independent engine -> results/stage1_crosscheck_formulas.json
python3 scripts/make_workbooks.py        # rebuild test workbooks (needs LibreOffice)
```

Check any other workbook: `npx tsx scripts/check.ts path/to/model.xlsx`.

## Layout

```
src/engine/   parser, values, functions, workbook (graph + recalculation), compare
src/xlsx/     .xlsx reader (zip + XML; works in browser and Node)
scripts/      test-workbook builder, LibreOffice recalc, checks, formulas cross-check
workbooks/    test workbooks (+ known_differences.json)
results/      machine-readable outputs every published number comes from
tests/        vitest suites
```
