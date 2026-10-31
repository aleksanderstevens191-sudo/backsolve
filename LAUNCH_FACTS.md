# LAUNCH_FACTS

Every number we might quote publicly, with the file it comes from and the machine that produced it.
If a number is not in this table, don't post it.

## Machines

| Id | What | Details |
|---|---|---|
| M1 | Cloud dev container (Stage 1 runs) | Linux 6.18.44, Intel Xeon (4 cores), Node v22.23.4, Python 3.13.18, `formulas` 1.3.4, LibreOffice 24.2.7.2 |

## Stage 1: forward pass (formula engine)

Tolerance everywhere: `|ours − cached| ≤ 1e-9 × max(1, |cached|)` (relative for big numbers, absolute near zero).

| Fact | Value | Source | Machine | Status |
|---|---|---|---|---|
| Formula cells matching cached values, three-statement model | 242 / 242 | `results/stage1_forward.json` | M1 | ✅ cached by **LibreOffice**, not Excel |
| Worst difference, three-statement model | 3.7e-15 (scaled), 4.1e-12 absolute | `results/stage1_forward.json` | M1 | ✅ |
| Formula cells matching cached values, circular LBO | 92 / 92 | `results/stage1_forward.json` | M1 | ✅ cached by **LibreOffice** (run to convergence), not Excel |
| Worst difference, circular LBO | 3.1e-13 (scaled), 2.3e-12 absolute | `results/stage1_forward.json` | M1 | ✅ |
| LBO circular blocks solved from zero | 5 blocks of 7 cells, 35–36 iterations each, residual ≤ 5.7e-13 | `results/stage1_forward.json` | M1 | ✅ |
| Formula cells matching, function-coverage workbook | 64 / 64 (6 checked against Excel's documented rule where LibreOffice differs) | `results/stage1_forward.json`, `workbooks/known_differences.json` | M1 | ✅ LibreOffice-cached |
| `formulas` library agrees with our engine, three-statement | 242 / 242 | `results/stage1_crosscheck_formulas.json` | M1 | ✅ |
| `formulas` library agrees with our engine, function coverage | 64 / 64 | `results/stage1_crosscheck_formulas.json` | M1 | ✅ |
| `formulas` library agrees with our engine, circular LBO | 44 / 92 (all 48 disagreements are in or downstream of a loop; `formulas` cuts loops instead of iterating) | `results/stage1_crosscheck_formulas.json` | M1 | ⚠️ expected |
| Unit tests | 46 engine + 3 workbook acceptance, all passing | `npm test` | M1 | ✅ |
| Unsupported functions hit in test workbooks | none | `results/stage1_forward.json` (`unsupported`) | M1 | ✅ |
| CFI DCF workbook | — | — | — | ⏳ waiting for Alek's file |
| SEC-generated DCF workbook | — | — | — | ⏳ Stage 4 |
| Excel-cached versions of the test workbooks | — | — | — | ⏳ needs a re-save in Microsoft Excel |

## Model outputs (reference only, not claims)

| Fact | Value | Source |
|---|---|---|
| Three-statement model value per share (`Valuation!B13`) | 26.0989490957 | `results/engine_values/three_statement.json` |
| LBO IRR (`Returns!B8`) | 19.888% | `results/engine_values/lbo_circular.json` |
| LBO MOIC (`Returns!B5`) | 2.4767x | `results/engine_values/lbo_circular.json` |

## Things that are NOT true yet (don't post)

- "Matches Excel": the test workbooks' cached values came from LibreOffice. Only say "matches Excel" for files saved by Excel (the CFI DCF, once added, or the test files after an Excel re-save).
- "Cross-checked against `formulas` on the LBO": `formulas` cannot iterate circular references, so it only confirms the 44 cells upstream of the loops.
