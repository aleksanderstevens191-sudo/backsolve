# LAUNCH_FACTS

Every number we might quote publicly, with the file it comes from and the machine that produced it.
If a number is not in this table, don't post it.

## Machines

| Id | What | Details |
|---|---|---|
| M2 | Alek's Mac (ground truth only) | Microsoft Excel for Mac 16.0300 recalculated and saved the three test workbooks on 2026-11-05 (iterative calc on, 1000 iterations, max change 1e-9) |
| M1 | Cloud dev container (Stage 1 runs) | Linux 6.18.44, Intel Xeon (4 cores), Node v22.23.4, Python 3.13.18, `formulas` 1.3.4, LibreOffice 24.2.7.2 |

## Stage 1: forward pass (formula engine)

Tolerance everywhere: `|ours − cached| ≤ 1e-9 × max(1, |cached|)` (relative for big numbers, absolute near zero).

| Fact | Value | Source | Machine | Status |
|---|---|---|---|---|
| Formula cells matching **Excel's** cached values, three-statement model | 242 / 242 | `results/stage1_forward.json` | M1 (engine), M2 (Excel) | ✅ cached by Microsoft Excel |
| Worst difference vs Excel, three-statement model | 2.7e-16 (scaled), 4.6e-13 absolute | `results/stage1_forward.json` | M1 | ✅ |
| Formula cells matching **Excel's** cached values, circular LBO | 92 / 92 | `results/stage1_forward.json` | M1, M2 | ✅ cached by Microsoft Excel |
| Worst difference vs Excel, circular LBO | 1.0e-10 (scaled), 5.6e-10 absolute | `results/stage1_forward.json` | M1 | ✅ (engine stops at the file's own max-change setting, 1e-9) |
| LBO circular blocks solved from zero | 5 blocks of 7 cells, 29 iterations each, residual ≤ 1.2e-10 | `results/stage1_forward.json` | M1 | ✅ |
| Formula cells matching **Excel's** cached values, function coverage | 64 / 64, including the 6 cells where LibreOffice differs (Excel's cache agrees with the documented rule) | `results/stage1_forward.json` | M1, M2 | ✅ no overrides applied |
| `formulas` library agrees with our engine, three-statement | 242 / 242 | `results/stage1_crosscheck_formulas.json` | M1 | ✅ |
| `formulas` library agrees with our engine, function coverage | 64 / 64 | `results/stage1_crosscheck_formulas.json` | M1 | ✅ |
| `formulas` library agrees with our engine, circular LBO | 44 / 92 (all 48 disagreements are in or downstream of a loop; `formulas` cuts loops instead of iterating) | `results/stage1_crosscheck_formulas.json` | M1 | ⚠️ expected |
| Unit tests | 46 engine + 3 workbook acceptance, all passing | `npm test` | M1 | ✅ |
| Unsupported functions hit in test workbooks | none | `results/stage1_forward.json` (`unsupported`) | M1 | ✅ |
| CFI DCF workbook | — | — | — | ⏳ waiting for Alek's file |
| SEC-generated DCF workbook | — | — | — | ⏳ Stage 4 |
| Excel-cached versions of the test workbooks | 3 / 3 | `workbooks/*.xlsx` (docProps/app.xml: Microsoft Macintosh Excel) | M2 | ✅ 2026-11-05 |

## Model outputs (reference only, not claims)

| Fact | Value | Source |
|---|---|---|
| Three-statement model value per share (`Valuation!B13`) | 26.0989490957 | `results/engine_values/three_statement.json` |
| LBO IRR (`Returns!B8`) | 19.888% | `results/engine_values/lbo_circular.json` |
| LBO MOIC (`Returns!B5`) | 2.4767x | `results/engine_values/lbo_circular.json` |

## Things that are NOT true yet (don't post)

- "Matches Excel" on any workbook other than these three test files (the CFI DCF and SEC model still need their own runs).
- "Cross-checked against `formulas` on the LBO": `formulas` cannot iterate circular references, so it only confirms the 44 cells upstream of the loops.
