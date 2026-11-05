# LAUNCH_FACTS

Every number we might quote publicly, with the file it comes from and the machine that produced it.
If a number is not in this table, don't post it.

## Machines

| Id | What | Details |
|---|---|---|
| M2 | Alek's Mac (ground truth only) | Microsoft Excel for Mac 16.0300 recalculated and saved the three test workbooks on 2026-11-05 (iterative calc on, 1000 iterations, max change 1e-9) |
| M1 | Cloud dev container (Stage 1 and 2 runs) | Linux 6.18.44, Intel Xeon (4 cores), Node v22.23.4, Python 3.13.18, `formulas` 1.3.4, LibreOffice 24.2.7.2 |

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

## Stage 2: gradients (reverse-mode tape)

Pass rule: `|AD − FD| ≤ 1e-6 × max(|AD|, |FD|)`; FD = central differences at h and h/2, Richardson-extrapolated, h = 1e-4 × max(1, |input|). Every numeric formula cell × every input cell. Buckets explained in `docs/stage2.md`.

| Fact | Value | Source | Machine | Status |
|---|---|---|---|---|
| Tape value = engine value, every formula cell | 0 mismatches on all 3 test workbooks (bit for bit) | `results/stage2_gradients.json` (`tape.forwardMismatches`) | M1 | ✅ |
| Gradient pairs checked, three-statement | 15,114: 12,744 both exactly 0, 2,199 within 1e-6, 171 within FD rounding noise (AD ≤ 1.2e-16), 0 failed | `results/stage2_gradients.json` | M1 | ✅ |
| Gradient pairs checked, circular LBO (incl. IRR + 5 loops) | 1,424: 841 exactly 0, 498 within 1e-6, 84 FD noise (AD = 0), 1 excluded (Year-0 counter on an `IF(year=5)` kink), 0 failed | `results/stage2_gradients.json` | M1 | ✅ |
| Gradient pairs checked, function coverage | 846: 748 exactly 0, 88 within 1e-6, 10 excluded (XNPV integer dates), 0 failed | `results/stage2_gradients.json` | M1 | ✅ |
| Worst relative error among passing pairs (3 test workbooks) | 6.7e-9 | `results/stage2_gradients.json` | M1 | ✅ |
| Loop fixed-point residual before differentiating (LBO) | ≤ 2.2e-16 (scaled) | `results/stage2_gradients.json` (`tape.loops`) | M1 | ✅ |
| Speed-up, tape vs one recalc per input, three-statement (66 inputs) | 22× (14.9 ms vs 0.671 ms) | `results/stage2_gradients.json` (`benchmark`) | M1 | ✅ ±30% run to run |
| Speed-up, synthetic DCF with 2,105 inputs | 761× (2,696 ms vs 3.54 ms) | `results/stage2_gradients.json` | M1 | ✅ synthetic model, no Excel file |
| Unit tests | 49 Stage 1 + 16 Stage 2, all passing | `npm test` | M1 | ✅ |
| Gradient check on CFI DCF / SEC workbooks | — | — | — | ⏳ need the files |

## Model outputs (reference only, not claims)

| Fact | Value | Source |
|---|---|---|
| Three-statement model value per share (`Valuation!B13`) | 26.0989490957 | `results/engine_values/three_statement.json` |
| LBO IRR (`Returns!B8`) | 19.888% | `results/engine_values/lbo_circular.json` |
| LBO MOIC (`Returns!B5`) | 2.4767x | `results/engine_values/lbo_circular.json` |
| d(value per share)/d(WACC), three-statement | −468.07 per 1.00 of WACC (≈ −$4.68 per +1 pp) | `results/stage2_gradients.json` (`headline.gradient`) |
| d(LBO IRR)/d(revenue growth) | +1.785 per 1.00 (≈ +1.8 pp IRR per +1 pp growth) | `results/stage2_gradients.json` |

## Things that are NOT true yet (don't post)

- "Gradients exact on any model": checked on these 3 test workbooks and 2 synthetic models only.
- Speed-ups as a fixed number: they are single-machine medians that vary by about ±30% between runs.

- "Matches Excel" on any workbook other than these three test files (the CFI DCF and SEC model still need their own runs).
- "Cross-checked against `formulas` on the LBO": `formulas` cannot iterate circular references, so it only confirms the 44 cells upstream of the loops.
