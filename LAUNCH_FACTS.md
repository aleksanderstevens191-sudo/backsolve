# LAUNCH_FACTS

Every number we might quote publicly, with the file it comes from and the machine that produced it.
If a number is not in this table, don't post it.

## Machines

| Id | What | Details |
|---|---|---|
| M2 | Alek's Mac (ground truth only) | Microsoft Excel for Mac 16.0300 recalculated and saved the three test workbooks on 2026-11-05 (iterative calc on, 1000 iterations, max change 1e-9) |
| M3 | Alek's Mac (Stage 3 Excel Goal Seek run, 2026-11-08) | Apple M4 Pro (14 cores), macOS 27.0.1, Microsoft Excel for Mac 16.84 (build 16.84.24041420) driven by `scripts/excel_goalseek.sh`, Node v24.22.1 for `npm run stage3` |
| M1 | Cloud dev container (Stage 1, 2 and 3 runs) | Linux 6.18.44, Intel Xeon (4 cores), Node v22.23.4, Python 3.13.18, `formulas` 1.3.4, LibreOffice 24.2.7.2 |

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
| Speed-up, tape vs one recalc per input, three-statement (66 inputs) | 28× (14.8 ms vs 0.528 ms) | `results/stage2_gradients.json` (`benchmark`) | M1 | ✅ ±30% run to run |
| Speed-up, synthetic DCF with 2,105 inputs | 874× (3,075 ms vs 3.52 ms) | `results/stage2_gradients.json` | M1 | ✅ synthetic model, no Excel file |
| Unit tests | 49 Stage 1 + 21 Stage 2, all passing | `npm test` | M1 | ✅ |
| Gradient check on CFI DCF / SEC workbooks | — | — | — | ⏳ need the files |

## Stage 3: solving backwards from a price

Rules: target hit when `|output − target| ≤ 1e-12 × max(1, |target|)`; closest point found when tangency (sine of the angle between the move and the price gradient, in σ units) ≤ 1e-9; one-input answers vs Goal Seek: `|ours − Goal Seek| ≤ 1e-6 × max(1, |Goal Seek|)`. Weighting and checks explained in `docs/stage3.md`.

Note on machines: rows marked M1 come from the cloud run of `results/stage3_backsolve.json` at commit edf3cd5 (`git show edf3cd5:results/stage3_backsolve.json`). The file was regenerated on M3 (Alek's Mac) on 2026-11-08 for the Excel comparison. On M3 the answers agree with M1 to about 1e-15 (last-digit rounding differences between machines), and the timings are different (M3 is faster). The rows marked M3 quote the current file.

| Fact | Value | Source | Machine | Status |
|---|---|---|---|---|
| One free input matches **Excel's** Goal Seek to 1e-6 | 3 / 3 cases where Excel finds an answer (gs1–gs3); differences 3.1e-16, −7.0e-13, −5.4e-14 | `results/stage3_backsolve.json` (`goalSeek[].excel`), files in `workbooks/goalseek/excel/` | M3 | ✅ |
| Every formula cell at Excel's Goal Seek answer matches Excel's cached values (1e-9) | 243/243, 242/242, 242/242 (gs1 includes the helper cell) | `results/stage3_backsolve.json` (`goalSeek[].excel.forwardCheckAtExcelsAnswer`) | M3 | ✅ |
| Excel's own Goal Seek on the LBO (gs4: margin step for 20% IRR, 5 loops + IRR) | Excel reports no solution: it pushes the margin step to 9.0e+70 and the IRR stays at 19.89%. Same result from the rebuilt case file (opens cleanly, IRR 19.89%) and from `lbo_circular.xlsx` with the helper added (2026-11-08). Backsolve answers 0.0051293 with IRR missing 20% by 2.5e-16 | `results/stage3_goalseek_excel_run.json` (`cases[3]`), `results/stage3_backsolve.json` (`goalSeek[3]`) | M3 | ✅ (Excel limitation; cause not proven) |
| Goal Seek case files keep Excel's saved numbers | every formula cell matches its cached value: 243/243, 242/242, 242/242, 93/93; Excel shows the target cell on open as $26.10 / $26.10 / $26.10 / 19.89% (gs4 used to open with #VALUE! when the files were written by openpyxl) | `results/stage3_backsolve.json` (`goalSeek[].caseFileCachedValues`), `results/stage3_goalseek_excel_run.json` (`cases[].setCellShownOnOpen`) | M3 | ✅ fixed 2026-11-09 |
| One free input vs LibreOffice's Goal Seek (second opinion, not ground truth) | 4 / 4 within 1e-6; differences −2.2e-16, 3.2e-12, −1.5e-13, 6.9e-7 | `results/stage3_backsolve.json` (`goalSeek[].libreoffice`), `results/stage3_goalseek_libreoffice.json` | M1 (LibreOffice 24.2.7.2) | ✅ |
| LBO Goal Seek (margin step for 20% IRR, 5 loops + IRR): output at our answer vs LibreOffice's | ours misses 20% by 2.5e-16; LibreOffice's own IRR at its answer misses by 6.0e-6 | `results/stage3_backsolve.json` (`goalSeek[3]`) | M1 | ✅ |
| Closest $30 point, three-statement, growth + gross margin, σ = 1 pp each | growth +1.357 pp (avg 6.00% → 7.36%), gross margin +1.684 pp (avg 41.60% → 43.28%), distance 2.163 | `results/stage3_backsolve.json` (`closestPoint[0]`) | M1 | ✅ |
| Same, margins stickier (σ_margin = 0.5 pp) | growth +2.470 pp, gross margin +0.814 pp, distance 2.958 | `closestPoint[1]` | M1 | ✅ |
| Same, four drivers (growth, gross margin, SG&A, capex; σ = 1 pp) | +0.530, +0.629, −0.627, −0.836 pp; distance 1.330 | `closestPoint[2]` | M1 | ✅ |
| LBO, IRR 19.89% → 20%, growth + margin step + exit multiple | +0.011 pp, +0.0034 pp, +0.020x; distance 0.0265 | `closestPoint[3]` | M1 | ✅ |
| Closest-point answers hit the target | worst miss 8.1e-13 (LBO IRR); tangency ≤ 7.7e-10; curvature along the price surface 1.0 to 1.04 (positive: a true minimum) | `closestPoint[].result` | M1 | ✅ |
| Brute force finds nothing closer | 2 drivers: closest traced point 2.16514 vs 2.16275 and 2.95795 vs 2.95764; 3+ drivers: 360 rays (131 reached) closest 1.33135 vs 1.32998, 210 rays (114 reached) closest 0.026526 vs 0.026516 | `closestPoint[].bruteForce` | M1 | ✅ |
| Every growth/margin pair for $30, three-statement (WACC 8 / 9 / 10%) | 122 / 130 / 137 points, one piece each; grid check 62/62, 66/66, 70/70 crossings covered; worst miss 5.3e-14 | `curves[0]`, `waccPlusMinus1pp.curves` | M1 | ✅ |
| Every growth/margin-step pair for 20% IRR, LBO | 120 points, one piece; 68/68 grid crossings covered; worst miss 3.5e-13 | `curves[1]` | M1 | ✅ |
| Implied average growth for $30 (growth only) at WACC 8% / 9% / 10% | 4.67% / 9.57% / 14.07% (base case 6.00%) | `waccPlusMinus1pp.summary` | M1 | ✅ |
| Implied growth move for WACC −1 pp / +1 pp | −4.90 pp / +4.49 pp | `waccPlusMinus1pp.summary` | M1 | ✅ |
| Slope of implied growth vs WACC (implicit function theorem) | 5.15 / 4.68 / 4.32 pp per pp at 8 / 9 / 10%; vs finite differences: 1.1e-11, 7.4e-15, 9.8e-13 relative | `waccPlusMinus1pp.slopeChecks` | M1 | ✅ |
| Scale, synthetic DCF, target +15% | 20 drivers: 3 Newton steps, 166 gradients, 113 ms; 200 drivers: 7 projection steps, 10 gradients, 31 ms | `scale` | M1 | ✅ single run, ±30% |
| Unit tests | 49 Stage 1 + 21 Stage 2 + 26 Stage 3 = 96, all passing | `npm test` | M1, M3 | ✅ |

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
- "Matches Excel's Goal Seek" without the qualifier: true for gs1–gs3. On the circular LBO (gs4) Excel's Goal Seek finds no answer, so there is nothing to match. Say "matches Excel's Goal Seek wherever Excel finds an answer".
- "Excel's Goal Seek can't handle circular models" in general: shown on one model (gs4). The cause is a guess, not proven.
- "The market implies X% growth" as a single fact: it's the closest answer under the chosen weights (σ), and it moves about 4.5 to 5 pp per 1 pp of WACC. Always quote the σ's and the WACC with it.
- "Finds the closest point on any model": the brute-force proof runs on these test cases in `npm run stage3`, not on every solve.
