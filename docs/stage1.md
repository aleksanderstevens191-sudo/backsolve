# Stage 1: the formula engine

## What it does

Reads an `.xlsx` (formulas and the values Excel saved with them), rebuilds every formula as a
tree, links the cells into a dependency graph, and recalculates the whole workbook from scratch.
Then it compares every formula cell with the value cached in the file.

Supported: `+ - * / ^ & %`, comparisons, `SUM`, `SUMPRODUCT`, `AVERAGE`, `MIN`, `MAX`, `IF`, `AND`,
`OR`, `NPV`, `XNPV`, `IRR`, plus small extras (`ABS`, `ROUND`, `NOT`, `IFERROR`, `COUNT`, `TRUE()`,
`FALSE()`). Absolute/relative/mixed references, cross-sheet and quoted sheet names, whole-column
ranges, workbook- and sheet-scoped named ranges (including names that are formulas), shared
formulas (how Excel stores a formula filled down), array formulas, and array constants.

Not supported, and reported (never skipped): any other function (cell shows `#NAME?` and is listed
under `unsupported` in the results), structured table refs (`Table1[Col]`), 3-D refs
(`Sheet1:Sheet3!A1`), external workbook links, What-If data tables, and dynamic references such as
`OFFSET`/`INDIRECT` (they will be listed as unsupported functions).

## Circular references, in plain English

An LBO often computes interest on the *average* debt balance. But the ending balance depends on
how much cash was left to repay debt, which depends on interest. Each cell needs the other one
first: a loop.

Excel's "Enable iterative calculation" setting handles this by guessing, recalculating the loop,
and repeating until no cell changes by more than *Maximum Change* (or it hits *Maximum
Iterations*). We do the same:

1. Find the loops. Tarjan's algorithm groups cells that can reach each other into "strongly
   connected components". A group of more than one cell (or a cell that reads itself) is a loop.
2. Everything else is evaluated in dependency order. Each loop is evaluated as one block.
3. For a loop block we start every cell at **zero** (not at the saved answer, which would be
   cheating) and sweep until the biggest change is below the workbook's Max Change setting.
4. We then check the answer really is a fixed point: plug it back in, recalculate once, and
   confirm nothing moves (residual ≤ 6e-13 on the LBO).

If the workbook has iterative calculation turned off, loop cells become `#CIRC!` and are reported.

## Where the "ground truth" values came from (important)

There is no Microsoft Excel in the build machine, so the three test workbooks were generated with
openpyxl and their cached values were computed by **LibreOffice 24.2**, not Excel. Two findings:

- **LibreOffice's single load-time recalculation does not finish circular loops.** Its cached LBO
  values had fixed-point residuals around 1.0 (for example interest income of 1.68 where the
  formula gives 0.60). Recalculating repeatedly (10 passes) until nothing changed fixed it, and the
  converged values match ours to 3e-13. `scripts/lo_recalc.py` does this.
- **LibreOffice differs from Excel on TRUE inside ranges.** LibreOffice has no true/false cell
  type, so it counts TRUE as 1 in `SUM(range)`, `AVERAGE`, `NPV`, `COUNT`, and it rejects
  `SUM(1,"2",TRUE)`. Excel's documented rules (and the independent `formulas` library) disagree
  with LibreOffice here and agree with us. Those 6 cells are listed in
  `workbooks/known_differences.json` and checked against Excel's rule; that override only applies
  while the file's cached values come from LibreOffice.

Re-saving the three test files in real Excel would make Excel their ground truth.

## Cross-check with the Python `formulas` library

`formulas` agrees with us on 242/242 cells of the three-statement model and 64/64 of the coverage
workbook. On the LBO it agrees on the 44 cells upstream of the loops and differs on 48 of the 52
cells in or after them, because `formulas` handles a loop by cutting it (one input becomes an
error/zero) rather than iterating. That's a limitation of `formulas`, not evidence either way.

## Workbooks

| Workbook | Built by | Cached by | Status |
|---|---|---|---|
| `three_statement.xlsx` (IS, BS, CF, DCF; revolver; names; cross-sheet) | `scripts/make_workbooks.py` | LibreOffice | 242/242 ✅ |
| `lbo_circular.xlsx` (interest on average debt and cash; IRR, XNPV, NPV) | `scripts/make_workbooks.py` | LibreOffice, run to convergence | 92/92 ✅ |
| `functions_coverage.xlsx` (every function + Excel edge cases) | `scripts/make_workbooks.py` | LibreOffice | 64/64 ✅ |
| CFI DCF | Alek | Excel | ⏳ need the file |
| SEC-generated DCF | Stage 4 | — | ⏳ |

## Unsupported functions hit

None in the three test workbooks. (LibreOffice writes TRUE/FALSE cells as `=TRUE()`/`=FALSE()`
formulas; we hit those, so we added them.)
