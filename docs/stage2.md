# Stage 2: running the workbook backwards

## What it does

Given any output cell (the share price, an IRR, anything numeric), Stage 2 computes its derivative
with respect to **every input cell at once**: "if this assumption goes up by a tiny amount, how
much does the price move?" Those derivatives are what Stage 3 uses to solve for the assumptions
that hit a target price.

```ts
import { loadWorkbook, Gradients } from './src/index';
const { wb } = loadWorkbook(bytes);
wb.recalc();
const g = new Gradients(wb);
g.gradientByAddr('Valuation!B13');   // { 'Valuation!B3': -468.07, 'Assumptions!B13': -40.38, ... }
g.kinkReport();                       // which IF / MIN / MAX decisions are close to flipping
```

Run every check with `npm run stage2` (writes `results/stage2_gradients.json`).

## How it works, in plain English

### The tape (reverse-mode automatic differentiation)

While recalculating, we write down every arithmetic step that depends on an input, plus its
*local slope*. For `C1 = A1 * B1` the tape gets one line: "C1 = A1 × B1; slope to A1 is B1's
value, slope to B1 is A1's value". A whole workbook becomes a few hundred such lines.

To get the gradient of the price, we put a 1 on the price's line and walk the tape **backwards**,
passing each line's number to its parents multiplied by the local slope (the chain rule). When we
reach the inputs, each one holds d(price)/d(input). One backward walk serves every input, so the
cost doesn't grow with the number of inputs; the obvious alternative (nudge one input, recalculate,
repeat) costs one full recalculation per input.

The tape is recorded by a second evaluator (`src/engine/ad.ts`) that copies the engine's Excel
rules. To prove the two agree, after recording we compare every cell: the tape's value must equal
the engine's value **bit for bit**. That holds on all three test workbooks (0 mismatches).

### IRR: the implicit function theorem

IRR has no formula; Excel finds it by trial and error (Newton's method). Differentiating the trial
and error would be wrong and slow. Instead we use what defines the answer: at the IRR `r`,
`NPV(r, cash flows) = 0`. If one cash flow moves, `r` must move just enough to keep that NPV at
zero, which gives the slope directly:

  ∂r/∂(cash flow i) = −(∂NPV/∂cash flow i) ÷ (∂NPV/∂r) = −(1+r)^(−i) ÷ NPV′(r).

### Circular references: the same idea at the loop's fixed point

An LBO loop (interest depends on debt, debt depends on cash left after interest) settles on values
`x` where recalculating the loop changes nothing: `x = F(x, inputs)`. When an input moves, the
settled values move too, and the implicit function theorem says by how much:

  dx/d(input) = (I − J)⁻¹ · ∂F/∂(input),   where J says how each loop cell's formula reacts to the other loop cells.

On the tape, each loop becomes one block: we record the loop formulas once at the settled values,
measure `J` (one small backward pass per loop cell; the LBO's loops have 7 cells each), and during
the backward pass solve a 7×7 linear system instead of looping. Before recording, loops are
re-solved to machine precision (scaled change ≤ 1e-15) instead of Excel's 1e-9 stopping rule,
because the formula above is exact only at a true fixed point. Recorded loop residuals: ≤ 2.2e-16.

### Kinks: IF, MIN, MAX (and ABS, ROUND, XNPV dates)

`IF(cash < minimum, draw on revolver, 0)` makes the model piecewise: on one side of the line the
slope is one thing, on the other side something else, and exactly on the line there is no slope.
We report the derivative of the branch Excel actually took, and we list every such decision that
the inputs can move, with how close it is to flipping:

- **margin**: a number that crosses zero exactly where the branch flips (for `A > B` it is `A − B`;
  for `MIN` it is the gap between the smallest and the runner-up).
- **uniform move %**: to first order, if every input moved by this percentage in the worst
  direction, the branch would flip. We flag a kink as **near** at ≤ 5%.
- **nearest input**: the single input that reaches the kink with the smallest move, and the move.

Kinks inside a loop are measured including the loop's own response (otherwise their distance would
be wrong). ROUND and XNPV's date truncation are staircases: their slope is 0, and integer dates sit
exactly on a step, so they always show as "on the kink".

## Checks

### Every gradient vs finite differences

Not just the price: for **every numeric formula cell** in each test workbook and **every input cell**,
we compare the tape's derivative with a finite difference (nudge the input up and down, recalculate,
measure the slope). The finite differences use central differences at steps `h` and `h/2`
combined by Richardson extrapolation, `h = 1e-4 × max(1, |input|)`, with loops re-solved at every
nudged point. Pass rule: `|AD − FD| ≤ 1e-6 × max(|AD|, |FD|)`.

Each pair lands in exactly one bucket:

| Bucket | Meaning |
|---|---|
| both exactly 0 | The output doesn't depend on that input; both methods say exactly 0. |
| within 1e-6 relative | The real test. |
| within FD rounding noise | The finite difference can't resolve slopes smaller than its own rounding error (≈ 10·ε·M/h, M = the largest value the output is computed from). In every such pair the tape's derivative is ≤ 1.2e-16, i.e. zero; the typical case is the balance-sheet check, which is 0 by construction. |
| excluded: step crosses a kink | A branch flips inside ±h, so the finite difference measures a jump, not a slope. Listed in the results, never silently dropped. |
| failed | None. |

Results (machine M1, `results/stage2_gradients.json`):

| Workbook | outputs × inputs | pairs | exactly 0 | within 1e-6 | FD noise | kink-excluded | failed | worst passing |
|---|---|---|---|---|---|---|---|---|
| three_statement.xlsx | 229 × 66 | 15,114 | 12,744 | 2,199 | 171 | 0 | **0** | 6.7e-9 |
| lbo_circular.xlsx | 89 × 16 | 1,424 | 841 | 498 | 84 | 1 | **0** | 2.5e-9 |
| functions_coverage.xlsx | 47 × 18 | 846 | 748 | 88 | 0 | 10 | **0** | 6.7e-11 |
| synthetic DCF, 10 segments | 233 × 215 | 50,095 | 45,875 | 4,220 | 0 | 0 | **0** | 2.0e-10 |
| synthetic DCF, 100 segments (price only) | 1 × 2,105 | 2,105 | 0 | 2,105 | 0 | 0 | **0** | 1.7e-9 |

The 11 kink-excluded pairs: on the LBO, `Model!B4` is the "Year 0" counter; the cash-flow row uses
`IF(year = 5, exit value, 0)`, so nudging it by any amount switches the exit year off and IRR fails.
On the coverage workbook, the 10 are XNPV dates (integers, so every nudge crosses a whole-day step).

### Benchmark: one backward pass vs one recalculation per input

| Workbook | inputs | one recalc | FD gradient (inputs + 1 recalcs) | tape: record + backward | speed-up |
|---|---|---|---|---|---|
| functions_coverage | 18 | 0.090 ms | 1.70 ms | 0.187 ms | 9× |
| lbo_circular | 16 | 0.100 ms | 1.69 ms | 0.312 ms | 5× |
| three_statement | 66 | 0.222 ms | 14.9 ms | 0.671 ms | 22× |
| synthetic, 10 segments | 215 | 0.136 ms | 29.4 ms | 0.374 ms | 79× |
| synthetic, 100 segments | 2,105 | 1.28 ms | 2,696 ms | 3.54 ms | 761× |

Medians on M1 (cloud container, Intel Xeon 2.10GHz, 4 cores, Node 22). Timings at this scale vary
by roughly ±30% between runs; the trend is the point: the tape costs about 3 recalculations, finite
differences cost one per input. The backward pass alone is 0.003 ms on the three-statement model;
most of the tape's cost is recording it. The synthetic models are generated in memory
(`scripts/synthetic.ts`) and have no Excel ground truth for their forward values; they only show
how the cost scales.

### Unit tests (`tests/grad.test.ts`)

Hand-derived derivatives (calculus, not our engine) for arithmetic, `^` in both arguments, SUM,
AVERAGE, SUMPRODUCT, NPV (rate and flows), IRR (implicit function theorem), IF branches, MIN/MAX
ties, ABS, ROUND; one- and two-cell loops with exact answers (e.g. `x = p + 0.5x` gives `dx/dp = 2`);
a kink inside a loop; and the share price / IRR of the real test workbooks vs finite differences.

## What the kink report finds on the test workbooks

- **LBO**: the "Meets target?" cell (`IRR ≥ 20%`) is 0.08% (uniform input move) from flipping: the
  IRR is 19.89% vs a 20% target. The `IF(year = 5, …)` cells sit exactly on their kink (equality
  tests on counters always do).
- **Three-statement**: nothing near the share price. Elsewhere, "Lowest EBITDA" (MIN) and "Highest
  EBITDA" (MAX) are 2.5% and 2.0% from switching years, and the terminal-value sanity flag is 4.1%
  away.

## Limits (reported, not hidden)

- Derivatives with respect to constants typed inside a formula or a named formula (e.g. a name
  defined as `=0.09`) aren't tracked; only input *cells* are.
- `x^y` with `x ≤ 0` and an input-driven `y` has no slope in `y`; we set it to 0 and add a warning.
- Array formulas inside a circular block aren't supported (none in the test workbooks); they throw.
- The CFI DCF workbook (still waiting for the file) and the SEC workbooks (Stage 4) get the same check
  when they arrive.
