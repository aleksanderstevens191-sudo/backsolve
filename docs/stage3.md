# Stage 3: from a price back to the assumptions

## What it does

Stage 1 runs the workbook forwards (assumptions in, price out). Stage 2 measures how the price
reacts to every assumption. Stage 3 runs it backwards: **you give a price, it gives the assumptions
that produce exactly that price.**

- **One free assumption** (like Excel's Goal Seek): there is usually one answer, and we find it.
- **Several free assumptions**: infinitely many combinations give the same price. We return the
  one that is **closest to the base case** under a weighting you can read and change, and, for two
  assumptions, the **whole curve** of combinations that give the price.

```ts
import { loadWorkbook, Backsolver, goalSeek } from './src/index';
const { wb } = loadWorkbook(bytes);
wb.recalc();

// One free input, like Goal Seek: the WACC that gives $30 per share
goalSeek(wb, 'Valuation!B13', 30, 'Valuation!B3').answer;          // 0.08261637706510873

// Several: the closest growth + margin to the base case that gives exactly $30
const s = new Backsolver(wb, 'Valuation!B13', [
  { name: 'Revenue growth', cells: 'Assumptions!C5:G5' },              // σ defaults to 1 pp
  { name: 'Gross margin',   cells: 'Assumptions!C6:G6', sigma: 0.005 }, // "margins are stickier"
]);
const r = s.solve(30);    // r.drivers[k].change, r.distance, r.tangency, r.nearKinks, ...
const c = s.traceCurve(30, 0, 1, [[-0.1, 0.2], [-0.15, 0.2]]);         // every pair that gives $30
```

A **driver** is a group of input cells that move together. By default a driver *shifts* every cell
by the same amount (growth 8%, 7%, 6%, 5%, 4% → each plus 1.36 pp); `mode: 'scale'` multiplies
them instead. The workbook is put back exactly as it was after every call.

Run every check with `npm run stage3` (writes `results/stage3_backsolve.json`). All numbers below
come from that file, machine M1 (see `LAUNCH_FACTS.md`).

## The weighting, in plain English

Take the three-statement test model. Its base case gives **$26.10** per share. Which growth and
gross margin give **$30**? A lot more growth and the same margin works. So does the same growth and
a lot more margin, and every mix in between (that's the curve in the chart below). We need a rule
for picking one.

The rule: **change the base case as little as possible**, where "little" is measured like this:

    distance² = (growth change ÷ σ_growth)² + (margin change ÷ σ_margin)² + ...

**σ ("sigma") is what a normal-sized move of that assumption looks like.** Dividing by σ turns
every change into "how many normal-sized moves", so a change in growth and a change in an exit
multiple are in the same units and can be added up. Squaring means two small changes count as less
than one big one: the answer spreads the move across the assumptions instead of piling it on one.

Default σ (each is overridable per driver, and the result says which rule it used):

| Driver | Default σ | Why |
|---|---|---|
| shift driver, every cell between −1 and 1 (growth, margins, rates) | 1 percentage point | These look like rates; 1 pp is a normal-sized disagreement. |
| other shift driver (exit multiple, a revenue figure) | 10% of the average base value | e.g. exit multiple 10.0x → σ = 1.0x |
| scale driver | 10% | |

**Who moves more.** At the answer, each driver's change is proportional to
**σ² × (how sensitive the price is to it)**. So an assumption moves more when the price reacts to it
more strongly, and more when you say you're less sure of it (bigger σ). In the $30 example the
price reacts to gross margin by $1.47 per pp and to growth by $1.18 per pp (at the answer), so with
equal σ the margin moves 1.24 times as much as growth, exactly the ratio of the two sensitivities.

**The statistics behind it.** If you think each assumption's error follows a bell curve of width σ
around the base case, then among all the combinations that hit the price exactly, this is the most
likely one. That's the whole justification; the answer is only as good as the σ's.

**So the answer is "the closest one under these weights", not "the" implied assumption.** Change
the weights and the answer changes (see the table below). That's why we also trace the full curve:
it shows every combination, with no weighting at all.

### Results on the test workbooks

| Case | Answer | Distance (normal moves) |
|---|---|---|
| Three-statement, $30, growth + gross margin, σ = 1 pp each | growth **+1.357 pp** (average 6.00% → 7.36%), gross margin **+1.684 pp** (41.60% → 43.28%) | 2.163 |
| Same, but margins stickier (σ_margin = 0.5 pp) | growth **+2.470 pp**, gross margin **+0.814 pp** | 2.958 |
| Three-statement, $30, four drivers, σ = 1 pp each | growth +0.530, gross margin +0.629, SG&A −0.627, capex −0.836 pp | 1.330 |
| LBO, IRR 19.89% → 20%, growth (σ 1 pp) + margin step (σ 0.25 pp) + exit multiple (σ 1.0x) | growth +0.011 pp, margin step +0.0034 pp, exit multiple +0.020x | 0.0265 |

Every answer hits the target to within 1e-12 (worst miss: 8.1e-13 of IRR on the LBO, 1.4e-14 of a
dollar on the three-statement model). With four drivers the distance drops from 2.16 to 1.33
because the move is shared across more levers; capex moves the most because the price is most
sensitive to it.

A straight-line estimate from the base case's slopes alone would say +1.352 pp growth and +1.808 pp
margin. The real answer needs less margin (+1.684) because margin's effect grows as growth rises
(more revenue to earn the margin on); the solver follows the actual curve, not the straight line.

## How the closest point is found and proven

**The condition.** At the closest point, the direction back to the base case is perpendicular to
the curve (surface, with 3+ drivers) of equal price. If it weren't, you could slide along the curve
towards the base case and get closer. In σ units this says the move `u` points exactly along the
price gradient: `u = λ × gradient`. We call the leftover angle the **tangency** (the sine of the
angle between the move and the gradient; 0 is perfect).

**The solver.** Newton's method on two conditions at once: "move is along the gradient" and
"price = target". Gradients are exact (the Stage 2 tape). Newton also needs the curvature of the
price; we get it by nudging each driver by 1e-4 σ and differencing the exact gradients (up to 40
drivers). Each step is checked with a line search, and if Newton can't make progress from the base
case, the solver walks the target in from the base price in smaller jumps. With more than 40 drivers
we skip the curvature and use plain projection steps (move to the price surface along the shortest
path), which is slower to converge per step but needs one gradient per step.

**What's checked on every answer** (in `npm run stage3`):

1. **Hits the target**: `|output − target| ≤ 1e-12 × max(1, |target|)`, recomputed from scratch.
2. **Tangency ≤ 1e-9.** Worst on the test workbooks: 7.7e-10 (LBO, through 5 loops and an IRR);
   1.0e-16 to 4.8e-16 on the three-statement model.
3. **A minimum, not a saddle.** The tangency condition also holds at the *farthest* point or a
   saddle. We check the curvature of the distance along the price surface; it must be positive.
   Smallest values: 1.044, 1.040, 1.000 (three-statement) and 1.0003 (LBO).
4. **Brute force finds nothing closer.**
   - 2 drivers: the curve is traced *independently* (no hint where the answer is) and we take its
     closest point. It must be no closer than the solver's answer (allowing 1e-9 for rounding) and
     within 1% of it (traced points are spaced along the curve, so the nearest one is a little
     farther). Results: 2.16514 vs 2.16275, and 2.95795 vs 2.95764.
   - 3+ drivers: fire rays from the base case in random directions (plus 60 near the answer) and
     find exactly where each one first reaches the target. Four drivers: 360 rays, 131 reached the
     target within 3× the solver's distance, closest 1.33135 vs 1.32998. LBO: 210 rays, 114 reached,
     closest 0.026526 vs 0.026516. Nothing beat the solver.
5. **Kinks nearby.** Every IF / MIN / MAX the drivers could flip by moving less than one normal
   move from the answer is listed (`r.nearKinks`). None on these cases.

## One free input: Goal Seek

With one free input there's no weighting: we need the input value where the output equals the
target. We use Newton's method with the exact slope, damped until the root is bracketed, then
bisection whenever a Newton step would leave the bracket or stops making fast progress. A flat slope at the start triggers an
outward search. Answers stop at `|output − target| ≤ 1e-12 × max(1, |target|)`.

**How the answer itself moves.** If something else in the model changes (say WACC), the answer
moves by `−(∂price/∂WACC) ÷ (∂price/∂answer)` (the implicit function theorem again, as with IRR in
Stage 2). `s.sensitivity1D(result, 'Valuation!B3')` returns it, and the WACC section below checks it.

### The check against Excel

The spec says Excel's Goal Seek is the reference, to 1e-6. Excel only runs on Alek's Mac, so the
four cases are files (`workbooks/goalseek/`, made by `scripts/make_goalseek.py` from the
Excel-saved test workbooks; `cases.json` lists them). Each opens on the sheet with the changing
cell, has the Goal Seek entries written on that sheet, and turns on iterative calculation with
**Maximum Change 1e-9**: Excel's Goal Seek stops once it's within Maximum Change of the goal, and
the default (0.001) is far too loose for a 1e-6 comparison.

When the `_excel` files come back (saved after Goal Seek), they go in `workbooks/goalseek/excel/`
and `npm run stage3` compares automatically: `|ours − Excel| ≤ 1e-6 × max(1, |Excel|)`, plus every
formula cell recomputed at Excel's answer vs Excel's own cached values (1e-9, the Stage 1 rule).
**Status: waiting for those files.**

Meanwhile, LibreOffice 24.2.7.2's Goal Seek (`scripts/lo_goalseek.py`) is a second opinion, not
ground truth:

| Case | Changing cell (base) | Our answer | LibreOffice | Ours − LibreOffice | Our miss | LibreOffice's miss (in LibreOffice) |
|---|---|---|---|---|---|---|
| gs1: 2026 revenue growth for $30 | `Assumptions!C5` (8.00%) | 0.23147236498254697 | 0.2314723649825472 | −2.2e-16 | −3.6e-15 | 1.8e-14 |
| gs2: terminal growth for $30 | `Valuation!B4` (2.50%) | 0.034115275755361524 | 0.034115275752115364 | 3.2e-12 | 7.0e-12 | −1.6e-9 |
| gs3: WACC for $30 | `Valuation!B3` (9.00%) | 0.08261637706510873 | 0.08261637706526173 | −1.5e-13 | 2.1e-13 | −9.1e-11 |
| gs4: LBO margin step for 20% IRR | `Inputs!B10` (0.50%) | 0.005129341235855554 | 0.005128652659766912 | 6.9e-7 | 2.5e-16 | −6.0e-6 |

All four agree within 1e-6. gs4 is the interesting one: LibreOffice's own IRR at its answer is
19.9994%, not 20%, because it doesn't fully re-solve the LBO's five circular blocks between Goal
Seek steps. Ours re-solves every loop to machine precision, and its IRR at the answer misses 20% by
2.5e-16. Excel's result on gs4 will tell us which way Excel leans.

**Is it the nearest answer?** Goal Seek finds *a* root, usually the one nearest the start. We scan
801 points between the base value and our answer (and as far again on the other side); no case has
a root closer to the base value.

## Every growth/margin pair that gives the price

![Every growth and margin pair that gives $30 per share, at WACC 8%, 9% and 10%](stage3_curve.svg)

`traceCurve` walks along the curve of equal price inside a box of driver changes:

1. **Step forward** along the curve's direction (perpendicular to the gradient), a small distance.
2. **Correct** back onto the curve with Newton, along the shortest path. Every kept point hits the
   price to within 1e-12.
3. **Adapt the step**: shorter where the curve bends (more than 15° per step), longer where it's
   straight. At a sharp corner (an IF or MAX switching), try turned directions up to 90°.
4. Stop exactly on the box edge (found by root-finding), or when the curve closes into a loop.
   Then walk the other way from the start.

**Proof that no piece was missed.** A curve can have several separate pieces. Independently of the
walk, we evaluate the price on a 41 × 41 grid over the box. Wherever two neighbouring grid points
are on opposite sides of the target, the curve must pass between them; we find that crossing
exactly and check that the traced curve passes through it (within 0.1 grid cell). Any crossing not
covered would start a new piece.

| Curve | Box (driver changes) | Points | Pieces | Grid crossings covered | Worst miss of the price |
|---|---|---|---|---|---|
| Three-statement $30, growth × gross margin, WACC 9% | growth −10 to +20 pp, margin −15 to +20 pp | 130 | 1, edge to edge | 66 / 66 | 5.3e-14 |
| Same at WACC 8% | same | 122 | 1 | 62 / 62 | 5.3e-14 |
| Same at WACC 10% | same | 137 | 1 | 70 / 70 | 5.0e-14 |
| LBO 20% IRR, growth × margin step | growth −10 to +10 pp, margin step −1 to +2 pp | 120 | 1, edge to edge | 68 / 68 | 3.5e-13 |

The chart draws both axes at the same scale (1 pp = 16 px), so "closest" looks right: the dashed
line from the base case meets the 9% curve at a right angle. (`npx tsx scripts/stage3_chart.ts`
redraws it from the results file.)

## How much the implied growth moves when WACC moves ±1 pp

The spec flags WACC as where these models are weakest, because nobody agrees on it to within a
percentage point. So: with WACC held at different values, what growth (every year shifted by the
same amount, margins at the base case) gives exactly $30?

| WACC | Implied average growth, 2026–2030 | Change vs WACC 9% | Slope: pp of growth per pp of WACC |
|---|---|---|---|
| 8.00% | 4.67% | **−4.90 pp** | 5.15 |
| 8.25% | 5.94% | −3.63 pp | 5.02 |
| 8.50% | 7.18% | −2.39 pp | 4.90 |
| 8.75% | 8.39% | −1.18 pp | 4.78 |
| **9.00% (base)** | **9.57%** | 0 | 4.68 |
| 9.25% | 10.73% | +1.16 pp | 4.58 |
| 9.50% | 11.86% | +2.29 pp | 4.49 |
| 9.75% | 12.98% | +3.40 pp | 4.40 |
| 10.00% | 14.07% | **+4.49 pp** | 4.32 |

(Base-case average growth: 6.00%.)

**In plain English:** one percentage point of WACC is worth about 4.5 to 5 percentage points of
growth, every year for five years. The base case is only 3.6 pp of growth away from $30 to begin
with, so a 1 pp disagreement about WACC is bigger than the whole gap we're explaining. "The market
is pricing in 9.6% growth" really means "9.6% growth *if* WACC is 9%"; at 8% it's 4.7%, at 10% it's
14.1%. The effect is lopsided (−4.90 vs +4.49) because value is more sensitive to WACC when WACC is
lower.

**Check:** the slope column comes from the implicit function theorem (one gradient, no re-solving).
At 8%, 9% and 10% it's checked against finite differences of fully re-solved answers (Richardson,
steps of 0.01 pp and 0.005 pp): relative errors 1.1e-11, 7.4e-15 and 9.8e-13 (limit 1e-6).

With growth *and* margin free (σ = 1 pp each), the closest point is 0.81 normal moves from the
base case at WACC 8%, 2.16 at 9% and 4.87 at 10%; the three curves are in the chart.

## Scale

A synthetic DCF built in memory (10 segments × 10 years, `scripts/synthetic.ts`; no Excel file, so
it's only a speed and convergence test). Target: base price + 15%.

| Drivers | What moves | Method | Steps | Gradient evaluations | Time | Miss | Tangency |
|---|---|---|---|---|---|---|---|
| 20 | one growth and one margin driver per segment | Newton with curvature | 3 | 166 | 83 ms | 1.8e-15 | 3.3e-16 |
| 200 | every growth and margin cell on its own | projection steps (no curvature matrix) | 7 | 10 | 20 ms | 7.1e-15 | 1.2e-10 |

The 200-driver case is faster because it skips the curvature matrix (2 gradients per driver per
step); one gradient costs about the same however many drivers there are, thanks to Stage 2. Single
run on M1; timings vary by roughly ±30% between runs.

## Unit tests (`tests/solve.test.ts`)

22 tests with answers derived by hand, not by our engine:

- **One input:** Gordon growth (`g = r − D/P`); a pole that throws plain Newton far off
  (`10/(1 − x) = 1000`); through a circular reference (`x = p + 0.5x`); through IRR (annuity
  factor); a flat start; the answer's slopes (`dg/dr = 1`, `dg/dD = −1/P`); the workbook restored.
- **Closest point:** a linear model vs the closed form `d = σ²·g·(T − P0) / Σσ²g²`; σ twice as big
  → 4 times the move; circle (from (3, 4) to radius 10 → (6, 8)); hyperbola (from (1, 1) to
  `a·b = 4` → (2, 2), curvature check sees a minimum); scale drivers; shift groups; refusing
  formula cells, text cells and cells claimed twice.
- **Curves:** `a·b = 4` ending exactly on the box; `a² − b² = 1` (two separate pieces, both found);
  a circle (closed loop); a MAX corner walked through in one piece.
- **Real workbooks:** the three-statement $30 answer (hit, minimum, workbook restored) and the LBO
  margin step for 20% IRR vs plain bisection on the engine (to 1e-12).

`npm test`: 92 tests (49 Stage 1 + 21 Stage 2 + 22 Stage 3), all passing.

## Limits (reported, not hidden)

- **The answer depends on σ.** The defaults are a convention, not a fact about the company. Every
  result lists each σ and the rule that chose it.
- **Closest is local.** Newton finds a nearby minimum; the brute-force checks that prove nothing is
  closer run in `npm run stage3` on the test cases, not on every `solve()` call. A model with several
  separate pieces of the price surface could have a closer point the solver didn't start near.
- **Kinks.** If the path to the answer crosses an IF / MIN / MAX switch, the slope jumps. The
  solver usually still converges (the line search protects it), but the answer can sit right next
  to a switch; `nearKinks` lists those within one normal move.
- **Curves are for two drivers.** With 3+ drivers the set of answers is a surface; we give the
  closest point, not the surface.
- **The 41 × 41 grid** can't see a piece of curve that fits entirely inside one grid cell (a tiny
  closed loop). Make the box smaller or the grid finer if that matters.
- **Excel Goal Seek comparison: pending** Alek's four saved files. Until then, only LibreOffice
  (not ground truth) has been compared.
- **Excel's own Goal Seek is not exact**: it stops within Maximum Change of the goal (1e-9 in our
  case files). The 1e-6 comparison leaves room for that.
