// Stage 3 check: solve backwards for the assumptions a target price implies, and prove every answer.
// Writes results/stage3_backsolve.json.
//
//   npx tsx scripts/stage3.ts
//
// What is checked:
//  1. One free input vs Goal Seek. Four cases (workbooks/goalseek/cases.json). Our answer must match
//     Excel's Goal Seek to 1e-6 (|ours − Excel| ≤ 1e-6 × max(1, |Excel|)). Excel's answers come from
//     files Alek saved after running Goal Seek on his Mac (workbooks/goalseek/excel/); for those we
//     also re-check every formula cell against Excel's cached values at Excel's answer. LibreOffice's
//     Goal Seek (scripts/lo_goalseek.py) is shown as a second opinion, not as ground truth.
//  2. Several free inputs: the closest point to the base case (weighted least squares) that hits
//     the target exactly. Checked: the target is hit, the tangency condition holds, the curvature
//     along the price surface is positive (a minimum, not a saddle), and brute force finds nothing
//     closer (the traced price curve for 2 drivers; random rays from the base case for 3+).
//  3. The full growth/margin curve at the target, with a grid scan proving no branch was missed.
//  4. How the implied growth moves when WACC changes by ±1 percentage point, with the slope from
//     the implicit function theorem checked against finite differences of re-solved answers.
//  5. Scale: a synthetic model with 20 and 200 free drivers.

import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import os from 'node:os';
import { loadWorkbook, compareToCached, Workbook } from '../src/index';
import { Backsolver, DriverSpec, SolveResult, CurveResult, goalSeek, rootBetween, cellAt } from '../src/engine/solve';
import { syntheticDcf } from './synthetic';

const root = new URL('..', import.meta.url).pathname;
const GS_TOL = 1e-6;
const IFT_TOL = 1e-6;
const machine = { os: `${os.type()} ${os.release()}`, cpu: os.cpus()[0]?.model, cores: os.cpus().length, node: process.version };
const load = (rel: string) => { const { wb, report } = loadWorkbook(new Uint8Array(readFileSync(join(root, rel)))); wb.recalc({ cycleStart: 'zero' }); return { wb, report }; };
const pct = (x: number, d = 2) => `${(100 * x).toFixed(d)}%`;
const pp = (x: number, d = 2) => `${x >= 0 ? '+' : ''}${(100 * x).toFixed(d)} pp`;
const failures: string[] = [];
const check = (ok: boolean, what: string) => { if (!ok) failures.push(what); return ok; };

/** Deterministic random numbers (mulberry32), so the brute-force checks are reproducible. */
function rng(seed: number) {
  return () => { seed |= 0; seed = (seed + 0x6d2b79f5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}
const gauss = (r: () => number) => Math.sqrt(-2 * Math.log(1 - r())) * Math.cos(2 * Math.PI * r());

// ------------------------------------------------------------------ 1. Goal Seek cases

interface Case { id: string; file: string; source: string; setCell: string; toValue: number; changingCell: string; what: string }
const cases: Case[] = JSON.parse(readFileSync(join(root, 'workbooks/goalseek/cases.json'), 'utf8')).cases;
const loPath = join(root, 'results/stage3_goalseek_libreoffice.json');
const lo = existsSync(loPath) ? JSON.parse(readFileSync(loPath, 'utf8')) : null;

/** Is there a root closer to the base case than ours? Scan output − target on a fine grid of [−|d|, +|d|]. */
function closerRoots(wb: Workbook, c: Case, d: number, n = 400) {
  const s = new Backsolver(wb, c.setCell, [{ name: c.changingCell, cells: c.changingCell }]);
  const f = (x: number) => s.value([x]) - c.toValue;
  const R = Math.abs(d);
  const xs = Array.from({ length: 2 * n + 1 }, (_, k) => -R + (R * k) / n);
  const fs = xs.map(f);
  s.reset();
  const found: number[] = [];
  for (let k = 0; k + 1 < xs.length; k++) {
    if (!Number.isFinite(fs[k]) || !Number.isFinite(fs[k + 1])) continue;
    const near = Math.min(Math.abs(xs[k]), Math.abs(xs[k + 1]));
    if (fs[k] * fs[k + 1] < 0 && near < R * (1 - 1e-9) && !(Math.abs(xs[k] - d) < (2 * R) / n || Math.abs(xs[k + 1] - d) < (2 * R) / n)) found.push((xs[k] + xs[k + 1]) / 2);
  }
  return { scanned: `${2 * n + 1} points on [${-R}, ${R}] around the base value`, closerRoots: found };
}

console.log('1. One free input: our answer vs Goal Seek');
const goalSeekRows = cases.map((c) => {
  const { wb } = load(`workbooks/goalseek/${c.file}`);
  const t = performance.now();
  const { answer, result } = goalSeek(wb, c.setCell, c.toValue, c.changingCell);
  const ms = performance.now() - t;
  const base = (cellAt(wb, c.changingCell).value as number);
  const scan = closerRoots(wb, c, answer - base);
  check(result.converged && Math.abs(result.missBy) <= 1e-12 * Math.max(1, Math.abs(c.toValue)), `${c.id}: did not hit the target`);
  check(scan.closerRoots.length === 0, `${c.id}: a root closer to the base case exists`);
  const row: Record<string, unknown> = {
    id: c.id, what: c.what, setCell: c.setCell, toValue: c.toValue, changingCell: c.changingCell, baseValue: base,
    ours: { answer, outputAtAnswer: result.value, missBy: result.missBy, iterations: result.iterations, evaluations: result.evaluations, ms, warnings: result.warnings },
    closestRootScan: scan,
  };
  const l = lo?.cases.find((x: { id: string }) => x.id === c.id);
  if (l) {
    const diff = answer - l.libreofficeAnswer;
    row.libreoffice = { engine: lo.engine, answer: l.libreofficeAnswer, outputAfterwardsInLibreOffice: l.setCellAfterwards, missBy: l.missBy, diffOursMinusLibreOffice: diff, withinTol: Math.abs(diff) <= GS_TOL * Math.max(1, Math.abs(l.libreofficeAnswer)) };
  }
  const xl = join(root, 'workbooks/goalseek/excel', c.file.replace(/\.xlsx$/, '_excel.xlsx'));
  if (existsSync(xl)) {
    const { wb: ewb, report } = load(`workbooks/goalseek/excel/${c.file.replace(/\.xlsx$/, '_excel.xlsx')}`);
    const excelAnswer = cellAt(ewb, c.changingCell).value as number;
    const excelOutput = cellAt(ewb, c.setCell).cached as number;
    const fwd = compareToCached(ewb, 1e-9);
    const diff = answer - excelAnswer;
    const byExcel = !!report.application?.includes('Excel');
    const ok = byExcel && Math.abs(diff) <= GS_TOL * Math.max(1, Math.abs(excelAnswer)) && fwd.mismatches.length === 0 && fwd.noCache === 0;
    check(ok, `${c.id}: Excel Goal Seek comparison failed`);
    row.excel = { savedBy: report.application, answer: excelAnswer, excelsOutputAtItsAnswer: excelOutput, excelMissBy: excelOutput - c.toValue,
      diffOursMinusExcel: diff, tolerance: `|ours − Excel| ≤ ${GS_TOL} × max(1, |Excel|)`, pass: ok,
      forwardCheckAtExcelsAnswer: { compared: fwd.compared, matched: fwd.matched, mismatched: fwd.mismatches.length, noCache: fwd.noCache, maxScaledDiff: fwd.maxScaledDiff, worstCell: fwd.worstCell } };
  } else row.excel = { status: 'waiting for the Excel-saved file from Alek' };
  const e = row.excel as Record<string, unknown>, lr = row.libreoffice as Record<string, number> | undefined;
  console.log(`  ${c.id} ${c.changingCell} for ${c.setCell} = ${c.toValue}: ours ${answer} (miss ${result.missBy.toExponential(1)}, ${result.iterations} Newton steps, ${ms.toFixed(1)} ms)`);
  if (lr) console.log(`      LibreOffice ${lr.answer}  diff ${lr.diffOursMinusLibreOffice.toExponential(2)}  (LibreOffice's own output there misses by ${lr.missBy.toExponential(1)})`);
  console.log(e.status ? `      Excel: ${e.status}` : `      Excel ${e.answer}  diff ${(e.diffOursMinusExcel as number).toExponential(2)}  ${e.pass ? 'PASS' : 'FAIL'} (Excel's own output misses by ${(e.excelMissBy as number).toExponential(1)}; forward check at Excel's answer ${(e.forwardCheckAtExcelsAnswer as { matched: number }).matched}/${(e.forwardCheckAtExcelsAnswer as { compared: number }).compared})`);
  return row;
});

// ------------------------------------------------------------------ 2. Closest point, several drivers

const G = (cells: string): DriverSpec => ({ name: 'Revenue growth (every year)', cells });
const TS_TARGET = 30;
const growth = G('Assumptions!C5:G5');
const gm: DriverSpec = { name: 'Gross margin (every year)', cells: 'Assumptions!C6:G6' };
const sga: DriverSpec = { name: 'SG&A % of revenue (every year)', cells: 'Assumptions!C7:G7' };
const capex: DriverSpec = { name: 'Capex % of revenue (every year)', cells: 'Assumptions!C9:G9' };

/** A driver's change in its own units: percentage points for rates, plain numbers otherwise. */
const change = (d: SolveResult['drivers'][number]) => d.cells.every((c) => Math.abs(c.base) <= 1) && d.mode === 'shift'
  ? `${d.change >= 0 ? '+' : ''}${(100 * d.change).toPrecision(3)} pp` : `${d.change >= 0 ? '+' : ''}${d.change.toPrecision(3)}`;
function describe(r: SolveResult) {
  return r.drivers.map((d) => `${d.name} ${change(d)} (${d.normalMoves.toFixed(3)} normal moves, ${(100 * d.shareOfDistance).toFixed(1)}% of the distance)`).join('; ');
}

/** Brute force for 3+ drivers: shoot rays from the base case, find where each first reaches the target surface. Nothing may be closer than the solver's answer. */
function rays(s: Backsolver, target: number, r: SolveResult, nRandom: number, nLocal: number) {
  const m = s.size, sig = s.drivers.map((d) => d.sigma);
  const rand = rng(20261106);
  const f = (u: number[]) => s.value(u.map((x, i) => x * sig[i])) - target;
  const dirs: number[][] = [];
  for (let k = 0; k < nRandom; k++) { const v = Array.from({ length: m }, () => gauss(rand)); const n = Math.hypot(...v); dirs.push(v.map((x) => x / n)); }
  // Directions close to the answer's (within about 10°): these test the neighbourhood hardest.
  const u0 = r.u.map((x) => x / r.distance);
  for (let k = 0; k < nLocal; k++) { const v = u0.map((x) => x + 0.18 * gauss(rand) / Math.sqrt(m)); const n = Math.hypot(...v); dirs.push(v.map((x) => x / n)); }
  const f0 = f(new Array(m).fill(0));
  const steps = 12, smax = 3 * r.distance;
  let best = Infinity, reached = 0, evals = 0;
  for (const v of dirs) {
    let a = 0, fa = f0;
    for (let k = 1; k <= steps; k++) {
      const b = (smax * k) / steps, fb = f(v.map((x) => x * b)); evals++;
      if (Number.isFinite(fa) && Number.isFinite(fb) && fa * fb <= 0) {
        const sStar = fb === 0 ? b : rootBetween((t) => { evals++; return f(v.map((x) => x * t)); }, a, b, fa, fb, 1e-12 * Math.max(1, Math.abs(target)));
        best = Math.min(best, sStar); reached++; break;
      }
      a = b; fa = fb;
    }
  }
  s.reset();
  return { rays: dirs.length, randomRays: nRandom, raysNearTheAnswer: nLocal, reachedTargetWithin3xDistance: reached, closestFound: best, solverDistance: r.distance, ratio: best / r.distance, pass: best >= r.distance * (1 - 1e-9), evaluations: evals };
}

function closest(label: string, wb: Workbook, output: string, specs: DriverSpec[], target: number, brute: 'curve' | 'rays', box?: [[number, number], [number, number]]) {
  const s = new Backsolver(wb, output, specs);
  const r = s.solve(target);
  const tolF = 1e-12 * Math.max(1, Math.abs(target));
  let bruteForce: Record<string, unknown>;
  if (brute === 'curve') {
    // Trace the price curve WITHOUT starting from the answer, then look for its closest point.
    const c = s.traceCurve(target, 0, 1, box!);
    let best = Infinity;
    for (const b of c.branches) for (const p of b.points) best = Math.min(best, Math.hypot(p.d[0] / s.drivers[0].sigma, p.d[1] / s.drivers[1].sigma));
    // Pass: no traced point is closer than the answer. (The nearest traced point is a little farther
    // because the points are spaced out along the curve; it must still come within 1%.)
    bruteForce = { method: 'closest point among the traced curve points (curve traced independently of the answer)', curvePoints: c.points, closestFound: best, solverDistance: r.distance, ratio: best / r.distance, pass: best >= r.distance * (1 - 1e-9) && best <= r.distance * 1.01 };
  } else bruteForce = { method: 'rays from the base case to the target surface', ...rays(s, target, r, specs.length >= 4 ? 300 : 150, 60) };
  const ok = r.converged && Math.abs(r.missBy) <= tolF && r.tangency <= 1e-9 && (r.reducedHessianMinEig ?? 0) > 0 && !!bruteForce.pass;
  check(ok, `closest point "${label}" failed a check`);
  console.log(`  ${label}: ${describe(r)}`);
  console.log(`      hits ${r.value} (miss ${r.missBy.toExponential(1)}), distance ${r.distance.toFixed(6)} normal moves, tangency ${r.tangency.toExponential(1)}, curvature along the surface ≥ ${r.reducedHessianMinEig?.toFixed(3)}, ${r.iterations} Newton steps, ${r.evaluations.gradients} gradient evaluations, ${r.ms.toFixed(0)} ms`);
  console.log(`      brute force (${bruteForce.method}): closest found ${(bruteForce.closestFound as number).toFixed(6)} vs solver ${r.distance.toFixed(6)} → ${bruteForce.pass ? 'nothing closer' : 'FAIL'}`);
  if (r.nearKinks.length) console.log(`      near kinks: ${r.nearKinks.map((k) => `${k.cell} ${k.kind} (${k.normalMovesToFlip.toFixed(2)} normal moves)`).join(', ')}`);
  return { label, output, target, result: r, bruteForce, pass: ok };
}

console.log('\n2. Several free inputs: the closest assumptions that hit the target exactly');
const ts = load('workbooks/three_statement.xlsx').wb;
const tsBox: [[number, number], [number, number]] = [[-0.10, 0.20], [-0.15, 0.20]];
const closestRows = [
  closest('three-statement, $30, growth + gross margin (σ = 1 pp each)', ts, 'Valuation!B13', [growth, gm], TS_TARGET, 'curve', tsBox),
  closest('three-statement, $30, growth + gross margin, margins stickier (σ_margin = 0.5 pp)', ts, 'Valuation!B13', [growth, { ...gm, sigma: 0.005 }], TS_TARGET, 'curve', [[-0.10, 0.20], [-0.30, 0.40]]),
  closest('three-statement, $30, growth + gross margin + SG&A + capex (σ = 1 pp each)', ts, 'Valuation!B13', [growth, gm, sga, capex], TS_TARGET, 'rays'),
];
const lbo = load('workbooks/lbo_circular.xlsx').wb;
const lboDrivers: DriverSpec[] = [
  { name: 'Revenue growth', cells: 'Inputs!B8' },
  { name: 'Margin improvement per year', cells: 'Inputs!B10', sigma: 0.0025 },
  { name: 'Exit multiple', cells: 'Inputs!B13' },
];
closestRows.push(closest('LBO, 20% IRR, growth + margin improvement (σ 0.25 pp) + exit multiple (σ 1.0x)', lbo, 'Returns!B8', lboDrivers, 0.2, 'rays'));

// ------------------------------------------------------------------ 3. The growth/margin curve

console.log('\n3. Every growth/margin pair that justifies the target');
function curve(label: string, s: Backsolver, target: number, box: [[number, number], [number, number]]) {
  const c: CurveResult = s.traceCurve(target, 0, 1, box);
  const ok = c.grid.covered === c.grid.crossings && c.maxAbsMissBy <= 1e-12 * Math.max(1, Math.abs(target)) && c.branches.length > 0;
  check(ok, `curve "${label}" failed a check`);
  console.log(`  ${label}: ${c.branches.length} branch(es), ${c.points} points, every point hits the target to ${c.maxAbsMissBy.toExponential(1)}; grid ${c.grid.n}×${c.grid.n}: ${c.grid.covered}/${c.grid.crossings} crossings on the traced curve (worst gap ${c.grid.worstGap.toExponential(1)}), ${c.ms.toFixed(0)} ms`);
  for (const b of c.branches) {
    const a = b.points[0], z = b.points[b.points.length - 1];
    console.log(`      from ${pct(a.values[0])} growth / ${pct(a.values[1])} margin (${b.ends[0]}) to ${pct(z.values[0])} growth / ${pct(z.values[1])} margin (${b.ends[1]})`);
  }
  return { label, pass: ok, curve: c };
}
const tsCurveSolver = new Backsolver(ts, 'Valuation!B13', [growth, gm]);
const curves = [curve('three-statement, value per share $30: average growth vs average gross margin', tsCurveSolver, TS_TARGET, tsBox)];
const lboCurveSolver = new Backsolver(lbo, 'Returns!B8', [lboDrivers[0], lboDrivers[1]]);
curves.push(curve('LBO, IRR 20%: revenue growth vs margin improvement per year', lboCurveSolver, 0.2, [[-0.10, 0.10], [-0.01, 0.02]]));

// ------------------------------------------------------------------ 4. WACC ±1 pp

console.log('\n4. Implied growth as WACC moves ±1 percentage point (value per share $30)');
const waccCell = cellAt(ts, 'Valuation!B3');
const wacc0 = waccCell.value as number;
const atWacc = (w: number, fn: () => unknown) => { waccCell.value = w; ts.recalc({ tight: true, cycleStart: 'current' }); try { return fn(); } finally { waccCell.value = wacc0; ts.recalc({ tight: true, cycleStart: 'current' }); } };
const sweep = [-0.01, -0.0075, -0.005, -0.0025, 0, 0.0025, 0.005, 0.0075, 0.01].map((dw) => atWacc(wacc0 + dw, () => {
  const s = new Backsolver(ts, 'Valuation!B13', [growth]);
  const r = s.solve(TS_TARGET);
  const sens = s.sensitivity1D(r, 'Valuation!B3');
  const cagr = s.cellsAt(r.d, ['Valuation!B20'])['Valuation!B20'] as number;
  const two = new Backsolver(ts, 'Valuation!B13', [growth, gm]).solve(TS_TARGET);
  check(r.converged && two.converged, `WACC ${wacc0 + dw}: did not converge`);
  return { wacc: wacc0 + dw, impliedGrowthShift: r.d[0], impliedAverageGrowth: s.meanImplied(0, r.d[0]), impliedRevenueCagr: cagr, missBy: r.missBy,
    dShiftDWaccIft: sens.dAnswer, closestPointGrowthPlusMargin: { growthShift: two.d[0], marginShift: two.d[1], distance: two.distance } };
})) as { wacc: number; impliedGrowthShift: number; impliedAverageGrowth: number; impliedRevenueCagr: number; missBy: number; dShiftDWaccIft: number; closestPointGrowthPlusMargin: { growthShift: number; marginShift: number; distance: number } }[];
for (const r of sweep) console.log(`  WACC ${pct(r.wacc)}: growth every year ${pp(r.impliedGrowthShift)} → average ${pct(r.impliedAverageGrowth)}, revenue CAGR ${pct(r.impliedRevenueCagr)}; slope ${(r.dShiftDWaccIft).toFixed(4)} pp of growth per pp of WACC`);
// IFT slope vs finite differences of re-solved answers (Richardson, like Stage 2).
const iftChecks = [wacc0 - 0.01, wacc0, wacc0 + 0.01].map((w) => {
  const solveAt = (x: number) => atWacc(x, () => new Backsolver(ts, 'Valuation!B13', [growth]).solve(TS_TARGET).d[0]) as number;
  const ift = atWacc(w, () => { const s = new Backsolver(ts, 'Valuation!B13', [growth]); return s.sensitivity1D(s.solve(TS_TARGET), 'Valuation!B3').dAnswer; }) as number;
  const h = 1e-4 * Math.max(1, Math.abs(w));
  const D1 = (solveAt(w + h) - solveAt(w - h)) / (2 * h), D2 = (solveAt(w + h / 2) - solveAt(w - h / 2)) / h;
  const fd = (4 * D2 - D1) / 3;
  const rel = Math.abs(ift - fd) / Math.max(Math.abs(ift), Math.abs(fd));
  check(rel <= IFT_TOL, `WACC ${w}: implicit-function slope vs finite differences ${rel}`);
  return { wacc: w, ift, finiteDifference: fd, relError: rel, pass: rel <= IFT_TOL };
});
for (const c of iftChecks) console.log(`  slope check at WACC ${pct(c.wacc)}: implicit function theorem ${c.ift.toFixed(9)}, finite differences ${c.finiteDifference.toFixed(9)}, rel. error ${c.relError.toExponential(1)} ${c.pass ? 'PASS' : 'FAIL'}`);
const lo1 = sweep[0], mid = sweep[4], hi1 = sweep[8];
const waccSummary = {
  targetValuePerShare: TS_TARGET, baseWacc: wacc0,
  impliedAverageGrowth: { waccMinus1pp: lo1.impliedAverageGrowth, base: mid.impliedAverageGrowth, waccPlus1pp: hi1.impliedAverageGrowth },
  changeInImpliedGrowth: { waccMinus1pp: lo1.impliedGrowthShift - mid.impliedGrowthShift, waccPlus1pp: hi1.impliedGrowthShift - mid.impliedGrowthShift },
  baseCaseAverageGrowth: new Backsolver(ts, 'Valuation!B13', [growth]).meanImplied(0, 0),
};
console.log(`  → at $${TS_TARGET}: WACC −1 pp moves the implied growth by ${pp(waccSummary.changeInImpliedGrowth.waccMinus1pp)}, WACC +1 pp by ${pp(waccSummary.changeInImpliedGrowth.waccPlus1pp)} (base WACC needs ${pct(mid.impliedAverageGrowth)} average growth vs ${pct(waccSummary.baseCaseAverageGrowth)} in the base case)`);
const waccCurves = [-0.01, 0.01].map((dw) => atWacc(wacc0 + dw, () => curve(`three-statement, $30, WACC ${pct(wacc0 + dw)}: growth vs gross margin`, new Backsolver(ts, 'Valuation!B13', [growth, gm]), TS_TARGET, tsBox)));

// ------------------------------------------------------------------ 5. Scale

console.log('\n5. Scale: synthetic DCF (in memory, no Excel file)');
const scale = [false, true].map((perCell) => {
  const { wb, output } = syntheticDcf(10);
  const P0 = wb.getCell('Valuation', 'A3')!.value as number;
  const specs: DriverSpec[] = [];
  for (let s = 1; s <= 10; s++) {
    if (perCell) for (let y = 1; y <= 10; y++) { specs.push({ name: `growth s${s} y${y}`, cells: `Inputs!${String.fromCharCode(65 + y)}${s}` }); specs.push({ name: `margin s${s} y${y}`, cells: `Inputs!${colName(11 + y)}${s}` }); }
    else { specs.push({ name: `growth, segment ${s}`, cells: `Inputs!B${s}:K${s}` }); specs.push({ name: `margin, segment ${s}`, cells: `Inputs!L${s}:U${s}` }); }
  }
  const s = new Backsolver(wb, output, specs);
  const r = s.solve(P0 * 1.15);
  const ok = r.converged && Math.abs(r.missBy) <= 1e-12 * Math.abs(P0 * 1.15) && r.tangency <= 1e-9;
  check(ok, `scale test with ${specs.length} drivers failed`);
  console.log(`  ${specs.length} drivers (${perCell ? 'every growth and margin cell on its own' : 'one growth and one margin driver per segment'}): target +15%, ${r.converged ? 'converged' : 'NOT converged'}, miss ${r.missBy.toExponential(1)}, tangency ${r.tangency.toExponential(1)}, ${r.iterations} steps, ${r.evaluations.gradients} gradient evaluations, ${r.ms.toFixed(0)} ms (${r.method})`);
  return { drivers: specs.length, target: P0 * 1.15, converged: r.converged, missBy: r.missBy, tangency: r.tangency, iterations: r.iterations, evaluations: r.evaluations, ms: r.ms, method: r.method, pass: ok };
});
function colName(n: number) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// ------------------------------------------------------------------ write

const excelDone = goalSeekRows.every((r) => !(r.excel as { status?: string }).status);
const overall = failures.length ? 'FAIL' : excelDone ? 'PASS' : 'PASS (Excel Goal Seek comparison pending)';
const slim = (c: { curve: CurveResult }) => ({ ...c.curve, branches: c.curve.branches.map((b) => ({ ...b, points: b.points.map((p) => ({ d: p.d, values: p.values, missBy: p.missBy })) })) });
mkdirSync(join(root, 'results'), { recursive: true });
writeFileSync(join(root, 'results', 'stage3_backsolve.json'), JSON.stringify({
  generatedAt: new Date().toISOString(), machine, overall, failures,
  tolerances: { goalSeek: `|ours − Goal Seek| ≤ ${GS_TOL} × max(1, |Goal Seek|)`, target: '|output − target| ≤ 1e-12 × max(1, |target|)', tangency: 'sin(angle between the move and the price gradient) ≤ 1e-9', impliedSlope: `|IFT − FD| ≤ ${IFT_TOL} × max(|IFT|, |FD|)` },
  goalSeek: goalSeekRows,
  closestPoint: closestRows.map((c) => ({ ...c, result: { ...c.result } })),
  curves: curves.map((c) => ({ label: c.label, pass: c.pass, ...slim(c) })),
  waccPlusMinus1pp: { summary: waccSummary, sweep, slopeChecks: iftChecks, curves: waccCurves.map((c) => ({ label: (c as { label: string }).label, pass: (c as { pass: boolean }).pass, ...slim(c as { curve: CurveResult }) })) },
  scale,
}, null, 1) + '\n');
console.log(`\nwrote results/stage3_backsolve.json — overall ${overall}`);
for (const f of failures) console.log(`  FAILED: ${f}`);
process.exit(failures.length ? 1 : 0);
