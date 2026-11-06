// Stage 3: solve the workbook backwards, from a target price (or IRR) to the assumptions it implies.
//
// The problem, in plain English
// -----------------------------
// A valuation model turns assumptions (growth, margins, WACC, ...) into a price. Here we know the
// price and want the assumptions.
//
//  * One free assumption: there is (usually) one answer, found by root-finding. This is what
//    Excel's Goal Seek does. We use Newton's method with the exact slope from the Stage 2 tape, and
//    fall back to bisection whenever Newton misbehaves.
//
//  * Several free assumptions: infinitely many combinations hit the price (a little more growth and
//    a little less margin can give the same price as the reverse). We pick ONE: the combination that
//    changes the base case the least, measured by a weighted sum of squares,
//
//        distance² = Σ_k (change_k / σ_k)²,
//
//    subject to the model's output being exactly the target. σ_k ("sigma") is what a normal-sized
//    move of driver k looks like, e.g. 1 percentage point of growth. Dividing by σ puts every
//    assumption in the same currency ("how many normal-sized moves"), so we can add them up.
//    Statistically, this is the most likely set of assumptions if each one's surprise follows a
//    bell curve of width σ around the base case and the price has to be hit exactly.
//
//    At the answer, the direction back to the base case is perpendicular to the surface of equal
//    price (the "tangency" condition): any step along that surface moves you away from the base.
//    To first order each driver moves in proportion to σ_k² × (sensitivity of the price to it), so
//    an assumption moves more when the price is more sensitive to it, and more when you say it is
//    less certain.
//
// Notation: d_k = change of driver k from the base case, u_k = d_k / σ_k (change in "normal moves").
// The solver works in u, where the distance is plain Euclidean length ‖u‖.

import type { Cell, Workbook } from './workbook';
import { Gradients } from './grad';
import { fullAddr, parseA1, toA1 } from './refs';
import { solveLinear } from './tape';
import type { Scalar } from './values';

export type DriverMode = 'shift' | 'scale';

/** A free assumption: one or more input cells that move together. */
export interface DriverSpec {
  /** Display name, e.g. "Revenue growth (every year)". */
  name: string;
  /** Input cells it moves: 'Assumptions!C5:G5', or a list of cells / ranges. */
  cells: string | string[];
  /**
   * 'shift' (default): every cell moves by the same amount d (growth 8%, 7%, ... → 8% + d, 7% + d, ...).
   * 'scale': every cell is multiplied by (1 + d).
   */
  mode?: DriverMode;
  /** A normal-sized move, in the driver's own units (its d). Default: see `defaultSigma`. */
  sigma?: number;
}

export interface Driver {
  name: string;
  mode: DriverMode;
  sigma: number;
  /** How sigma was chosen, in words. */
  sigmaRule: string;
  cells: Cell[];
  addrs: string[];
  /** Each cell's value in the base case. */
  base: number[];
}

export interface SolverOptions {
  /** The target counts as hit when |output − target| ≤ targetTol × max(1, |target|). Default 1e-12. */
  targetTol?: number;
  /**
   * Several drivers: the closest point counts as found when the part of the move that is NOT
   * along the price gradient is ≤ tangencyTol × the move (the sine of the angle). Default 1e-9.
   */
  tangencyTol?: number;
  maxIterations?: number;
  /** Up to this many drivers, Newton steps use a curvature matrix (finite differences of exact gradients). */
  hessianUpTo?: number;
  /** Finite-difference step for that matrix, in normal-sized moves (σ units). Default 1e-4. */
  hessianStep?: number;
}

export interface DriverResult {
  name: string;
  mode: DriverMode;
  sigma: number;
  sigmaRule: string;
  /** Change from the base case, in the driver's units. */
  change: number;
  /** Change in normal-sized moves (change / sigma). */
  normalMoves: number;
  /** Share of the total squared distance carried by this driver (sums to 1). */
  shareOfDistance: number;
  cells: { addr: string; base: number; implied: number }[];
}

export interface SolveResult {
  output: string;
  target: number;
  /** Output in the base case. */
  baseValue: number;
  converged: boolean;
  /** Output at the answer (recomputed), and how far it is from the target. */
  value: number;
  missBy: number;
  /** Change of each driver (d) and in normal-sized moves (u = d / σ). */
  d: number[];
  u: number[];
  /** ‖u‖: the weighted least-squares distance from the base case. */
  distance: number;
  /** Lagrange multiplier: u = λ × (gradient in σ units) at the answer. */
  lambda: number;
  /** Sine of the angle between the move u and the price gradient (0 at the exact closest point). */
  tangency: number;
  /**
   * Smallest curvature of the distance along the surface of equal price at the answer (>0: a true
   * local minimum). Only for 2+ drivers, when the curvature matrix was computed.
   */
  reducedHessianMinEig?: number;
  drivers: DriverResult[];
  /** The straight-line (first-order) prediction from the base-case gradient, for comparison. */
  linearPrediction: number[];
  gradientAtAnswer: number[];
  iterations: number;
  evaluations: { values: number; gradients: number };
  ms: number;
  method: string;
  /**
   * IF / MIN / MAX / ... decisions upstream of the output that the drivers could flip by moving less
   * than one normal-sized move (first order) from the answer. Near a kink the slope changes
   * abruptly, so the answer's neighbourhood looks different on the other side.
   */
  nearKinks: DriverKink[];
  warnings: string[];
}

export interface DriverKink {
  cell: string;
  kind: string;
  detail: string;
  /** The kink's margin (0 = exactly on it). */
  margin: number;
  /** First-order distance to the kink, in normal-sized moves of the drivers: |margin| / ‖∂margin/∂u‖. */
  normalMovesToFlip: number;
}

export interface CurvePoint { d: [number, number]; values: [number, number]; value: number; missBy: number }
export interface CurveBranch { points: CurvePoint[]; ends: [string, string]; closed: boolean }
export interface CurveResult {
  output: string;
  target: number;
  drivers: [string, string];
  /** The box searched, in driver changes d: [[lo_i, hi_i], [lo_j, hi_j]]. */
  box: [[number, number], [number, number]];
  branches: CurveBranch[];
  points: number;
  maxAbsMissBy: number;
  /** Brute-force check: a grid scan of the box, and whether the traced curve passes through every crossing it found. */
  grid: { n: number; crossings: number; covered: number; worstGap: number; invalidNodes: number };
  evaluations: { values: number; gradients: number };
  ms: number;
}

const EPS = Number.EPSILON;
const norm = (v: ArrayLike<number>) => { let s = 0; for (let i = 0; i < v.length; i++) s += v[i] * v[i]; return Math.sqrt(s); };
const dot = (a: ArrayLike<number>, b: ArrayLike<number>) => { let s = 0; for (let i = 0; i < a.length; i++) s += a[i] * b[i]; return s; };

/**
 * How far the move u is from pointing along the gradient G: the length of the part of u that is
 * perpendicular to G, divided by the length of u (the sine of the angle). Computed by projection,
 * not as sqrt(1 − cos²), which can't resolve angles below ~1e-8.
 */
export function tangencyOf(u: ArrayLike<number>, G: ArrayLike<number>): number {
  const nu = norm(u), g2 = dot(G, G);
  if (nu === 0 || g2 === 0) return 0;
  const c = dot(u, G) / g2;
  let s = 0;
  for (let i = 0; i < u.length; i++) { const r = u[i] - c * G[i]; s += r * r; }
  return Math.sqrt(s) / nu;
}

/** A sheet-qualified reference ('Sheet'!A1 or Sheet!A1:B3) → sheet name and corners. */
export function parseRef(ref: string): { sheet: string; r1: number; c1: number; r2: number; c2: number } {
  const i = ref.lastIndexOf('!');
  if (i < 0) throw new Error(`reference needs a sheet name: ${ref}`);
  const sheet = ref.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'");
  const [p, q] = ref.slice(i + 1).replace(/\$/g, '').split(':');
  const a = parseA1(p), b = parseA1(q ?? p);
  if (!a || !b) throw new Error(`bad reference ${ref}`);
  return { sheet, r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
}

export function cellAt(wb: Workbook, ref: string): Cell {
  const r = parseRef(ref);
  const c = wb.getCell(r.sheet, toA1(r.r1, r.c1));
  if (!c) throw new Error(`no cell ${ref}`);
  return c;
}

/** The default "normal-sized move" for a driver. */
export function defaultSigma(mode: DriverMode, base: number[]): { sigma: number; rule: string } {
  if (mode === 'scale') return { sigma: 0.1, rule: 'scale driver: a 10% change is one normal-sized move' };
  if (base.every((x) => Math.abs(x) <= 1)) return { sigma: 0.01, rule: 'values between −1 and 1 look like rates: 1 percentage point is one normal-sized move' };
  const m = base.reduce((s, x) => s + Math.abs(x), 0) / base.length;
  return { sigma: m > 0 ? 0.1 * m : 1, rule: '10% of the average base value is one normal-sized move' };
}

export class Backsolver {
  readonly output: Cell;
  readonly outputAddr: string;
  readonly drivers: Driver[];
  readonly opts: Required<SolverOptions>;
  /** Model evaluations so far: value-only recalculations, and recalculation + tape + backward pass. */
  counts = { values: 0, gradients: 0 };
  private last?: Gradients;

  constructor(readonly wb: Workbook, output: string, specs: DriverSpec[], opts: SolverOptions = {}) {
    this.opts = { targetTol: 1e-12, tangencyTol: 1e-9, maxIterations: 60, hessianUpTo: 40, hessianStep: 1e-4, ...opts };
    this.output = cellAt(wb, output);
    if (!this.output.node) throw new Error(`${output} is not a formula cell`);
    this.outputAddr = fullAddr(this.output.sheet, this.output.row, this.output.col);
    const seen = new Set<Cell>();
    this.drivers = specs.map((s) => {
      const refs = Array.isArray(s.cells) ? s.cells : [s.cells];
      const cells: Cell[] = [];
      for (const ref of refs) {
        const r = parseRef(ref);
        for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
          const c = wb.getCell(r.sheet, toA1(row, col));
          const at = `${r.sheet}!${toA1(row, col)}`;
          if (!c || c.node || c.arrayMaster) throw new Error(`driver "${s.name}": ${at} is not an input cell`);
          if (typeof c.value !== 'number') throw new Error(`driver "${s.name}": ${at} is not a number`);
          if (seen.has(c)) throw new Error(`driver "${s.name}": ${at} already belongs to another driver`);
          seen.add(c); cells.push(c);
        }
      }
      if (!cells.length) throw new Error(`driver "${s.name}" has no cells`);
      const mode = s.mode ?? 'shift';
      const base = cells.map((c) => c.value as number);
      const def = defaultSigma(mode, base);
      const sigma = s.sigma ?? def.sigma;
      if (!(sigma > 0)) throw new Error(`driver "${s.name}": sigma must be > 0`);
      return { name: s.name, mode, sigma, sigmaRule: s.sigma !== undefined ? 'set by the user' : def.rule, cells, addrs: cells.map((c) => fullAddr(c.sheet, c.row, c.col)), base };
    });
  }

  get size() { return this.drivers.length; }

  /** Write driver changes d into the input cells (no recalculation). */
  setCells(d: ArrayLike<number>) {
    this.drivers.forEach((dr, k) => dr.cells.forEach((c, i) => { c.value = dr.mode === 'shift' ? dr.base[i] + d[k] : dr.base[i] * (1 + d[k]); }));
  }

  /** Put the base case back and recalculate. */
  reset() {
    this.setCells(new Float64Array(this.size));
    this.wb.recalc({ tight: true, cycleStart: 'current' });
  }

  /** The output at driver changes d (NaN when it isn't a number or a loop didn't settle). */
  value(d: ArrayLike<number>): number {
    this.counts.values++;
    this.setCells(d);
    this.wb.recalc({ tight: true, cycleStart: 'current' });
    if (this.wb.cycles.some((c) => !c.converged)) return NaN;
    const v = this.output.value;
    return typeof v === 'number' ? v : NaN;
  }

  /** The output and its slope with respect to each driver: one recalculation, one tape, one backward pass. */
  valueAndGradient(d: ArrayLike<number>): { value: number; g: Float64Array } {
    this.counts.gradients++;
    this.setCells(d);
    const g = new Float64Array(this.size);
    let gr: Gradients;
    try { gr = new Gradients(this.wb); } catch { return { value: NaN, g: g.fill(NaN) }; }
    this.last = gr;
    if (typeof this.output.value !== 'number') return { value: NaN, g: g.fill(NaN) };
    const r = gr.gradient(this.output);
    const idx = new Map<Cell, number>();
    gr.inputs.forEach((x, i) => idx.set(x.cell, i));
    this.drivers.forEach((dr, k) => {
      let s = 0;
      dr.cells.forEach((c, i) => { const j = idx.get(c); if (j !== undefined) s += r.d[j] * (dr.mode === 'shift' ? 1 : dr.base[i]); });
      g[k] = s;
    });
    return { value: this.output.value, g };
  }

  /** Values of other cells at driver changes d (e.g. an implied revenue CAGR). Leaves the workbook at the base case. */
  cellsAt(d: ArrayLike<number>, refs: string[]): Record<string, Scalar> {
    this.value(d);
    const out = Object.fromEntries(refs.map((r) => [r, cellAt(this.wb, r).value]));
    this.reset();
    return out;
  }

  /**
   * d(answer)/d(another input) for a one-driver answer, by the implicit function theorem: the
   * answer keeps output(d, p) = target, so dd/dp = −(∂output/∂p) / (∂output/∂d).
   */
  sensitivity1D(res: SolveResult, inputRef: string, k = 0): { dAnswer: number; dOutputDInput: number; dOutputDDriver: number } {
    const p = cellAt(this.wb, inputRef);
    const { g } = this.valueAndGradient(res.d);
    const gr = this.last!;
    const j = gr.inputs.findIndex((x) => x.cell === p);
    const dp = j < 0 ? 0 : gr.gradient(this.output).d[j];
    this.reset();
    return { dAnswer: -dp / g[k], dOutputDInput: dp, dOutputDDriver: g[k] };
  }

  // ---------------------------------------------------------------- one free driver (Goal Seek)

  /**
   * One free driver k (the others stay at the base case): find d with output(d) = target, starting
   * from the base case. Newton's method with the exact slope; once the target is bracketed, any
   * Newton step that leaves the bracket is replaced by bisection, so it cannot run away.
   */
  solve1D(target: number, k = 0): SolveResult {
    const t0 = performance.now();
    const c0 = { ...this.counts };
    const m = this.size;
    const at = (x: number) => { const d = new Float64Array(m); d[k] = x; return d; };
    const fdf = (x: number) => { const r = this.valueAndGradient(at(x)); return { f: r.value - target, df: r.g[k] }; };
    const tol = this.opts.targetTol * Math.max(1, Math.abs(target));
    const r = newtonBracketed(fdf, 0, tol, this.drivers[k].sigma, this.opts.maxIterations * 2);
    const res = this.finish(target, at(r.x), r.converged, r.iterations, 'Newton with bisection fallback (one driver)', t0, c0, r.warnings);
    return res;
  }

  // ---------------------------------------------------------------- several free drivers

  /**
   * The closest assumptions (smallest Σ (d_k/σ_k)²) for which the output equals the target exactly.
   * With one driver this is the root nearest the base case.
   *
   * Method: Newton's method on the optimality conditions u = λ·G(u), output(u) = target (G = the
   * price gradient in σ units). Exact gradients come from the Stage 2 tape; the curvature matrix
   * comes from finite differences of those exact gradients, and only affects the speed, not the
   * answer: convergence is judged on exact quantities. A line search keeps every step honest, and if
   * the full target is too far for Newton, we walk the target there in smaller steps (continuation).
   */
  solve(target: number): SolveResult {
    if (this.size === 1) return this.solve1D(target, 0);
    const t0 = performance.now();
    const c0 = { ...this.counts };
    const m = this.size;
    const sig = this.drivers.map((d) => d.sigma);
    const warnings: string[] = [];
    const tolF = this.opts.targetTol * Math.max(1, Math.abs(target));
    const evalU = (u: ArrayLike<number>) => {
      const r = this.valueAndGradient(Float64Array.from(u, (x, i) => x * sig[i]));
      return { P: r.value, G: Float64Array.from(r.g, (x, i) => x * sig[i]) };
    };
    const base = evalU(new Float64Array(m));
    if (!Number.isFinite(base.P)) throw new Error(`${this.outputAddr} is not a number in the base case`);
    const Gref = norm(base.G);
    if (Gref === 0) throw new Error(`${this.outputAddr} does not depend on any driver`);
    const useH = m <= this.opts.hessianUpTo;
    const hess = (u: Float64Array): Float64Array => {
      const H = new Float64Array(m * m);
      if (!useH) return H;
      const h = this.opts.hessianStep;
      for (let j = 0; j < m; j++) {
        const up = Float64Array.from(u), dn = Float64Array.from(u);
        up[j] += h; dn[j] -= h;
        const a = evalU(up).G, b = evalU(dn).G;
        for (let i = 0; i < m; i++) H[i * m + j] = (a[i] - b[i]) / (2 * h);
      }
      for (let i = 0; i < m; i++) for (let j = 0; j < i; j++) { const s = (H[i * m + j] + H[j * m + i]) / 2; H[i * m + j] = s; H[j * m + i] = s; }
      return H;
    };
    // Merit: how far from optimal. Stationarity ‖u − λG‖² plus the miss expressed in σ units.
    const merit = (u: Float64Array, lam: number, P: number, G: Float64Array, T: number) => {
      let s = 0; for (let i = 0; i < m; i++) { const r = u[i] - lam * G[i]; s += r * r; }
      const f = (P - T) / Gref;
      return s + f * f;
    };

    // Newton-KKT from (u, λ) toward target T. Returns the final iterate and whether it converged.
    let iterations = 0;
    const kkt = (u0: Float64Array, lam0: number, T: number, final: boolean) => {
      let u = Float64Array.from(u0), lam = lam0;
      let { P, G } = evalU(u);
      const tolT = this.opts.targetTol * Math.max(1, Math.abs(T));
      for (let it = 0; it < this.opts.maxIterations; it++) {
        if (!Number.isFinite(P) || G.some((x) => !Number.isFinite(x))) return { ok: false, u, lam, P, G };
        // Converged? Judged on exact quantities: the miss, and the tangency of u to G.
        // (Intermediate targets of the continuation only need to be roughly right.)
        const tang = tangencyOf(u, G);
        if (Math.abs(P - T) <= (final ? tolF : tolT) && tang <= (final ? this.opts.tangencyTol : 1e-6)) return { ok: true, u, lam: dot(u, G) / dot(G, G), P, G };
        iterations++;
        const H = hess(u);
        let step: { du: Float64Array; dl: number } | null = null;
        for (const withH of useH ? [true, false] : [false]) {
          const K = new Float64Array((m + 1) * (m + 1));
          const n1 = m + 1;
          for (let i = 0; i < m; i++) {
            for (let j = 0; j < m; j++) K[i * n1 + j] = (i === j ? 1 : 0) - (withH ? lam * H[i * m + j] : 0);
            K[i * n1 + m] = -G[i];
            K[m * n1 + i] = G[i];
          }
          const rhs: number[] = [];
          for (let i = 0; i < m; i++) rhs.push(-(u[i] - lam * G[i]));
          rhs.push(-(P - T));
          let sol: number[];
          try { sol = solveLinear(K, n1, rhs); } catch { continue; }
          if (sol.some((x) => !Number.isFinite(x))) continue;
          const du = Float64Array.from(sol.slice(0, m)), dl = sol[m];
          // Backtracking line search on the merit function.
          const phi0 = merit(u, lam, P, G, T);
          let alpha = 1, accepted = false;
          for (let ls = 0; ls < 30; ls++, alpha /= 2) {
            const un = Float64Array.from(u, (x, i) => x + alpha * du[i]);
            const ln = lam + alpha * dl;
            const e = evalU(un);
            if (!Number.isFinite(e.P) || e.G.some((x) => !Number.isFinite(x))) continue;
            const phi = merit(un, ln, e.P, e.G, T);
            if (phi <= (1 - 1e-4 * alpha) * phi0 || phi < 1e-30) { u = un; lam = ln; P = e.P; G = e.G; accepted = true; break; }
          }
          if (accepted) { step = { du, dl }; break; }
        }
        if (!step) return { ok: false, u, lam, P, G };
      }
      return { ok: false, u, lam, P, G };
    };

    // Start from the straight-line answer: u = λ0·G0 with λ0 chosen so the linearised output hits T.
    const P0 = base.P;
    const lin = Float64Array.from(base.G, (x) => (x * (target - P0)) / (Gref * Gref));
    let u = new Float64Array(m), lam = 0, s = 0, ds = 1, converged = false, method = 'Newton on the optimality conditions';
    if (Math.abs(P0 - target) <= tolF) { converged = true; s = 1; }
    while (s < 1) {
      const sn = Math.min(1, s + ds), T = P0 + sn * (target - P0);
      const lam0 = s === 0 ? (T - P0) / (Gref * Gref) : lam;
      const u0 = s === 0 ? Float64Array.from(base.G, (x) => x * lam0) : u;
      const r = kkt(u0, lam0, T, sn === 1);
      if (r.ok) { u = r.u; lam = r.lam; s = sn; ds = Math.min(1 - s, 2 * ds); converged = s === 1; continue; }
      ds /= 2;
      if (ds < 1e-6) { warnings.push(`could not reach the target: stuck at ${(100 * s).toFixed(4)}% of the way from the base output`); u = r.u; break; }
      method = 'Newton on the optimality conditions, with the target walked in from the base output (continuation)';
    }
    if (!useH) method += '; no curvature matrix (too many drivers), so steps are minimum-norm projections';
    const res = this.finish(target, Float64Array.from(u, (x, i) => x * sig[i]), converged, iterations, method, t0, c0, warnings, Float64Array.from(lin, (x, i) => x * sig[i]));
    // Second-order check: curvature of the distance along the surface of equal price.
    if (converged && useH) {
      const uu = Float64Array.from(res.u);
      const G = Float64Array.from(res.gradientAtAnswer, (x, i) => x * sig[i]);
      const H = hess(uu);
      this.reset();
      res.reducedHessianMinEig = reducedMinEig(H, G, res.lambda, m);
      if (!(res.reducedHessianMinEig > 0)) res.warnings.push('the answer is a stationary point but not a local minimum of the distance (the curvature along the price surface is not positive)');
      res.evaluations = { values: this.counts.values - c0.values, gradients: this.counts.gradients - c0.gradients };
    }
    return res;
  }

  /** Kinks upstream of the output that the drivers can flip within `within` normal-sized moves (first order). */
  kinksNear(gr: Gradients, within = 1): DriverKink[] {
    const up = gr.upstreamCells(this.output);
    const idx = new Map<Cell, number>();
    gr.inputs.forEach((x, i) => idx.set(x.cell, i));
    const out: DriverKink[] = [];
    for (const k of gr.kinks) {
      if (!up.has(k.cell)) continue;
      const adj = gr.rec.tape.backward([[k.margin.i, 1]]);
      let s2 = 0;
      for (const dr of this.drivers) {
        let s = 0;
        dr.cells.forEach((c, i) => { const j = idx.get(c); if (j !== undefined) s += adj[gr.inputs[j].node] * (dr.mode === 'shift' ? 1 : dr.base[i]); });
        s2 += (s * dr.sigma) ** 2;
      }
      if (s2 === 0) continue; // the drivers can't move this decision
      const moves = Math.abs(k.margin.v) / Math.sqrt(s2);
      if (moves <= within) out.push({ cell: k.cell, kind: k.kind, detail: k.detail, margin: k.margin.v, normalMovesToFlip: moves });
    }
    return out.sort((a, b) => a.normalMovesToFlip - b.normalMovesToFlip);
  }

  /** Recompute everything at the answer, describe it, and put the base case back. */
  private finish(target: number, d: Float64Array, converged: boolean, iterations: number, method: string, t0: number, c0: { values: number; gradients: number }, warnings: string[], linear?: Float64Array): SolveResult {
    const m = this.size;
    const zero = new Float64Array(m);
    const baseValue = this.value(zero);
    const { value, g } = this.valueAndGradient(d);
    const near = this.last ? this.kinksNear(this.last) : [];
    const sig = this.drivers.map((x) => x.sigma);
    const u = Array.from(d, (x, i) => x / sig[i]);
    const G = Array.from(g, (x, i) => x * sig[i]);
    const dist = norm(u);
    const lambda = dot(G, G) > 0 ? dot(u, G) / dot(G, G) : 0;
    const tangency = tangencyOf(u, G);
    if (!linear) {
      // Straight-line prediction from the base gradient.
      const b = this.valueAndGradient(zero);
      const Gb = Array.from(b.g, (x, i) => x * sig[i]);
      const n2 = dot(Gb, Gb);
      linear = Float64Array.from(Gb, (x, i) => (n2 > 0 ? (x * (target - b.value)) / n2 : NaN) * sig[i]);
    }
    const drivers: DriverResult[] = this.drivers.map((dr, k) => ({
      name: dr.name, mode: dr.mode, sigma: dr.sigma, sigmaRule: dr.sigmaRule, change: d[k], normalMoves: u[k],
      shareOfDistance: dist > 0 ? (u[k] * u[k]) / (dist * dist) : 0,
      cells: dr.cells.map((c, i) => ({ addr: dr.addrs[i], base: dr.base[i], implied: dr.mode === 'shift' ? dr.base[i] + d[k] : dr.base[i] * (1 + d[k]) })),
    }));
    if (!converged) warnings.push('did not converge: the numbers below are the best point found, not an answer');
    if (near.length) warnings.push(`${near.length} kink(s) would flip if the drivers moved less than one normal-sized move from the answer: ${near.slice(0, 3).map((k) => `${k.cell} ${k.kind} (${k.normalMovesToFlip.toPrecision(2)})`).join(', ')}`);
    this.reset();
    return {
      output: this.outputAddr, target, baseValue, converged, value, missBy: value - target,
      d: Array.from(d), u, distance: dist, lambda, tangency, drivers, linearPrediction: Array.from(linear),
      gradientAtAnswer: Array.from(g), iterations,
      evaluations: { values: this.counts.values - c0.values, gradients: this.counts.gradients - c0.gradients },
      ms: performance.now() - t0, method, nearKinks: near, warnings,
    };
  }

  // ---------------------------------------------------------------- the curve of equal price

  /**
   * Every combination of drivers i and j (the others at the base case) that gives exactly the
   * target, inside a box of changes. Traced by predictor-corrector continuation: step along the
   * curve's tangent, then pull back onto the curve with Newton steps along the gradient.
   *
   * Then a brute-force check: evaluate the model on an n×n grid over the box, find every grid edge
   * where output − target changes sign (the curve must cross there), and confirm the traced curve
   * passes through each crossing. A crossing that isn't covered means another branch: we trace
   * from it too, until every crossing is covered.
   */
  traceCurve(target: number, i: number, j: number, box: [[number, number], [number, number]], opts: { grid?: number; maxPoints?: number; seed?: [number, number] } = {}): CurveResult {
    const t0 = performance.now();
    const c0 = { ...this.counts };
    const m = this.size;
    const n = opts.grid ?? 41;
    const sI = this.drivers[i].sigma, sJ = this.drivers[j].sigma;
    const tolF = this.opts.targetTol * Math.max(1, Math.abs(target));
    // Work in σ units so a step means the same in both directions.
    const lo = [box[0][0] / sI, box[1][0] / sJ], hi = [box[0][1] / sI, box[1][1] / sJ];
    const diag = Math.hypot(hi[0] - lo[0], hi[1] - lo[1]);
    const toD = (a: number, b: number) => { const d = new Float64Array(m); d[i] = a * sI; d[j] = b * sJ; return d; };
    const F = (a: number, b: number) => this.value(toD(a, b)) - target;
    const FG = (a: number, b: number) => { const r = this.valueAndGradient(toD(a, b)); return { f: r.value - target, gx: r.g[i] * sI, gy: r.g[j] * sJ }; };
    const inside = (a: number, b: number) => a >= lo[0] - 1e-12 && a <= hi[0] + 1e-12 && b >= lo[1] - 1e-12 && b <= hi[1] + 1e-12;

    // 1. Grid scan.
    const gx = (k: number) => lo[0] + ((hi[0] - lo[0]) * k) / (n - 1), gy = (k: number) => lo[1] + ((hi[1] - lo[1]) * k) / (n - 1);
    const val: number[][] = [];
    let invalid = 0;
    for (let a = 0; a < n; a++) { val.push([]); for (let b = 0; b < n; b++) { const v = F(gx(a), gy(b)); if (!Number.isFinite(v)) invalid++; val[a].push(v); } }
    // Every grid edge where the sign changes, located exactly by 1-D root finding along the edge.
    const crossings: [number, number][] = [];
    const refine = (ax: number, ay: number, bx: number, by: number, fa: number, fb: number): [number, number] => {
      const x = rootBetween((t) => F(ax + t * (bx - ax), ay + t * (by - ay)), 0, 1, fa, fb, tolF);
      return [ax + x * (bx - ax), ay + x * (by - ay)];
    };
    for (let a = 0; a < n; a++) for (let b = 0; b < n; b++) {
      const v = val[a][b];
      if (!Number.isFinite(v)) continue;
      if (v === 0) { crossings.push([gx(a), gy(b)]); continue; }
      if (a + 1 < n && Number.isFinite(val[a + 1][b]) && v * val[a + 1][b] < 0) crossings.push(refine(gx(a), gy(b), gx(a + 1), gy(b), v, val[a + 1][b]));
      if (b + 1 < n && Number.isFinite(val[a][b + 1]) && v * val[a][b + 1] < 0) crossings.push(refine(gx(a), gy(b), gx(a), gy(b + 1), v, val[a][b + 1]));
    }

    // 2. Continuation from a point on the curve, in one direction.
    const hMax = diag / 150, hMin = diag * 1e-9, cornerH = hMax / 1024;
    const maxPoints = opts.maxPoints ?? 4000;
    const project = (a: number, b: number) => {
      // Newton steps along the gradient (minimum-norm correction) until the target is hit.
      for (let k = 0; k < 12; k++) {
        const e = FG(a, b);
        if (!Number.isFinite(e.f)) return null;
        if (Math.abs(e.f) <= tolF) return { a, b, ...e, k };
        const g2 = e.gx * e.gx + e.gy * e.gy;
        if (g2 === 0) return null;
        a -= (e.f * e.gx) / g2; b -= (e.f * e.gy) / g2;
      }
      return null;
    };
    const walk = (start: { a: number; b: number; gx: number; gy: number }, sign: 1 | -1) => {
      const pts: [number, number][] = [];
      let p = start, h = hMax / 4;
      let tx = (-p.gy * sign) / Math.hypot(p.gx, p.gy), ty = (p.gx * sign) / Math.hypot(p.gx, p.gy);
      let reason = 'max points';
      let mx = tx, my = ty; // direction of the last actual move
      // One predictor-corrector step in direction (dx, dy). Rejected if the correction had to move
      // the point far (the curve isn't where the tangent said), or if we didn't actually advance.
      const step = (dx: number, dy: number) => {
        const qa = p.a + h * dx, qb = p.b + h * dy;
        const q = project(qa, qb);
        if (!q || Math.hypot(q.a - qa, q.b - qb) > 0.5 * h) return null;
        const ma = q.a - p.a, mb = q.b - p.b, len = Math.hypot(ma, mb);
        if (len < 0.25 * h || ma * mx + mb * my < -0.1 * len) return null; // stalled, or going back
        const gn = Math.hypot(q.gx, q.gy);
        let nx = -q.gy / gn, ny = q.gx / gn;
        if (nx * ma + ny * mb < 0) { nx = -nx; ny = -ny; } // keep travelling the same way
        // Landed exactly on a corner (an IF/MIN/MAX tie): the slope there belongs to the other arm
        // and its tangent is across our path. Keep going the way we moved instead.
        if (Math.abs(nx * ma + ny * mb) < 0.5 * len) { nx = ma / len; ny = mb / len; }
        return { q, nx, ny, ma: ma / len, mb: mb / len };
      };
      while (pts.length < maxPoints) {
        let r = step(tx, ty);
        let turned = false;
        // Turned more than 15° in one step: take a smaller step, unless the step is already tiny.
        if (r && r.nx * tx + r.ny * ty < Math.cos(Math.PI / 12) && h > cornerH) r = null;
        if (!r && h <= cornerH) {
          // Still failing with a tiny step: a corner, where an IF / MIN / MAX flips and the curve
          // changes direction abruptly. Look for the way on by turning the predictor up to 90°.
          for (const deg of [15, -15, 30, -30, 45, -45, 60, -60, 75, -75, 90, -90]) {
            const c = Math.cos((deg * Math.PI) / 180), sn = Math.sin((deg * Math.PI) / 180);
            r = step(c * tx - sn * ty, sn * tx + c * ty);
            if (r) { turned = true; break; }
          }
        }
        if (!r) {
          h /= 2;
          if (h < hMin) { reason = project(p.a + 2 * hMin * tx, p.b + 2 * hMin * ty) === null ? 'the model stops giving a number (or the slope vanishes) here' : 'step too small (a sharp turn the tracer could not follow)'; break; }
          continue;
        }
        const q = r.q;
        if (!inside(q.a, q.b)) {
          // Find exactly where the curve leaves the box: root of F on the box edge it crossed.
          const e = boxExit(p.a, p.b, q.a, q.b, lo, hi);
          const pt = e && edgeRoot(e, p.a, p.b, q.a, q.b, F, tolF);
          const last = pts.length ? pts[pts.length - 1] : [p.a, p.b];
          if (pt && Math.hypot(pt[0] - last[0], pt[1] - last[1]) > 1e-12 * diag) pts.push(pt);
          reason = 'edge of the box';
          break;
        }
        pts.push([q.a, q.b]);
        // Closed curve: back near the start after going around.
        if (pts.length > 10 && Math.hypot(q.a - start.a, q.b - start.b) < 0.75 * h) { reason = 'closed curve'; break; }
        tx = r.nx; ty = r.ny; mx = r.ma; my = r.mb; p = q;
        if (q.k <= 2 && !turned) h = Math.min(hMax, h * 1.5);
      }
      return { pts, reason };
    };

    // 3. Trace branches until every grid crossing is covered.
    const branches: { pts: [number, number][]; ends: [string, string]; closed: boolean }[] = [];
    const gap = (p: [number, number]) => {
      let best = Infinity;
      for (const br of branches) for (let k = 0; k < br.pts.length; k++) {
        const a = br.pts[k], b = br.pts[k + 1] ?? (br.closed ? br.pts[0] : undefined);
        best = Math.min(best, b ? segDist(p, a, b) : Math.hypot(p[0] - a[0], p[1] - a[1]));
      }
      return best;
    };
    const cell = Math.min((hi[0] - lo[0]) / (n - 1), (hi[1] - lo[1]) / (n - 1));
    const coverTol = 0.1 * cell;
    const seeds: [number, number][] = [];
    if (opts.seed) seeds.push([opts.seed[0] / sI, opts.seed[1] / sJ]);
    let guard = 0;
    for (;;) {
      let s = seeds.shift();
      if (!s) { const unc = crossings.find((c) => gap(c) > coverTol); if (!unc) break; s = unc; }
      if (++guard > 50) break;
      const st = project(s[0], s[1]);
      if (!st || !inside(st.a, st.b)) continue;
      const fw = walk(st, 1);
      let pts: [number, number][], ends: [string, string], closed = false;
      if (fw.reason === 'closed curve') { pts = [[st.a, st.b], ...fw.pts]; ends = ['closed curve', 'closed curve']; closed = true; }
      else { const bw = walk(st, -1); pts = [...bw.pts.reverse(), [st.a, st.b], ...fw.pts]; ends = [bw.reason, fw.reason]; }
      branches.push({ pts, ends, closed });
    }
    let worstGap = 0, covered = 0;
    for (const c of crossings) { const g = gap(c); worstGap = Math.max(worstGap, g * Math.min(sI, sJ)); if (g <= coverTol) covered++; }

    // 4. Report every point with its exact output (recomputed).
    let maxMiss = 0, count = 0;
    const out: CurveBranch[] = branches.map((br) => ({
      ends: br.ends, closed: br.closed,
      points: br.pts.map(([a, b]) => {
        const v = this.value(toD(a, b));
        maxMiss = Math.max(maxMiss, Math.abs(v - target)); count++;
        return { d: [a * sI, b * sJ] as [number, number], values: [this.meanImplied(i, a * sI), this.meanImplied(j, b * sJ)] as [number, number], value: v, missBy: v - target };
      }),
    }));
    this.reset();
    return {
      output: this.outputAddr, target, drivers: [this.drivers[i].name, this.drivers[j].name], box, branches: out, points: count, maxAbsMissBy: maxMiss,
      grid: { n, crossings: crossings.length, covered, worstGap, invalidNodes: invalid },
      evaluations: { values: this.counts.values - c0.values, gradients: this.counts.gradients - c0.gradients }, ms: performance.now() - t0,
    };
  }

  /** Average implied cell value of driver k at change d (e.g. the average implied growth rate). */
  meanImplied(k: number, d: number): number {
    const dr = this.drivers[k];
    return dr.base.reduce((s, x) => s + (dr.mode === 'shift' ? x + d : x * (1 + d)), 0) / dr.base.length;
  }
}

// ---------------------------------------------------------------- helpers

/**
 * Safeguarded Newton for f(x) = 0 from x0. Unbracketed steps are halved until |f| decreases (or
 * the sign flips, which brackets the root). Once bracketed, Newton steps that leave the bracket,
 * or don't at least halve |f| twice running, are replaced by bisection.
 */
export function newtonBracketed(fdf: (x: number) => { f: number; df: number }, x0: number, tol: number, scale: number, maxIt = 120) {
  const warnings: string[] = [];
  let x = x0, { f, df } = fdf(x);
  if (!Number.isFinite(f)) throw new Error('the output is not a number at the starting point');
  let best = { x, f };
  let A: { x: number; f: number } | null = null, B: { x: number; f: number } | null = null;
  let slow = 0, probe = 0, it = 0;
  for (it = 1; it <= maxIt; it++) {
    if (Math.abs(f) <= tol) return { x, f, iterations: it - 1, converged: true, warnings };
    let xn: number;
    const newton = df !== 0 && Number.isFinite(df) ? x - f / df : NaN;
    if (A && B) {
      const a = Math.min(A.x, B.x), b = Math.max(A.x, B.x);
      xn = Number.isFinite(newton) && newton > a && newton < b && slow < 2 ? newton : (A.x + B.x) / 2;
      if (b - a <= 4 * EPS * Math.max(1, Math.abs(x))) { warnings.push('stopped at the limit of floating-point precision'); break; }
    } else if (Number.isFinite(newton)) xn = newton;
    else { probe++; xn = x0 + scale * 2 ** (probe >> 1) * (probe & 1 ? 1 : -1); } // flat: probe outward both ways
    let r = fdf(xn);
    if (!(A && B)) {
      // Damped Newton: halve the step until |f| shrinks or the sign flips.
      for (let k = 0; k < 40 && (!Number.isFinite(r.f) || (Math.sign(r.f) === Math.sign(f) && Math.abs(r.f) >= Math.abs(f))); k++) {
        if (!Number.isFinite(newton)) break;
        xn = x + (xn - x) / 2; r = fdf(xn);
      }
    } else if (!Number.isFinite(r.f)) {
      // Inside a bracket but the model errors here: bisect toward the current point instead.
      for (let k = 0; k < 60 && !Number.isFinite(r.f); k++) { xn = (x + xn) / 2; r = fdf(xn); }
    }
    if (!Number.isFinite(r.f)) { warnings.push(`the output is not a number near ${xn}`); continue; }
    slow = Math.abs(r.f) > 0.5 * Math.abs(f) ? slow + 1 : 0;
    if (Math.sign(r.f) !== Math.sign(f) && r.f !== 0) { A = { x, f }; B = { x: xn, f: r.f }; }
    else if (A && B) { if (Math.sign(r.f) === Math.sign(A.f)) A = { x: xn, f: r.f }; else B = { x: xn, f: r.f }; }
    x = xn; f = r.f; df = r.df;
    if (Math.abs(f) < Math.abs(best.f)) best = { x, f };
  }
  return { x: best.x, f: best.f, iterations: it, converged: Math.abs(best.f) <= tol, warnings };
}

/** Root of f between a and b (f(a), f(b) of opposite signs), value-only: Illinois false position with bisection safeguard. */
export function rootBetween(f: (x: number) => number, a: number, b: number, fa: number, fb: number, tol: number): number {
  let side = 0;
  for (let it = 0; it < 200; it++) {
    let c = (a * fb - b * fa) / (fb - fa);
    if (!(c > Math.min(a, b) && c < Math.max(a, b))) c = (a + b) / 2;
    let fc = f(c);
    if (!Number.isFinite(fc)) { c = (a + b) / 2; fc = f(c); if (!Number.isFinite(fc)) return c; }
    if (Math.abs(fc) <= tol || Math.abs(b - a) <= 4 * EPS * Math.max(1, Math.abs(c))) return c;
    if (fc * fb < 0) { a = b; fa = fb; b = c; fb = fc; side = 0; }
    else { b = c; fb = fc; if (side === 1) fa /= 2; side = 1; }
    if (it % 8 === 7) { const mid = (a + b) / 2, fm = f(mid); if (Number.isFinite(fm)) { if (fm * fb < 0) { a = mid; fa = fm; } else { b = mid; fb = fm; } } }
  }
  return (a + b) / 2;
}

function segDist(p: [number, number], a: [number, number], b: [number, number]) {
  const vx = b[0] - a[0], vy = b[1] - a[1];
  const L = vx * vx + vy * vy;
  const t = L === 0 ? 0 : Math.max(0, Math.min(1, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L));
  return Math.hypot(p[0] - a[0] - t * vx, p[1] - a[1] - t * vy);
}

/** Which box edge the segment p→q leaves through: { axis, at } (axis 0 = first coordinate fixed). */
function boxExit(pa: number, pb: number, qa: number, qb: number, lo: number[], hi: number[]): { axis: 0 | 1; at: number } | null {
  let best: { axis: 0 | 1; at: number; t: number } | null = null;
  const tryEdge = (axis: 0 | 1, at: number) => {
    const p = axis === 0 ? pa : pb, q = axis === 0 ? qa : qb;
    if (p === q) return;
    const t = (at - p) / (q - p);
    if (t < 0 || t > 1) return;
    const o = axis === 0 ? pb + t * (qb - pb) : pa + t * (qa - pa);
    const [olo, ohi] = axis === 0 ? [lo[1], hi[1]] : [lo[0], hi[0]];
    if (o < olo - 1e-12 || o > ohi + 1e-12) return;
    if (!best || t < best.t) best = { axis, at, t };
  };
  tryEdge(0, lo[0]); tryEdge(0, hi[0]); tryEdge(1, lo[1]); tryEdge(1, hi[1]);
  return best;
}

/** The exact point where the curve meets a box edge, between the inside point p and the outside point q. */
function edgeRoot(e: { axis: 0 | 1; at: number }, pa: number, pb: number, qa: number, qb: number, F: (a: number, b: number) => number, tol: number): [number, number] | null {
  // On the edge, the other coordinate is free: search between where p and q project onto the edge,
  // widened a little, for a sign change.
  const lo = e.axis === 0 ? Math.min(pb, qb) : Math.min(pa, qa), hi = e.axis === 0 ? Math.max(pb, qb) : Math.max(pa, qa);
  const w = Math.max(hi - lo, 1e-9);
  const g = (o: number) => (e.axis === 0 ? F(e.at, o) : F(o, e.at));
  for (const pad of [0, 0.5, 2, 8]) {
    const a = lo - pad * w, b = hi + pad * w;
    const fa = g(a), fb = g(b);
    if (Number.isFinite(fa) && Number.isFinite(fb) && fa * fb <= 0) {
      const o = fa === 0 ? a : fb === 0 ? b : rootBetween(g, a, b, fa, fb, tol);
      return e.axis === 0 ? [e.at, o] : [o, e.at];
    }
  }
  return null;
}

/** Smallest eigenvalue of (I − λH) restricted to the directions perpendicular to G. */
function reducedMinEig(H: Float64Array, G: Float64Array, lambda: number, m: number): number {
  const gn = norm(G);
  const gh = Float64Array.from(G, (x) => x / gn);
  // Orthonormal basis of G⊥ by Gram-Schmidt on the unit vectors.
  const Z: Float64Array[] = [];
  for (let k = 0; k < m && Z.length < m - 1; k++) {
    const v = new Float64Array(m); v[k] = 1;
    const sub = (w: Float64Array) => { const c = dot(v, w); for (let i = 0; i < m; i++) v[i] -= c * w[i]; };
    sub(gh); for (const z of Z) sub(z);
    const nv = norm(v);
    if (nv > 1e-8) Z.push(Float64Array.from(v, (x) => x / nv));
  }
  const r = Z.length;
  const M = new Float64Array(r * r);
  for (let a = 0; a < r; a++) for (let b = 0; b < r; b++) {
    let s = 0;
    for (let i = 0; i < m; i++) for (let j = 0; j < m; j++) s += Z[a][i] * ((i === j ? 1 : 0) - lambda * H[i * m + j]) * Z[b][j];
    M[a * r + b] = s;
  }
  return Math.min(...jacobiEigenvalues(M, r));
}

/** Eigenvalues of a small symmetric matrix (cyclic Jacobi rotations). */
export function jacobiEigenvalues(S: Float64Array, n: number): number[] {
  const A = Float64Array.from(S);
  for (let sweep = 0; sweep < 100; sweep++) {
    let off = 0;
    for (let p = 0; p < n; p++) for (let q = p + 1; q < n; q++) off += A[p * n + q] ** 2;
    if (off < 1e-30) break;
    for (let p = 0; p < n; p++) for (let q = p + 1; q < n; q++) {
      const apq = A[p * n + q];
      if (Math.abs(apq) < 1e-300) continue;
      const theta = (A[q * n + q] - A[p * n + p]) / (2 * apq);
      const t = Math.sign(theta || 1) / (Math.abs(theta) + Math.sqrt(theta * theta + 1));
      const c = 1 / Math.sqrt(t * t + 1), s = t * c;
      for (let k = 0; k < n; k++) {
        const akp = A[k * n + p], akq = A[k * n + q];
        A[k * n + p] = c * akp - s * akq; A[k * n + q] = s * akp + c * akq;
      }
      for (let k = 0; k < n; k++) {
        const apk = A[p * n + k], aqk = A[q * n + k];
        A[p * n + k] = c * apk - s * aqk; A[q * n + k] = s * apk + c * aqk;
      }
    }
  }
  return Array.from({ length: n }, (_, i) => A[i * n + i]);
}

/** Excel-style Goal Seek: set `setCell` to `toValue` by changing the single input `changingCell`. */
export function goalSeek(wb: Workbook, setCell: string, toValue: number, changingCell: string, opts: SolverOptions = {}) {
  const s = new Backsolver(wb, setCell, [{ name: changingCell, cells: changingCell }], opts);
  const res = s.solve1D(toValue, 0);
  return { answer: res.drivers[0].cells[0].implied, result: res, solver: s };
}

