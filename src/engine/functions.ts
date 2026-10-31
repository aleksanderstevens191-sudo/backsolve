// Built-in Excel functions.
//
// Functions receive their arguments as unevaluated AST nodes, so IF can skip the branch it
// doesn't take (just like Excel) and functions like SUM can tell "a value that came from a
// cell reference" (text is ignored) from "a value typed into the formula" (text is an error).

import { Node } from './parser';
import { Matrix, RangeRef, Scalar, Value, XLError, isError } from './values';
import { toBool, toMatrix, toNumber } from './coerce';

export interface Ctx { sheet: string; row: number; col: number; array: boolean }
export interface Evaluator {
  eval(n: Node, ctx: Ctx): Value;
  /** Evaluate to a single value, applying Excel's implicit intersection to ranges. */
  scalar(n: Node, ctx: Ctx): Scalar;
}
export type Fn = (args: Node[], ctx: Ctx, ev: Evaluator) => Value;

const err = (code: XLError['code'], detail?: string) => new XLError(code, detail);
const fromRefOrArray = (v: Value) => v instanceof RangeRef || v instanceof Matrix;

/**
 * Gather numbers the way SUM/AVERAGE/MIN/MAX/NPV do:
 *  - from a reference or array: numbers only (text, TRUE/FALSE and blanks are skipped), errors propagate;
 *  - from a value typed directly: TRUE=1, "3"=3, other text is #VALUE!.
 */
export function collectNumbers(args: Node[], ctx: Ctx, ev: Evaluator): number[] | XLError {
  const out: number[] = [];
  for (const a of args) {
    if (a.t === 'missing') { out.push(0); continue; }
    const v = ev.eval(a, ctx);
    if (fromRefOrArray(v)) {
      for (const x of toMatrix(v).values()) {
        if (typeof x === 'number') out.push(x);
        else if (isError(x)) return x;
      }
    } else {
      const n = toNumber(v as Scalar);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function num(n: Node | undefined, ctx: Ctx, ev: Evaluator, dflt?: number): number | XLError {
  if (!n || n.t === 'missing') return dflt ?? 0;
  return toNumber(ev.scalar(n, ctx));
}

const checkNum = (x: number): number | XLError => (Number.isFinite(x) ? x : err('#NUM!', 'result is not a finite number'));

/** Shared by AND / OR. */
function collectBools(args: Node[], ctx: Ctx, ev: Evaluator): boolean[] | XLError {
  const out: boolean[] = [];
  for (const a of args) {
    const v = ev.eval(a, ctx);
    if (fromRefOrArray(v)) {
      for (const x of toMatrix(v).values()) {
        if (typeof x === 'boolean') out.push(x);
        else if (typeof x === 'number') out.push(x !== 0);
        else if (isError(x)) return x;
      }
    } else {
      const b = toBool(v as Scalar);
      if (isError(b)) return b;
      out.push(b);
    }
  }
  if (out.length === 0) return err('#VALUE!', 'AND/OR got no logical values');
  return out;
}

/** NPV of cash flows at times t (in periods). */
export function npvAt(rate: number, flows: number[], times: number[]): number {
  let s = 0;
  for (let i = 0; i < flows.length; i++) s += flows[i] / Math.pow(1 + rate, times[i]);
  return s;
}

/**
 * IRR by Newton's method starting from `guess`, like Excel. We iterate to full double
 * precision (Excel documents 0.00001% accuracy, but in practice Newton converges
 * quadratically so both land on the same double to ~1e-15).
 */
export function solveIrr(flows: number[], guess: number): number | XLError {
  if (!flows.some((x) => x > 0) || !flows.some((x) => x < 0)) return err('#NUM!', 'IRR needs both positive and negative cash flows');
  let r = guess;
  for (let it = 0; it < 100; it++) {
    let f = 0, df = 0;
    for (let i = 0; i < flows.length; i++) {
      const d = Math.pow(1 + r, i);
      f += flows[i] / d;
      df -= (i * flows[i]) / (d * (1 + r));
    }
    if (df === 0 || !Number.isFinite(f)) return err('#NUM!', 'IRR did not converge');
    const step = f / df;
    r -= step;
    if (r <= -1) return err('#NUM!', 'IRR left the domain r > -1');
    if (Math.abs(step) <= 1e-15 * Math.max(1, Math.abs(r))) return r;
  }
  return err('#NUM!', 'IRR did not converge in 100 iterations');
}

export const FUNCTIONS: Record<string, Fn> = {
  SUM(args, ctx, ev) {
    const xs = collectNumbers(args, ctx, ev);
    return isError(xs) ? xs : checkNum(xs.reduce((a, b) => a + b, 0));
  },
  AVERAGE(args, ctx, ev) {
    const xs = collectNumbers(args.filter((a) => a.t !== 'missing'), ctx, ev);
    if (isError(xs)) return xs;
    if (xs.length === 0) return err('#DIV/0!', 'AVERAGE of no numbers');
    return xs.reduce((a, b) => a + b, 0) / xs.length;
  },
  MIN(args, ctx, ev) {
    const xs = collectNumbers(args, ctx, ev);
    if (isError(xs)) return xs;
    return xs.length ? xs.reduce((a, b) => (b < a ? b : a)) : 0;
  },
  MAX(args, ctx, ev) {
    const xs = collectNumbers(args, ctx, ev);
    if (isError(xs)) return xs;
    return xs.length ? xs.reduce((a, b) => (b > a ? b : a)) : 0;
  },
  SUMPRODUCT(args, ctx, ev) {
    // Arguments are evaluated as arrays (so =SUMPRODUCT((A1:A3>0)*B1:B3) works without Ctrl+Shift+Enter).
    const actx = { ...ctx, array: true };
    const ms = args.map((a) => toMatrix(ev.eval(a, actx)));
    const h = ms[0].height, w = ms[0].width;
    if (ms.some((m) => m.height !== h || m.width !== w)) return err('#VALUE!', 'SUMPRODUCT arrays differ in size');
    let s = 0;
    for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) {
      let p = 1;
      for (const m of ms) {
        const x = m.get(r, c);
        if (isError(x)) return x;
        p *= typeof x === 'number' ? x : 0; // text, booleans and blanks count as 0
      }
      s += p;
    }
    return s;
  },
  IF(args, ctx, ev) {
    if (args.length < 1 || args.length > 3) return err('#VALUE!', 'IF takes 1 to 3 arguments');
    const cv = ev.eval(args[0], ctx);
    const branch = (b: boolean): Value => {
      const n = args[b ? 1 : 2];
      if (n === undefined) return b; // IF(cond, x) with cond FALSE returns FALSE (and TRUE when no x)
      if (n.t === 'missing') return 0;
      return ev.eval(n, ctx);
    };
    if (ctx.array && (cv instanceof Matrix || (cv instanceof RangeRef && !cv.isCell))) {
      const m = toMatrix(cv);
      const t = toMatrix(branch(true)), f = toMatrix(branch(false));
      const pick = (x: Matrix, r: number, c: number) => x.get(Math.min(r, x.height - 1), Math.min(c, x.width - 1));
      return new Matrix(m.rows.map((row, r) => row.map((x, c) => {
        const b = toBool(x);
        return isError(b) ? b : pick(b ? t : f, r, c);
      })));
    }
    const b = toBool(ev.scalar(args[0], ctx));
    if (isError(b)) return b;
    return branch(b);
  },
  AND(args, ctx, ev) { const bs = collectBools(args, ctx, ev); return isError(bs) ? bs : bs.every((x) => x); },
  OR(args, ctx, ev) { const bs = collectBools(args, ctx, ev); return isError(bs) ? bs : bs.some((x) => x); },
  NOT(args, ctx, ev) { const b = toBool(ev.scalar(args[0], ctx)); return isError(b) ? b : !b; },
  NPV(args, ctx, ev) {
    if (args.length < 2) return err('#VALUE!', 'NPV needs a rate and at least one value');
    const rate = num(args[0], ctx, ev);
    if (isError(rate)) return rate;
    const xs = collectNumbers(args.slice(1), ctx, ev);
    if (isError(xs)) return xs;
    if (rate === -1) return err('#DIV/0!', 'NPV rate of -100%');
    return checkNum(npvAt(rate, xs, xs.map((_, i) => i + 1)));
  },
  XNPV(args, ctx, ev) {
    const rate = num(args[0], ctx, ev);
    if (isError(rate)) return rate;
    const vals = [...toMatrix(ev.eval(args[1], ctx)).values()];
    const dates = [...toMatrix(ev.eval(args[2], ctx)).values()];
    for (const x of [...vals, ...dates]) { if (isError(x)) return x; if (typeof x !== 'number') return err('#VALUE!', 'XNPV values/dates must be numbers'); }
    if (vals.length !== dates.length) return err('#NUM!', 'XNPV values and dates differ in length');
    if (rate <= -1) return err('#NUM!', 'XNPV rate must be > -100%');
    const d = (dates as number[]).map(Math.trunc);
    if (d.some((x) => x < d[0])) return err('#NUM!', 'XNPV date precedes the first date');
    return checkNum(npvAt(rate, vals as number[], d.map((x) => (x - d[0]) / 365)));
  },
  IRR(args, ctx, ev) {
    const flows: number[] = [];
    for (const x of toMatrix(ev.eval(args[0], ctx)).values()) {
      if (isError(x)) return x;
      if (typeof x === 'number') flows.push(x);
    }
    const guess = num(args[1], ctx, ev, 0.1);
    if (isError(guess)) return guess;
    return solveIrr(flows, guess);
  },
  ABS(args, ctx, ev) { const x = num(args[0], ctx, ev); return isError(x) ? x : Math.abs(x); },
  ROUND(args, ctx, ev) {
    const x = num(args[0], ctx, ev), d = num(args[1], ctx, ev);
    if (isError(x)) return x; if (isError(d)) return d;
    const f = Math.pow(10, Math.trunc(d));
    // Round half away from zero, nudged to undo binary representation error (Excel shows 2.675 -> 2.68).
    const y = Math.abs(x) * f;
    const r = Math.round(y + y * Number.EPSILON);
    return (Math.sign(x) * r) / f;
  },
  IFERROR(args, ctx, ev) {
    const v = ev.eval(args[0], ctx);
    const s = v instanceof RangeRef || v instanceof Matrix ? ev.scalar(args[0], ctx) : v;
    if (isError(s)) return args[1] && args[1].t !== 'missing' ? ev.eval(args[1], ctx) : 0;
    return v;
  },
  COUNT(args, ctx, ev) {
    let n = 0;
    for (const a of args) {
      const v = ev.eval(a, ctx);
      if (fromRefOrArray(v)) { for (const x of toMatrix(v).values()) if (typeof x === 'number') n++; }
      else if (typeof v === 'number' || typeof v === 'boolean' || (typeof v === 'string' && !isError(toNumber(v)))) n++;
    }
    return n;
  },
  TRUE: () => true,
  FALSE: () => false,
  /** The '@' implicit-intersection operator (stored as _xlfn.SINGLE in files). */
  SINGLE(args, ctx, ev) { return ev.scalar(args[0], ctx); },
};

export const SUPPORTED_FUNCTIONS = Object.keys(FUNCTIONS).filter((f) => f !== 'SINGLE').sort();
