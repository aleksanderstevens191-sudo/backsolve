// Excel's rules for turning one kind of value into another.

import { Matrix, RangeRef, Scalar, Value, XLError } from './values';

export function toNumber(v: Scalar): number | XLError {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof XLError) return v;
  const s = v.trim();
  if (s === '') return new XLError('#VALUE!', 'empty text is not a number');
  const pct = s.endsWith('%');
  const n = Number((pct ? s.slice(0, -1) : s).replace(/,/g, ''));
  if (!Number.isFinite(n)) return new XLError('#VALUE!', `"${v}" is not a number`);
  return pct ? n / 100 : n;
}

export function toBool(v: Scalar): boolean | XLError {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v instanceof XLError) return v;
  const s = v.toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  return new XLError('#VALUE!', `"${v}" is not TRUE/FALSE`);
}

/** Number -> text the way Excel's General format does it (up to 15 significant digits). */
export function numberToText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  let s = n.toPrecision(15);
  if (s.includes('e')) {
    let [m, e] = s.split('e');
    if (m.includes('.')) m = m.replace(/0+$/, '').replace(/\.$/, '');
    const ev = Number(e);
    return `${m}E${ev < 0 ? '-' : '+'}${String(Math.abs(ev)).padStart(2, '0')}`;
  }
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

export function toText(v: Scalar): string | XLError {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof XLError) return v;
  return numberToText(v);
}

/** Excel ordering for comparisons: numbers < text < booleans; text compares case-insensitively. */
export function compare(a: Scalar, b: Scalar): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const rank = (x: Scalar) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2);
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number') return a < (b as number) ? -1 : a > (b as number) ? 1 : 0;
  if (typeof a === 'string') { const x = a.toUpperCase(), y = (b as string).toUpperCase(); return x < y ? -1 : x > y ? 1 : 0; }
  return a === b ? 0 : a ? 1 : -1;
}

export function toMatrix(v: Value): Matrix {
  if (v instanceof Matrix) return v;
  if (v instanceof RangeRef) return v.toMatrix();
  return new Matrix([[v]]);
}
