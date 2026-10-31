// Compare every formula cell the engine computed against the value cached in the file.

import { Workbook } from './workbook';
import { Scalar, XLError } from './values';
import { fullAddr } from './refs';

/** Numbers match when |ours - cached| <= tol * max(1, |cached|): relative for big numbers, absolute near zero. */
export const DEFAULT_TOL = 1e-9;

export interface Mismatch { cell: string; formula?: string; ours: string; cached: string; diff?: number }
export interface CompareResult {
  formulaCells: number; compared: number; matched: number; noCache: number;
  maxAbsDiff: number; maxScaledDiff: number; worstCell?: string;
  mismatches: Mismatch[];
}

const show = (v: Scalar) => (v instanceof XLError ? v.code : JSON.stringify(v));

/**
 * `overrides` replaces the cached value for specific cells (e.g. where the file was cached by
 * LibreOffice and Excel's documented rule gives a different answer).
 */
export function compareToCached(wb: Workbook, tol = DEFAULT_TOL, overrides: Map<string, Scalar> = new Map()): CompareResult {
  const res: CompareResult = { formulaCells: 0, compared: 0, matched: 0, noCache: 0, maxAbsDiff: 0, maxScaledDiff: 0, mismatches: [] };
  for (const c of wb.allFormulaCells()) {
    res.formulaCells++;
    const addr = fullAddr(c.sheet, c.row, c.col);
    if (c.cached === null && !c.arrayMaster) { res.noCache++; continue; }
    res.compared++;
    const a = c.value, b = overrides.has(addr) ? overrides.get(addr)! : c.cached;
    let ok: boolean; let diff: number | undefined;
    if (typeof a === 'number' && typeof b === 'number') {
      diff = Math.abs(a - b);
      const scaled = diff / Math.max(1, Math.abs(b));
      if (diff > res.maxAbsDiff) res.maxAbsDiff = diff;
      if (scaled > res.maxScaledDiff) { res.maxScaledDiff = scaled; res.worstCell = addr; }
      ok = scaled <= tol;
    } else if (a instanceof XLError && b instanceof XLError) ok = a.code === b.code;
    else if (a === null && b === '') ok = true; // empty-text result vs blank
    else ok = a === b || (a === null && b === 0) || (typeof a === 'string' && typeof b === 'string' && a === b);
    if (ok) res.matched++;
    else res.mismatches.push({ cell: addr, formula: c.formula, ours: show(a), cached: show(b), diff });
  }
  return res;
}
