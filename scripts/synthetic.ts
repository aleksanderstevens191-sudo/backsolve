// A large synthetic DCF, built in memory (no .xlsx, so no Excel ground truth for its forward values).
// It exists only to show how gradient cost scales with the number of inputs.
//
// S business segments × 10 years. Inputs per segment: base revenue, a growth rate per year and an
// EBIT margin per year. Shared inputs: tax rate, WACC, terminal growth, net debt, shares.

import { Workbook, WorkbookInput } from '../src/index';
import type { CellInput } from '../src/engine/workbook';
import { numToCol } from '../src/engine/refs';

export function syntheticDcf(segments: number, years = 10): { wb: Workbook; output: string; inputCount: number } {
  const inp = new Map<string, CellInput>(), mdl = new Map<string, CellInput>(), val = new Map<string, CellInput>();
  const set = (m: Map<string, CellInput>, a: string, v: number | string) =>
    m.set(a, typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1), value: null } : { value: v });
  // Inputs sheet: row s+1 = segment s. Col A base revenue; cols B.. growth per year; then margins.
  for (let s = 1; s <= segments; s++) {
    set(inp, `A${s}`, 100 + (s % 17) * 7);
    for (let y = 1; y <= years; y++) {
      set(inp, `${numToCol(1 + y)}${s}`, 0.02 + ((s * 7 + y * 3) % 11) / 200);
      set(inp, `${numToCol(1 + years + y)}${s}`, 0.08 + ((s * 5 + y) % 13) / 100);
    }
  }
  const T = segments + 2;
  set(inp, `A${T}`, 0.25); set(inp, `B${T}`, 0.09); set(inp, `C${T}`, 0.025); set(inp, `D${T}`, 500); set(inp, `E${T}`, 100);
  // Model sheet: revenue block rows 1..S, EBIT block rows S+1..2S, totals below.
  for (let s = 1; s <= segments; s++) {
    for (let y = 1; y <= years; y++) {
      const c = numToCol(y), prev = y === 1 ? `Inputs!A${s}` : `${numToCol(y - 1)}${s}`;
      set(mdl, `${c}${s}`, `=${prev}*(1+Inputs!${numToCol(1 + y)}${s})`);
      set(mdl, `${c}${segments + s}`, `=${c}${s}*Inputs!${numToCol(1 + years + y)}${s}`);
    }
  }
  const tot = 2 * segments + 1;
  for (let y = 1; y <= years; y++) {
    const c = numToCol(y);
    set(mdl, `${c}${tot}`, `=SUM(${c}${segments + 1}:${c}${2 * segments})`);            // EBIT
    set(mdl, `${c}${tot + 1}`, `=${c}${tot}-MAX(0,${c}${tot})*Inputs!$A$${T}`);         // NOPAT (no tax on losses)
    set(mdl, `${c}${tot + 2}`, `=${c}${tot + 1}*(1-0.3)`);                              // FCF after reinvestment
  }
  const L = numToCol(years);
  set(val, 'A1', `=NPV(Inputs!B${T},Model!A${tot + 2}:${L}${tot + 2})`);
  set(val, 'A2', `=Model!${L}${tot + 2}*(1+Inputs!C${T})/(Inputs!B${T}-Inputs!C${T})/(1+Inputs!B${T})^${years}`);
  set(val, 'A3', `=(A1+A2-Inputs!D${T})/Inputs!E${T}`);
  const input: WorkbookInput = {
    sheets: [{ name: 'Inputs', cells: inp }, { name: 'Model', cells: mdl }, { name: 'Valuation', cells: val }],
    names: [], calc: { iterate: false, maxIterations: 100, maxChange: 0.001 },
  };
  const wb = new Workbook(input);
  wb.recalc();
  return { wb, output: 'Valuation!A3', inputCount: inp.size };
}
