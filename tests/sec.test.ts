// Stage 4 acceptance test: every SEC-generated workbook matches its cached values (1e-9), and the
// implied-growth and LBO solves hit their targets. Full report: npm run stage4.
import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { loadWorkbook, compareToCached } from '../src/index';
import { goalSeek } from '../src/engine/solve';

const dir = join(__dirname, '..', 'workbooks', 'sec');
const cells = existsSync(join(dir, 'cells.json')) ? JSON.parse(readFileSync(join(dir, 'cells.json'), 'utf8')).companies : [];

describe.each(cells.map((c: { ticker: string }) => [c.ticker, c]))('%s', (_t, c: any) => {
  const load = () => { const { wb } = loadWorkbook(new Uint8Array(readFileSync(join(dir, '..', '..', c.file)))); wb.recalc({ cycleStart: 'zero' }); return wb; };
  it('every formula cell matches the cached value to 1e-9, with no unsupported functions', () => {
    const wb = load();
    const res = compareToCached(wb, 1e-9);
    expect([...wb.unsupported.keys()]).toEqual([]);
    expect(res.noCache).toBe(0);
    expect(res.mismatches).toEqual([]);
  });
  it('implied growth hits the market price, and the LBO margin hits a 20% IRR', () => {
    const wb = load();
    const g = goalSeek(wb, c.valuePerShare, c.priceValue, c.growth);
    expect(g.result.converged).toBe(true);
    expect(Math.abs(g.result.value - c.priceValue)).toBeLessThanOrEqual(1e-12 * Math.max(1, c.priceValue));
    const l = goalSeek(wb, c.lbo.irr, c.lbo.target, c.lbo.marginImprovement);
    expect(l.result.converged).toBe(true);
    expect(Math.abs(l.result.value - c.lbo.target)).toBeLessThanOrEqual(1e-12);
  });
});
