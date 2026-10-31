// Stage 1 acceptance test: every formula cell in every test workbook matches its cached value.
import { describe, expect, it } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { loadWorkbook, compareToCached } from '../src/index';
import type { Scalar } from '../src/engine/values';

const dir = join(__dirname, '..', 'workbooks');
const known = JSON.parse(readFileSync(join(dir, 'known_differences.json'), 'utf8'));

describe.each(readdirSync(dir).filter((f) => f.endsWith('.xlsx')))('%s', (file) => {
  it('every formula cell matches the cached value to 1e-9, with no unsupported functions', () => {
    const { wb, report } = loadWorkbook(new Uint8Array(readFileSync(join(dir, file))));
    wb.recalc({ cycleStart: 'zero' });
    const k = report.application?.startsWith('LibreOffice') ? known[file] ?? {} : {};
    const overrides = new Map<string, Scalar>(Object.entries(k).map(([c, e]) => [c, (e as { excel: Scalar }).excel]));
    const res = compareToCached(wb, 1e-9, overrides);
    expect([...wb.unsupported.keys()]).toEqual([]);
    expect(res.noCache).toBe(0);
    expect(res.mismatches).toEqual([]);
    for (const c of wb.cycles) expect(c.converged).toBe(true);
  });
});
