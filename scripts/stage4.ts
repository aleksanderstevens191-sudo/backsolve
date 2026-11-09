// Stage 4 checks and results: the ten SEC-generated workbooks (workbooks/sec/*.xlsx).
// Writes results/stage4_sec.json.
//
//   npx tsx scripts/stage4.ts            (npm run stage4)
//
// Per company:
//  1. Forward pass: every formula cell vs the value cached in the file (1e-9), as in Stage 1.
//     The file says which program cached the values (Microsoft Excel or LibreOffice).
//  2. Gradients: every numeric formula cell × every input cell vs finite differences (1e-6), as in
//     Stage 2.
//  3. Implied growth: the one revenue-growth rate (same every year for 10 years) that makes the
//     model's value per share equal the market price, next to the actual 5-year growth from the
//     same SEC data. Repeated with the WACC 1 percentage point lower and higher.
//  4. LBO: the EBITDA-margin improvement (by year 5) that gives a 20% IRR at the offer price, for
//     the default 30% premium and for premiums from 0% to 50%.
//  5. If Excel's own Goal Seek results exist (workbooks/sec/excel/, from scripts/excel_sec.sh on a
//     Mac), our answers are compared with Excel's to 1e-6.

import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { loadWorkbook, compareToCached, DEFAULT_TOL, Workbook } from '../src/index';
import { Backsolver, goalSeek, cellAt } from '../src/engine/solve';
import { checkWorkbook, machine } from './gradcheck';

const root = new URL('..', import.meta.url).pathname;
const GS_TOL = 1e-6;
const cellsFile = JSON.parse(readFileSync(join(root, 'workbooks/sec/cells.json'), 'utf8'));
const data = JSON.parse(readFileSync(join(root, 'data/stage4/companies.json'), 'utf8'));
const only = process.argv.slice(2);

const num = (wb: Workbook, ref: string) => cellAt(wb, ref).value as number;
const pct = (x: number, d = 2) => `${(100 * x).toFixed(d)}%`;
const pp = (x: number, d = 2) => `${x >= 0 ? '+' : ''}${(100 * x).toFixed(d)} pp`;
function setInput(wb: Workbook, ref: string, v: number) { const [s, a] = ref.split('!'); wb.setValue(s, a, v); wb.recalc({ cycleStart: 'current' }); }

const rows: Record<string, unknown>[] = [];
const failures: string[] = [];
for (const c of cellsFile.companies) {
  if (only.length && !only.includes(c.ticker)) continue;
  const bytes = new Uint8Array(readFileSync(join(root, c.file)));
  const { wb, report } = loadWorkbook(bytes);
  wb.recalc({ cycleStart: 'zero' });
  console.log(`\n${c.ticker} ${c.name}  (${c.file}, cached by ${report.application})`);

  // 1. forward pass
  const cmp = compareToCached(wb, DEFAULT_TOL);
  const fwdOk = cmp.mismatches.length === 0 && cmp.noCache === 0 && wb.unsupported.size === 0;
  if (!fwdOk) failures.push(`${c.ticker}: forward check`);
  console.log(`  forward: ${cmp.matched}/${cmp.compared} formula cells match the cached values (worst scaled diff ${cmp.maxScaledDiff.toExponential(1)} at ${cmp.worstCell}), ${cmp.noCache} without a cached value, ${wb.unsupported.size} unsupported functions  ${fwdOk ? 'PASS' : 'FAIL'}`);
  for (const m of cmp.mismatches.slice(0, 5)) console.log(`    MISMATCH ${m.cell} =${m.formula}: ours ${m.ours}, cached ${m.cached}`);

  // 2. gradients
  const gc = checkWorkbook(c.file, wb, c.valuePerShare, true, true);
  const g = gc.result.gradientCheck;
  if (!gc.ok) failures.push(`${c.ticker}: gradient check`);
  console.log(`  gradients: ${g.outputsChecked} outputs × ${g.inputs} inputs = ${g.pairs} pairs: ${g.exactZero} both 0, ${g.passRelative} within 1e-6 (worst ${g.maxRelErrorAmongPassing.toExponential(1)}), ${g.withinFdRoundingNoise} FD noise, ${g.excludedStepCrossesKink} excluded (kink), ${g.failed} FAILED  ${gc.ok ? 'PASS' : 'FAIL'}`);
  wb.recalc({ cycleStart: 'current' });

  // 3. implied growth
  const price = num(wb, c.price);
  const base = { valuePerShare: num(wb, c.valuePerShare), growth: num(wb, c.growth), wacc: num(wb, c.wacc), costOfEquity: num(wb, c.costOfEquity),
    costOfDebt: num(wb, c.costOfDebt), beta: num(wb, c.beta), taxRate: num(wb, c.taxRate), riskFree: num(wb, c.riskFree), erp: num(wb, c.erp) };
  const cagr5 = num(wb, c.cagr5), cagr9 = num(wb, c.cagr9);
  const gs = goalSeek(wb, c.valuePerShare, price, c.growth);
  const implied = gs.answer;
  const hit = gs.result;
  const solver = new Backsolver(wb, c.valuePerShare, [{ name: 'Revenue growth', cells: c.growth }]);
  const slope = solver.sensitivity1D(solver.solve1D(price), c.waccAdjustment).dAnswer; // d(implied growth)/d(WACC), implicit function theorem
  const byWacc: Record<string, number | null> = {};
  for (const dw of [-0.01, 0.01]) {
    setInput(wb, c.waccAdjustment, dw);
    try { byWacc[dw > 0 ? '+1pp' : '-1pp'] = goalSeek(wb, c.valuePerShare, price, c.growth).answer; } catch (e) { byWacc[dw > 0 ? '+1pp' : '-1pp'] = null; }
    setInput(wb, c.waccAdjustment, 0);
  }
  const growthOk = gs.result.converged && Math.abs(hit.value - price) <= 1e-12 * Math.max(1, price);
  if (!growthOk) failures.push(`${c.ticker}: implied growth did not converge`);
  console.log(`  value per share at base case (growth ${pct(base.growth)}): ${base.valuePerShare.toFixed(2)} vs price ${price.toFixed(2)} on ${c.priceDate}; WACC ${pct(base.wacc)} (beta ${base.beta.toFixed(2)})`);
  console.log(`  implied 10-year revenue growth ${pct(implied)} vs actual 5-year ${pct(cagr5)} (gap ${pp(implied - cagr5)}); WACC −1 pp: ${byWacc['-1pp'] == null ? 'n/a' : pct(byWacc['-1pp']!)}, +1 pp: ${byWacc['+1pp'] == null ? 'n/a' : pct(byWacc['+1pp']!)}; slope ${(slope).toFixed(2)} pp per pp`);

  // 4. LBO
  const L = c.lbo;
  const lboBase = { irr: num(wb, L.irr), moic: num(wb, L.moic), entryMultiple: num(wb, L.entryMultiple), ebitdaMargin0: num(wb, L.ebitdaMargin0), premium: num(wb, L.premium), growth: num(wb, L.growth) };
  const lbo = goalSeek(wb, L.irr, L.target, L.marginImprovement);
  const lboOk = lbo.result.converged && Math.abs(lbo.result.value - L.target) <= 1e-12;
  if (!lboOk) failures.push(`${c.ticker}: LBO margin did not converge`);
  const premiums: { premium: number; marginImprovement: number | null; entryMultiple: number }[] = [];
  for (const p of [0, 0.1, 0.2, 0.3, 0.4, 0.5]) {
    setInput(wb, L.premium, p);
    let m: number | null = null;
    try { const r = goalSeek(wb, L.irr, L.target, L.marginImprovement); m = r.result.converged ? r.answer : null; } catch { m = null; }
    premiums.push({ premium: p, marginImprovement: m, entryMultiple: num(wb, L.entryMultiple) });
  }
  setInput(wb, L.premium, lboBase.premium);
  console.log(`  LBO at ${pct(lboBase.premium, 0)} premium (entry ${lboBase.entryMultiple.toFixed(1)}x EBITDA): base IRR ${pct(lboBase.irr)}; 20% IRR needs EBITDA margin ${pp(lbo.answer)} by year 5 (${pct(lboBase.ebitdaMargin0, 1)} → ${pct(lboBase.ebitdaMargin0 + lbo.answer, 1)})`);
  console.log(`    by premium: ${premiums.map((p) => `${pct(p.premium, 0)}: ${p.marginImprovement == null ? 'n/a' : pp(p.marginImprovement, 1)}`).join(', ')}`);

  // 5. Excel's Goal Seek, if run
  const excel: Record<string, unknown> = {};
  for (const [kind, setRef, changing, goal, ours] of [['growth', c.valuePerShare, c.growth, price, implied], ['lbo', L.irr, L.marginImprovement, L.target, lbo.answer]] as const) {
    const f = join(root, 'workbooks/sec/excel', `${c.ticker}_${kind}_excel.xlsx`);
    if (!existsSync(f)) { excel[kind] = { status: 'not run' }; continue; }
    const { wb: ewb, report: er } = loadWorkbook(new Uint8Array(readFileSync(f)));
    const xa = cellAt(ewb, changing).value as number, xo = cellAt(ewb, setRef).cached as number;
    ewb.recalc({ cycleStart: 'zero' });
    const fwd = compareToCached(ewb, DEFAULT_TOL);
    const diff = (ours as number) - xa;
    const ok = Math.abs(diff) <= GS_TOL * Math.max(1, Math.abs(xa)) && fwd.mismatches.length === 0 && fwd.noCache === 0;
    if (!ok) failures.push(`${c.ticker}: Excel Goal Seek (${kind})`);
    excel[kind] = { savedBy: er.application, excelAnswer: xa, excelOutputAtItsAnswer: xo, excelMissBy: xo - (goal as number), diffOursMinusExcel: diff, pass: ok,
      forwardCheckAtExcelsAnswer: { compared: fwd.compared, matched: fwd.matched, maxScaledDiff: fwd.maxScaledDiff } };
    console.log(`  Excel Goal Seek (${kind}): Excel ${xa}, ours ${ours}, diff ${diff.toExponential(2)}  ${ok ? 'PASS' : 'FAIL'}`);
  }

  const co = data.companies.find((x: { ticker: string }) => x.ticker === c.ticker);
  rows.push({
    ticker: c.ticker, name: c.name, file: c.file, cachedBy: report.application, fiscalYears: c.fiscalYears,
    price: { value: price, date: c.priceDate, source: co.price.source, url: co.price.url, note: co.price.note },
    shares: co.shares,
    forward: { ok: fwdOk, formulaCells: cmp.formulaCells, compared: cmp.compared, matched: cmp.matched, noCachedValue: cmp.noCache, maxScaledDiff: cmp.maxScaledDiff, worstCell: cmp.worstCell, mismatches: cmp.mismatches.slice(0, 20), tolerance: `|ours − cached| ≤ ${DEFAULT_TOL} × max(1, |cached|)` },
    gradients: { ok: gc.ok, ...g, failures: g.failures.slice(0, 10), excluded: g.excluded.slice(0, 10), noiseSamples: undefined, benchmark: gc.result.benchmark, kinksUpstreamOfValuePerShare: gc.result.kinksUpstreamOfHeadline },
    baseCase: base,
    impliedGrowth: {
      definition: 'the one revenue growth rate, the same in each of the next 10 years, at which the model value per share equals the market price (everything else held at the base case)',
      implied, actual5yCagr: cagr5, actual9yCagr: cagr9, gapVs5y: implied - cagr5, outputAtAnswer: hit.value, missBy: hit.value - price, iterations: gs.result.iterations,
      waccMinus1pp: byWacc['-1pp'], waccPlus1pp: byWacc['+1pp'], slopePerWacc: slope,
    },
    lbo: { ...lboBase, targetIrr: L.target, marginImprovementNeeded: lbo.answer, ebitdaMarginAtYear5: lboBase.ebitdaMargin0 + lbo.answer, irrAtAnswer: lbo.result.value, missBy: lbo.result.value - L.target, byPremium: premiums },
    excelGoalSeek: excel,
  });
}

const excelRan = rows.some((r) => Object.values(r.excelGoalSeek as Record<string, { status?: string }>).some((x) => !x.status));
const excelCached = rows.filter((r) => String(r.cachedBy).startsWith('Microsoft')).length;
const overall = failures.length ? 'FAIL' : `PASS${excelCached < rows.length ? ` (${rows.length - excelCached} of ${rows.length} workbooks still cached by LibreOffice, not Excel)` : ''}${excelRan ? '' : ' (Excel Goal Seek comparison not run)'}`;
mkdirSync(join(root, 'results'), { recursive: true });
if (!only.length) writeFileSync(join(root, 'results/stage4_sec.json'), JSON.stringify({
  generatedAt: new Date().toISOString(), machine, overall, failures,
  priceDate: data.priceDate, riskFreeRate: data.riskFreeRate, equityRiskPremium: data.equityRiskPremium, assumptions: cellsFile.assumptions,
  companies: rows,
}, null, 1) + '\n');
console.log(`\n${failures.length ? `FAILURES: ${failures.join('; ')}\n` : ''}${only.length ? '' : 'wrote results/stage4_sec.json — '}overall ${overall}`);
process.exit(failures.length ? 1 : 0);
