// Draw docs/stage3_curve.svg from results/stage3_backsolve.json: every (average growth, average
// gross margin) pair that gives the three-statement model a value per share of exactly $30, at WACC
// 8%, 9% and 10%, plus the base case and the closest point to it.
//
//   npx tsx scripts/stage3_chart.ts
//
// Both axes use the same scale (pixels per percentage point), so the shortest move from the base
// case really looks perpendicular to the curve, which is what "closest point" means when each
// percentage point of growth or margin counts the same (σ = 1 pp each).

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const root = new URL('..', import.meta.url).pathname;
const R = JSON.parse(readFileSync(join(root, 'results/stage3_backsolve.json'), 'utf8'));

type Pt = { values: [number, number] };
const curveOf = (c: { branches: { points: Pt[] }[] }) => c.branches.flatMap((b) => b.points.map((p) => p.values));
const curves = [
  { label: 'WACC 8%', cls: 's1', pts: curveOf(R.waccPlusMinus1pp.curves[0]) },
  { label: 'WACC 9% (base)', cls: 's2', pts: curveOf(R.curves[0]) },
  { label: 'WACC 10%', cls: 's3', pts: curveOf(R.waccPlusMinus1pp.curves[1]) },
];
const sol = R.closestPoint[0].result;
const mean = (xs: number[]) => xs.reduce((a, b) => a + b, 0) / xs.length;
const base: [number, number] = [mean(sol.drivers[0].cells.map((c: { base: number }) => c.base)), mean(sol.drivers[1].cells.map((c: { base: number }) => c.base))];
const best: [number, number] = [mean(sol.drivers[0].cells.map((c: { implied: number }) => c.implied)), mean(sol.drivers[1].cells.map((c: { implied: number }) => c.implied))];

// Plot geometry: equal pixels per percentage point on both axes.
const X0 = -0.04, X1 = 0.26, Y0 = 0.30, Y1 = 0.62, PX = 16 / 0.01;
const L = 72, T = 112, W = (X1 - X0) * PX, H = (Y1 - Y0) * PX, RM = 132, B = 64;
const sx = (x: number) => L + (x - X0) * PX, sy = (y: number) => T + (Y1 - y) * PX;
const f = (n: number) => n.toFixed(1);
const pctLabel = (x: number) => `${Math.round(x * 100)}%`;

let grid = '';
for (let x = -0.0; x <= X1 + 1e-9; x += 0.05) grid += `<line class="grid" x1="${f(sx(x))}" x2="${f(sx(x))}" y1="${T}" y2="${T + H}"/><text class="tick" x="${f(sx(x))}" y="${T + H + 18}" text-anchor="middle">${pctLabel(x)}</text>`;
for (let y = Y0; y <= Y1 + 1e-9; y += 0.05) grid += `<line class="grid" x1="${L}" x2="${L + W}" y1="${f(sy(y))}" y2="${f(sy(y))}"/><text class="tick" x="${L - 8}" y="${f(sy(y) + 4)}" text-anchor="end">${pctLabel(y)}</text>`;

const lines = curves.map((c) => `<polyline class="line ${c.cls}" points="${c.pts.map(([x, y]) => `${f(sx(x))},${f(sy(y))}`).join(' ')}"><title>${c.label}: every point here gives exactly $30 per share</title></polyline>`).join('');
// Direct labels at the right-hand ends (growth = 26%), nudged apart if they would collide.
const ends = curves.map((c) => { const p = c.pts.reduce((a, b) => (b[0] > a[0] ? b : a)); return { c, y: sy(p[1]), x: sx(p[0]) }; }).sort((a, b) => a.y - b.y);
for (let k = 1; k < ends.length; k++) if (ends[k].y - ends[k - 1].y < 15) ends[k].y = ends[k - 1].y + 15;
const labels = ends.map((e) => `<text class="dl" x="${f(e.x + 8)}" y="${f(e.y + 4)}">${e.c.label}</text>`).join('');

const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${L + W + RM} ${T + H + B}" width="${L + W + RM}" height="${T + H + B}" role="img" aria-labelledby="t d">
<title id="t">Every growth and margin pair that gives $30 per share</title>
<desc id="d">Three-statement test model. Curves of average revenue growth versus average gross margin that give a value per share of exactly $30, at WACC 8%, 9% and 10%. The base case (${(100 * base[0]).toFixed(2)}% growth, ${(100 * base[1]).toFixed(2)}% margin) gives $${R.closestPoint[0].result.baseValue.toFixed(2)}. The closest $30 point at 9% WACC is ${(100 * best[0]).toFixed(2)}% growth, ${(100 * best[1]).toFixed(2)}% margin. Data: results/stage3_backsolve.json.</desc>
<style>
  svg { --surface: #fcfcfb; --ink: #0b0b0b; --ink2: #52514e; --muted: #898781; --grid: #e1e0d9; --axis: #c3c2b7; --s1: #86b6ef; --s2: #2a78d6; --s3: #104281; font-family: system-ui, -apple-system, "Segoe UI", sans-serif; }
  @media (prefers-color-scheme: dark) { svg { --surface: #1a1a19; --ink: #f0efec; --ink2: #c3c2b7; --grid: #2c2c2a; --axis: #383835; --s1: #cde2fb; --s2: #86b6ef; --s3: #5598e7; } }
  .bg { fill: var(--surface); }
  .grid { stroke: var(--grid); stroke-width: 1; }
  .axis { stroke: var(--axis); stroke-width: 1; }
  .tick { fill: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }
  .title { fill: var(--ink); font-size: 16px; font-weight: 600; }
  .sub { fill: var(--ink2); font-size: 12px; }
  .lab { fill: var(--ink2); font-size: 12px; }
  .dl { fill: var(--ink2); font-size: 12px; }
  .line { fill: none; stroke-width: 2; stroke-linejoin: round; stroke-linecap: round; }
  .s1 { stroke: var(--s1); } .s2 { stroke: var(--s2); } .s3 { stroke: var(--s3); }
  .sw1 { fill: var(--s1); } .sw2 { fill: var(--s2); } .sw3 { fill: var(--s3); }
  .move { stroke: var(--ink2); stroke-width: 1.5; stroke-dasharray: 4 3; }
  .dot { fill: var(--surface); stroke: var(--ink); stroke-width: 2; }
  .dot2 { fill: var(--s2); stroke: var(--surface); stroke-width: 2; }
</style>
<rect class="bg" width="100%" height="100%"/>
<text class="title" x="${L}" y="28">Every growth and margin pair that gives $30 per share</text>
<text class="sub" x="${L}" y="47">Three-statement test model, other assumptions at the base case. Both axes: 1 percentage point = 16 px.</text>
<g transform="translate(${L},62)">
  <rect class="sw1" x="0" y="0" width="12" height="4" rx="2"/><text class="lab" x="18" y="6">WACC 8%</text>
  <rect class="sw2" x="90" y="0" width="12" height="4" rx="2"/><text class="lab" x="108" y="6">WACC 9% (base)</text>
  <rect class="sw3" x="222" y="0" width="12" height="4" rx="2"/><text class="lab" x="240" y="6">WACC 10%</text>
  <circle class="dot" cx="5" cy="24" r="4"/><text class="lab" x="18" y="28">Base case (${pctLabel2(base[0])} growth, ${pctLabel2(base[1])} margin): $${sol.baseValue.toFixed(2)}</text>
  <circle class="dot2" cx="300" cy="24" r="4.5"/><text class="lab" x="313" y="28">Closest $30 point (${pctLabel2(best[0])} growth, ${pctLabel2(best[1])} margin)</text>
</g>
${grid}
<line class="axis" x1="${L}" x2="${L + W}" y1="${T + H}" y2="${T + H}"/>
<line class="axis" x1="${L}" x2="${L}" y1="${T}" y2="${T + H}"/>
<text class="lab" x="${L + W / 2}" y="${T + H + 42}" text-anchor="middle">Average revenue growth per year, 2026–2030</text>
<text class="lab" transform="translate(18,${T + H / 2}) rotate(-90)" text-anchor="middle">Average gross margin</text>
${lines}
${labels}
<line class="move" x1="${f(sx(base[0]))}" y1="${f(sy(base[1]))}" x2="${f(sx(best[0]))}" y2="${f(sy(best[1]))}"/>
<circle class="dot" cx="${f(sx(base[0]))}" cy="${f(sy(base[1]))}" r="5"><title>Base case: ${(100 * base[0]).toFixed(2)}% growth, ${(100 * base[1]).toFixed(2)}% margin → $${sol.baseValue.toFixed(2)}</title></circle>
<circle class="dot2" cx="${f(sx(best[0]))}" cy="${f(sy(best[1]))}" r="5"><title>Closest $30 point: ${(100 * best[0]).toFixed(2)}% growth, ${(100 * best[1]).toFixed(2)}% margin</title></circle>
</svg>
`;
function pctLabel2(x: number) { return `${(100 * x).toFixed(1)}%`; }
writeFileSync(join(root, 'docs/stage3_curve.svg'), svg);
console.log('wrote docs/stage3_curve.svg');
