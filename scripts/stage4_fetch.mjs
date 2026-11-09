#!/usr/bin/env node
// Stage 4: download every public input the SEC DCFs need, exactly as served, and record where and
// when each file came from. Nothing here is a workbook and nothing is uploaded anywhere: it only
// downloads public data into data/stage4/raw/.
//
// Run it on a machine with normal internet access (the cloud container can't reach these hosts):
//
//     SEC_USER_AGENT="Your Name your@email" node scripts/stage4_fetch.mjs
//
// SEC asks automated tools to send a User-Agent naming a contact
// (https://www.sec.gov/os/accessing-edgar-data). Without SEC_USER_AGENT we send a generic one that
// names this repo; if SEC answers 403, set SEC_USER_AGENT and run again.
//
// What it downloads (all gzip-compressed on disk, byte-for-byte what the server sent):
//   sec/company_tickers.json.gz          SEC's ticker -> CIK list (to check the CIKs below)
//   sec/CIK##########.json.gz            SEC EDGAR company facts (XBRL API), one per company
//   prices/<TICKER>.json.gz              daily prices, 10 years (Yahoo Finance chart API)
//   prices/GSPC.json.gz                  S&P 500 index, same source (for beta)
//   rates/DGS10.csv.gz                   10-year US Treasury yield (FRED, Federal Reserve Bank of St. Louis)
//   rates/histimpl.xls.gz                Damodaran's implied equity risk premium history (NYU Stern)
//   manifest.json                        URL, time, HTTP status, size and SHA-256 of each file
//
// Yahoo sometimes refuses scripts; then Stooq's CSV is tried for that ticker and the manifest says so.

import { createHash } from 'node:crypto';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUT = join(ROOT, 'data', 'stage4', 'raw');

// The ten companies. Large, non-financial (a free-cash-flow DCF doesn't fit banks or insurers).
export const COMPANIES = [
  { ticker: 'AAPL', cik: 320193, name: 'Apple' },
  { ticker: 'MSFT', cik: 789019, name: 'Microsoft' },
  { ticker: 'NVDA', cik: 1045810, name: 'NVIDIA' },
  { ticker: 'AMZN', cik: 1018724, name: 'Amazon' },
  { ticker: 'GOOGL', cik: 1652044, name: 'Alphabet' },
  { ticker: 'META', cik: 1326801, name: 'Meta Platforms' },
  { ticker: 'WMT', cik: 104169, name: 'Walmart' },
  { ticker: 'JNJ', cik: 200406, name: 'Johnson & Johnson' },
  { ticker: 'PG', cik: 80424, name: 'Procter & Gamble' },
  { ticker: 'KO', cik: 21344, name: 'Coca-Cola' },
];

const SEC_UA = process.env.SEC_USER_AGENT || 'Backsolve research script (github.com/aleksanderstevens191-sudo/backsolve)';
const WEB_UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0) Backsolve/1.0';
const manifest = { fetchedBy: 'scripts/stage4_fetch.mjs', startedAt: new Date().toISOString(), node: process.version, platform: `${process.platform} ${process.arch}`, files: [] };

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function get(url, headers, tries = 3) {
  let last;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await fetch(url, { headers, redirect: 'follow' });
      const body = Buffer.from(await res.arrayBuffer());
      if (res.ok) return { res, body };
      last = new Error(`HTTP ${res.status} for ${url}: ${body.toString('utf8', 0, 200)}`);
      if (res.status === 403 || res.status === 404) break;
    } catch (e) { last = e; }
    await sleep(1500 * (i + 1));
  }
  throw last;
}

function save(rel, url, { res, body }, note) {
  const path = join(OUT, rel + '.gz');
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, gzipSync(body, { level: 9 }));
  const entry = {
    file: `data/stage4/raw/${rel}.gz`, url, fetchedAt: new Date().toISOString(), status: res.status,
    serverDate: res.headers.get('date'), lastModified: res.headers.get('last-modified'),
    bytes: body.length, sha256: createHash('sha256').update(body).digest('hex'), ...(note ? { note } : {}),
  };
  manifest.files.push(entry);
  console.log(`ok  ${rel}  ${(body.length / 1e6).toFixed(2)} MB`);
}

async function tryAll(rel, attempts) {
  const errors = [];
  for (const [url, headers, note] of attempts) {
    try { save(rel, url, await get(url, headers), note); return true; } catch (e) { errors.push(String(e.message || e)); }
  }
  manifest.files.push({ file: `data/stage4/raw/${rel}.gz`, failed: true, errors });
  console.log(`FAIL ${rel}\n  ${errors.join('\n  ')}`);
  return false;
}

const secH = { 'User-Agent': SEC_UA, 'Accept-Encoding': 'gzip, deflate' };
const webH = { 'User-Agent': WEB_UA, Accept: '*/*' };
const now = Math.floor(Date.now() / 1000);
const tenYearsAgo = now - Math.round(10.2 * 365.25 * 86400);
const yahoo = (sym) => `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(sym)}?period1=${tenYearsAgo}&period2=${now}&interval=1d&events=div%2Csplit&includeAdjustedClose=true`;
const yahoo2 = (sym) => yahoo(sym).replace('query1', 'query2');
const stooq = (sym) => `https://stooq.com/q/d/l/?s=${sym}&i=d`;

let fails = 0;
fails += !(await tryAll('sec/company_tickers.json', [['https://www.sec.gov/files/company_tickers.json', secH]]));
for (const c of COMPANIES) {
  const cik = String(c.cik).padStart(10, '0');
  fails += !(await tryAll(`sec/CIK${cik}.json`, [[`https://data.sec.gov/api/xbrl/companyfacts/CIK${cik}.json`, secH]]));
  await sleep(250); // SEC allows 10 requests/second; stay far below
}
for (const [file, sym, stq] of [...COMPANIES.map((c) => [c.ticker, c.ticker, `${c.ticker.toLowerCase()}.us`]), ['GSPC', '^GSPC', '^spx']]) {
  fails += !(await tryAll(`prices/${file}.json`, [[yahoo(sym), webH], [yahoo2(sym), webH]]))
    && !(await tryAll(`prices/${file}.csv`, [[stooq(stq), webH, 'Stooq fallback (Yahoo refused)']]));
}
fails += !(await tryAll('rates/DGS10.csv', [['https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10', webH]]));
fails += !(await tryAll('rates/histimpl.xls', [
  ['https://pages.stern.nyu.edu/~adamodar/pc/datasets/histimpl.xls', webH],
  ['https://www.stern.nyu.edu/~adamodar/pc/datasets/histimpl.xls', webH],
]));

manifest.finishedAt = new Date().toISOString();
mkdirSync(OUT, { recursive: true });
writeFileSync(join(OUT, 'manifest.json'), JSON.stringify(manifest, null, 1) + '\n');
console.log(fails ? `\n${fails} download(s) failed; see manifest.json` : '\nall downloads ok');
process.exit(fails ? 1 : 0);
