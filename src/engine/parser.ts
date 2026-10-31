// Formula text -> abstract syntax tree (AST).
//
// Excel operator precedence, tightest first:  ':'  then unary minus,  '%',  '^',  '* /',  '+ -',  '&',
// comparisons. Note that unary minus binds tighter than '^', so =-2^2 is 4 in Excel, and '^' is
// left-associative, so =2^3^2 is 64.

import { ErrorCode } from './values';
import { MAX_COL, MAX_ROW, colToNum, numToCol, parseA1, quoteSheet } from './refs';

export type RefNode = {
  t: 'ref'; sheet?: string;
  r1: number; c1: number; r2: number; c2: number;
  abs: [boolean, boolean, boolean, boolean]; // absRow1, absCol1, absRow2, absCol2
  range: boolean; kind?: 'cell' | 'cols' | 'rows';
};

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: ErrorCode }
  | RefNode
  | { t: 'name'; name: string; sheet?: string }
  | { t: 'fn'; name: string; args: Node[] }
  | { t: 'bin'; op: string; l: Node; r: Node }
  | { t: 'neg'; e: Node }
  | { t: 'pos'; e: Node }
  | { t: 'pct'; e: Node }
  | { t: 'arr'; rows: Node[][] }
  | { t: 'missing' };

export class ParseError extends Error {}

type Tok =
  | { k: 'num'; v: number } | { k: 'str'; v: string } | { k: 'bool'; v: boolean } | { k: 'err'; v: ErrorCode }
  | { k: 'ref'; node: RefNode } | { k: 'name'; v: string; sheet?: string } | { k: 'func'; v: string }
  | { k: 'op'; v: string } | { k: 'punct'; v: '(' | ')' | ',' | ';' | '{' | '}' };

const ERRORS: ErrorCode[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];
const WORD = /[A-Za-z0-9_.$\\?]/;

function cellPart(word: string): { row: number; col: number; absRow: boolean; absCol: boolean } | null {
  return parseA1(word);
}

function colPart(word: string): { col: number; abs: boolean } | null {
  const m = /^(\$?)([A-Za-z]{1,3})$/.exec(word);
  if (!m) return null;
  const col = colToNum(m[2]);
  return col <= MAX_COL ? { col, abs: m[1] === '$' } : null;
}

function rowPart(word: string): { row: number; abs: boolean } | null {
  const m = /^(\$?)(\d+)$/.exec(word);
  if (!m) return null;
  const row = Number(m[2]);
  return row >= 1 && row <= MAX_ROW ? { row, abs: m[1] === '$' } : null;
}

function tokenize(src: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const peek = (o = 0) => src[i + o];

  const readWord = () => { const s = i; while (i < src.length && WORD.test(src[i])) i++; return src.slice(s, i); };

  /** After an optional sheet prefix, read a cell / range / name reference. */
  const readRefOrName = (sheet: string | undefined, first: string): Tok => {
    // Range "A1:B2", "A:C" or "1:3".
    if (peek() === ':') {
      const save = i; i++;
      const second = readWord();
      const a = cellPart(first), b = cellPart(second);
      if (a && b) {
        return { k: 'ref', node: { t: 'ref', sheet, r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col), abs: [a.absRow, a.absCol, b.absRow, b.absCol], range: true, kind: 'cell' } };
      }
      const ca = colPart(first), cb = colPart(second);
      if (ca && cb) {
        return { k: 'ref', node: { t: 'ref', sheet, r1: 1, c1: Math.min(ca.col, cb.col), r2: MAX_ROW, c2: Math.max(ca.col, cb.col), abs: [true, ca.abs, true, cb.abs], range: true, kind: 'cols' } };
      }
      const ra = rowPart(first), rb = rowPart(second);
      if (ra && rb) {
        return { k: 'ref', node: { t: 'ref', sheet, r1: Math.min(ra.row, rb.row), c1: 1, r2: Math.max(ra.row, rb.row), c2: MAX_COL, abs: [ra.abs, true, rb.abs, true], range: true, kind: 'rows' } };
      }
      i = save; // not a static range; ':' between non-refs is unsupported and will error in the parser
    }
    const a = cellPart(first);
    if (a) return { k: 'ref', node: { t: 'ref', sheet, r1: a.row, c1: a.col, r2: a.row, c2: a.col, abs: [a.absRow, a.absCol, a.absRow, a.absCol], range: false, kind: 'cell' } };
    if (sheet === undefined && /^(TRUE|FALSE)$/i.test(first)) return { k: 'bool', v: first.toUpperCase() === 'TRUE' };
    if (!/^[A-Za-z_\\][A-Za-z0-9_.\\?]*$/.test(first)) throw new ParseError(`Bad reference or name "${first}"`);
    return { k: 'name', v: first, sheet };
  };

  while (i < src.length) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
    if (ch === '"') {
      let s = ''; i++;
      while (true) {
        if (i >= src.length) throw new ParseError('Unterminated string');
        if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2; continue; } i++; break; }
        s += src[i++];
      }
      toks.push({ k: 'str', v: s }); continue;
    }
    if (ch === '#') {
      const e = ERRORS.find((e) => src.startsWith(e, i) || src.toUpperCase().startsWith(e, i));
      if (!e) throw new ParseError(`Unknown error literal at ${src.slice(i)}`);
      i += e.length; toks.push({ k: 'err', v: e }); continue;
    }
    if (/[0-9.]/.test(ch) && !(ch === '.' && !/[0-9]/.test(peek(1) ?? ''))) {
      // Row range like 3:5 (digits ':' digits) — rare but valid.
      const rr = /^(\d+):(\$?\d+)/.exec(src.slice(i));
      if (rr) { i += rr[1].length; toks.push(readRefOrName(undefined, rr[1])); continue; }
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i))!;
      i += m[0].length; toks.push({ k: 'num', v: Number(m[0]) }); continue;
    }
    if (ch === "'") {
      let s = ''; i++;
      while (true) {
        if (i >= src.length) throw new ParseError('Unterminated sheet name');
        if (src[i] === "'") { if (src[i + 1] === "'") { s += "'"; i += 2; continue; } i++; break; }
        s += src[i++];
      }
      if (src[i] !== '!') throw new ParseError(`Expected ! after '${s}'`);
      i++;
      if (s.includes(':')) throw new ParseError(`3-D references ('${s}'!) are unsupported`);
      if (src[i] === '#') { const e = ERRORS.find((e) => src.startsWith(e, i))!; i += e.length; toks.push({ k: 'err', v: e }); continue; }
      toks.push(readRefOrName(s, readWord())); continue;
    }
    if (WORD.test(ch)) {
      const w = readWord();
      if (src[i] === '(') { toks.push({ k: 'func', v: w.toUpperCase().replace(/^_XLFN\.|^_XLWS\./, '').replace(/^_XLFN\./, '') }); continue; }
      if (src[i] === '!') {
        i++;
        if (src[i] === '#') { const e = ERRORS.find((e) => src.startsWith(e, i))!; i += e.length; toks.push({ k: 'err', v: e }); continue; }
        toks.push(readRefOrName(w, readWord())); continue;
      }
      if (src[i] === '[') throw new ParseError(`Structured table references (${w}[...]) are unsupported`);
      toks.push(readRefOrName(undefined, w)); continue;
    }
    if (ch === '[') throw new ParseError('External workbook references ([1]Sheet!A1) are unsupported');
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') { toks.push({ k: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%'.includes(ch)) { toks.push({ k: 'op', v: ch }); i++; continue; }
    if ('(),;{}'.includes(ch)) { toks.push({ k: 'punct', v: ch as '(' }); i++; continue; }
    if (ch === '@') { toks.push({ k: 'func', v: 'SINGLE' }); i++; continue; } // implicit intersection operator
    throw new ParseError(`Unexpected character "${ch}"`);
  }
  return toks;
}

export function parseFormula(src: string): Node {
  const text = src.startsWith('=') ? src.slice(1) : src;
  const toks = tokenize(text);
  let p = 0;
  const peek = () => toks[p];
  const isOp = (v: string) => { const t = toks[p]; return t?.k === 'op' && t.v === v; };
  const isP = (v: string) => { const t = toks[p]; return t?.k === 'punct' && t.v === v; };
  const expectP = (v: string) => { if (!isP(v)) throw new ParseError(`Expected "${v}" in ${src}`); p++; };

  const parseComparison = (): Node => {
    let l = parseConcat();
    while (peek()?.k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes((peek() as { v: string }).v)) {
      const op = (toks[p++] as { v: string }).v; l = { t: 'bin', op, l, r: parseConcat() };
    }
    return l;
  };
  const parseConcat = (): Node => { let l = parseAdd(); while (isOp('&')) { p++; l = { t: 'bin', op: '&', l, r: parseAdd() }; } return l; };
  const parseAdd = (): Node => {
    let l = parseMul();
    while (isOp('+') || isOp('-')) { const op = (toks[p++] as { v: string }).v; l = { t: 'bin', op, l, r: parseMul() }; }
    return l;
  };
  const parseMul = (): Node => {
    let l = parsePow();
    while (isOp('*') || isOp('/')) { const op = (toks[p++] as { v: string }).v; l = { t: 'bin', op, l, r: parsePow() }; }
    return l;
  };
  const parsePow = (): Node => { let l = parsePct(); while (isOp('^')) { p++; l = { t: 'bin', op: '^', l, r: parsePct() }; } return l; };
  const parsePct = (): Node => { let e = parseUnary(); while (isOp('%')) { p++; e = { t: 'pct', e }; } return e; };
  const parseUnary = (): Node => {
    if (isOp('-')) { p++; return { t: 'neg', e: parseUnary() }; }
    if (isOp('+')) { p++; return { t: 'pos', e: parseUnary() }; }
    return parsePrimary();
  };
  const parseArg = (): Node => (isP(',') || isP(')') ? { t: 'missing' } : parseComparison());
  const parsePrimary = (): Node => {
    const t = toks[p++];
    if (!t) throw new ParseError(`Unexpected end of formula: ${src}`);
    switch (t.k) {
      case 'num': return { t: 'num', v: t.v };
      case 'str': return { t: 'str', v: t.v };
      case 'bool': return { t: 'bool', v: t.v };
      case 'err': return { t: 'err', v: t.v };
      case 'ref': return t.node;
      case 'name': return { t: 'name', name: t.v, sheet: t.sheet };
      case 'func': {
        expectP('(');
        const args: Node[] = [];
        if (!isP(')')) { args.push(parseArg()); while (isP(',')) { p++; args.push(parseArg()); } }
        expectP(')');
        return { t: 'fn', name: t.v, args };
      }
      case 'punct':
        if (t.v === '(') { const e = parseComparison(); expectP(')'); return e; }
        if (t.v === '{') {
          const rows: Node[][] = [[]];
          while (true) {
            let neg = false;
            if (isOp('-')) { p++; neg = true; }
            const c = toks[p++];
            if (!c || !['num', 'str', 'bool', 'err'].includes(c.k)) throw new ParseError('Bad array constant');
            let n: Node = c.k === 'num' ? { t: 'num', v: neg ? -c.v : c.v } : c.k === 'str' ? { t: 'str', v: c.v } : c.k === 'bool' ? { t: 'bool', v: c.v } : { t: 'err', v: (c as { v: ErrorCode }).v };
            rows[rows.length - 1].push(n);
            if (isP(',')) { p++; continue; }
            if (isP(';')) { p++; rows.push([]); continue; }
            expectP('}'); break;
          }
          return { t: 'arr', rows };
        }
    }
    throw new ParseError(`Unexpected token ${JSON.stringify(t)} in ${src}`);
  };

  const node = parseComparison();
  if (p < toks.length) throw new ParseError(`Unexpected trailing token ${JSON.stringify(toks[p])} in ${src}`);
  return node;
}

/** Move relative references by (dr, dc): used to expand Excel "shared formulas". */
export function shiftNode(n: Node, dr: number, dc: number): Node {
  switch (n.t) {
    case 'ref': {
      const [ar1, ac1, ar2, ac2] = n.abs;
      return {
        ...n,
        r1: n.kind === 'cols' || ar1 ? n.r1 : n.r1 + dr, c1: n.kind === 'rows' || ac1 ? n.c1 : n.c1 + dc,
        r2: n.kind === 'cols' || ar2 ? n.r2 : n.r2 + dr, c2: n.kind === 'rows' || ac2 ? n.c2 : n.c2 + dc,
      };
    }
    case 'fn': return { ...n, args: n.args.map((a) => shiftNode(a, dr, dc)) };
    case 'bin': return { ...n, l: shiftNode(n.l, dr, dc), r: shiftNode(n.r, dr, dc) };
    case 'neg': case 'pos': case 'pct': return { ...n, e: shiftNode(n.e, dr, dc) };
    default: return n;
  }
}

/** AST -> formula text (without the leading '='). Used for debugging and reports. */
export function printNode(n: Node): string {
  const pr = (x: Node) => printNode(x);
  switch (n.t) {
    case 'num': return String(n.v);
    case 'str': return `"${n.v.replace(/"/g, '""')}"`;
    case 'bool': return n.v ? 'TRUE' : 'FALSE';
    case 'err': return n.v;
    case 'missing': return '';
    case 'ref': {
      const s = n.sheet !== undefined ? `${quoteSheet(n.sheet)}!` : '';
      const d = (b: boolean) => (b ? '$' : '');
      if (n.kind === 'cols') return `${s}${d(n.abs[1])}${numToCol(n.c1)}:${d(n.abs[3])}${numToCol(n.c2)}`;
      if (n.kind === 'rows') return `${s}${d(n.abs[0])}${n.r1}:${d(n.abs[2])}${n.r2}`;
      const a = `${d(n.abs[1])}${numToCol(n.c1)}${d(n.abs[0])}${n.r1}`;
      return n.range ? `${s}${a}:${d(n.abs[3])}${numToCol(n.c2)}${d(n.abs[2])}${n.r2}` : `${s}${a}`;
    }
    case 'name': return n.sheet ? `${quoteSheet(n.sheet)}!${n.name}` : n.name;
    case 'fn': return `${n.name}(${n.args.map(pr).join(',')})`;
    case 'bin': return `(${pr(n.l)}${n.op}${pr(n.r)})`;
    case 'neg': return `-${pr(n.e)}`;
    case 'pos': return `+${pr(n.e)}`;
    case 'pct': return `${pr(n.e)}%`;
    case 'arr': return `{${n.rows.map((r) => r.map(pr).join(',')).join(';')}}`;
  }
}

/** Every function name used in an AST. */
export function functionsUsed(n: Node, out = new Set<string>()): Set<string> {
  if (n.t === 'fn') { out.add(n.name); n.args.forEach((a) => functionsUsed(a, out)); }
  else if (n.t === 'bin') { functionsUsed(n.l, out); functionsUsed(n.r, out); }
  else if (n.t === 'neg' || n.t === 'pos' || n.t === 'pct') functionsUsed(n.e, out);
  return out;
}
