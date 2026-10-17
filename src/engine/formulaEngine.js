/**
 * Minimal Excel-compatible formula engine.
 *
 * Supports everything the Inclement Weather workbook uses, plus a few common
 * extras so the formulas can be extended later without touching the engine:
 *   - numbers, strings, TRUE/FALSE
 *   - + - * / ^ & = <> < > <= >=, unary + and -, postfix %
 *   - A1 / $A$1 references, ranges (A1:B2), sheet-qualified refs ('BOM Extract'!B12)
 *   - defined names (e.g. HoursPerDayRatio)
 *   - SUM, AVERAGE, MIN, MAX, COUNT, ABS, ROUND, IF, IFERROR
 *
 * Formulas are kept as the exact Excel text (e.g. "=+B11+2.5*(B9-B11)/5").
 */

// ---------- errors ----------
export class XlError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
export const isErr = (v) => v instanceof XlError;
const ERR = {
  value: () => new XlError('#VALUE!'),
  div0: () => new XlError('#DIV/0!'),
  ref: () => new XlError('#REF!'),
  name: () => new XlError('#NAME?'),
  circ: () => new XlError('#CIRC!'),
};

// ---------- address helpers ----------
export function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function numToCol(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function parseA1(a1) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a1);
  if (!m) return null;
  return { r: parseInt(m[2], 10), c: colToNum(m[1]) };
}

// ---------- tokenizer ----------
const SHEET_PREFIX = "(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!";
const CELL = '\\$?[A-Za-z]{1,3}\\$?\\d+';
const RE_REF = new RegExp(`^(${SHEET_PREFIX})?(${CELL})(?::(${CELL}))?`);
const RE_NUM = /^(\d+\.?\d*([eE][+-]?\d+)?|\.\d+([eE][+-]?\d+)?)/;
const RE_IDENT = /^[A-Za-z_][A-Za-z0-9_.]*/;

function tokenize(src) {
  const tokens = [];
  let i = 0;
  while (i < src.length) {
    const rest = src.slice(i);
    const ch = rest[0];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let j = 1, out = '';
      while (j < rest.length) {
        if (rest[j] === '"') { if (rest[j + 1] === '"') { out += '"'; j += 2; continue; } break; }
        out += rest[j++];
      }
      tokens.push({ t: 'str', v: out });
      i += j + 1;
      continue;
    }
    // reference (must be tried before identifiers / numbers)
    let m = RE_REF.exec(rest);
    if (m) {
      // A cell-looking token directly followed by "(" is a function name (e.g. LOG10( ) – not used here but be safe
      const after = rest[m[0].length];
      if (after !== '(') {
        tokens.push({ t: 'ref', sheet: m[1] ? m[1].slice(0, -1) : null, a: m[2], b: m[3] || null });
        i += m[0].length;
        continue;
      }
    }
    m = RE_NUM.exec(rest);
    if (m) { tokens.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
    m = RE_IDENT.exec(rest);
    if (m) {
      const name = m[0];
      const isFn = rest[name.length] === '(';
      tokens.push({ t: isFn ? 'fn' : 'name', v: name });
      i += name.length;
      continue;
    }
    const two = rest.slice(0, 2);
    if (two === '<=' || two === '>=' || two === '<>') { tokens.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>(),%'.includes(ch)) { tokens.push({ t: 'op', v: ch }); i++; continue; }
    throw new Error(`Unexpected character "${ch}" in formula "${src}"`);
  }
  return tokens;
}

// ---------- parser (recursive descent -> AST) ----------
function parse(src) {
  const text = src.startsWith('=') ? src.slice(1) : src;
  const tokens = tokenize(text);
  let p = 0;
  const peek = () => tokens[p];
  const eat = (v) => {
    const tk = tokens[p];
    if (tk && tk.t === 'op' && tk.v === v) { p++; return true; }
    return false;
  };

  function parseComparison() {
    let left = parseConcat();
    for (;;) {
      const tk = peek();
      if (tk && tk.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(tk.v)) {
        p++;
        left = { k: 'bin', op: tk.v, l: left, r: parseConcat() };
      } else return left;
    }
  }
  function parseConcat() {
    let left = parseAdd();
    while (eat('&')) left = { k: 'bin', op: '&', l: left, r: parseAdd() };
    return left;
  }
  function parseAdd() {
    let left = parseMul();
    for (;;) {
      if (eat('+')) left = { k: 'bin', op: '+', l: left, r: parseMul() };
      else if (eat('-')) left = { k: 'bin', op: '-', l: left, r: parseMul() };
      else return left;
    }
  }
  function parseMul() {
    let left = parsePow();
    for (;;) {
      if (eat('*')) left = { k: 'bin', op: '*', l: left, r: parsePow() };
      else if (eat('/')) left = { k: 'bin', op: '/', l: left, r: parsePow() };
      else return left;
    }
  }
  function parsePow() {
    let left = parseUnary();
    while (eat('^')) left = { k: 'bin', op: '^', l: left, r: parseUnary() };
    return left;
  }
  function parseUnary() {
    if (eat('+')) return { k: 'un', op: '+', e: parseUnary() };
    if (eat('-')) return { k: 'un', op: '-', e: parseUnary() };
    return parsePostfix();
  }
  function parsePostfix() {
    let e = parsePrimary();
    while (eat('%')) e = { k: 'pct', e };
    return e;
  }
  function parsePrimary() {
    const tk = tokens[p++];
    if (!tk) throw new Error(`Unexpected end of formula "${src}"`);
    if (tk.t === 'num') return { k: 'num', v: tk.v };
    if (tk.t === 'str') return { k: 'str', v: tk.v };
    if (tk.t === 'ref') return { k: 'ref', sheet: tk.sheet, a: tk.a, b: tk.b };
    if (tk.t === 'name') {
      const up = tk.v.toUpperCase();
      if (up === 'TRUE') return { k: 'bool', v: true };
      if (up === 'FALSE') return { k: 'bool', v: false };
      return { k: 'name', v: tk.v };
    }
    if (tk.t === 'fn') {
      if (!eat('(')) throw new Error(`Expected "(" after ${tk.v}`);
      const args = [];
      if (!eat(')')) {
        do { args.push(parseComparison()); } while (eat(','));
        if (!eat(')')) throw new Error(`Expected ")" in "${src}"`);
      }
      return { k: 'fn', name: tk.v.toUpperCase(), args };
    }
    if (tk.t === 'op' && tk.v === '(') {
      const e = parseComparison();
      if (!eat(')')) throw new Error(`Expected ")" in "${src}"`);
      return e;
    }
    throw new Error(`Unexpected token "${tk.v}" in "${src}"`);
  }

  const ast = parseComparison();
  if (p < tokens.length) throw new Error(`Unexpected token "${tokens[p].v}" in "${src}"`);
  return ast;
}

// ---------- coercion ----------
function toNum(v) {
  if (isErr(v)) return v;
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const s = v.trim();
    if (s !== '' && !isNaN(Number(s))) return Number(s);
    return ERR.value();
  }
  return ERR.value();
}
const toStr = (v) => (v === null || v === undefined ? '' : String(v));

// ---------- workbook ----------
/**
 * sheets: { [sheetName]: { cells: Map("r,c" -> { f?: string, v?: any }) } }
 * names:  { [definedName]: "'Sheet'!$A$1" }
 */
export class Workbook {
  constructor(sheets, names = {}) {
    this.sheets = {};
    this.sheetKeys = {};
    for (const [name, sh] of Object.entries(sheets)) {
      this.sheets[name] = sh;
      this.sheetKeys[name.toLowerCase()] = name;
    }
    this.names = names;
    this.cache = new Map();
    this.stack = new Set();
    this.astCache = new Map();
  }

  resolveSheet(name, current) {
    if (!name) return current;
    let n = name;
    if (n.startsWith("'")) n = n.slice(1, -1).replace(/''/g, "'");
    return this.sheetKeys[n.toLowerCase()] || null;
  }

  getValue(sheet, r, c) {
    const sh = this.sheets[sheet];
    if (!sh) return ERR.ref();
    const key = `${sheet}!${r},${c}`;
    if (this.cache.has(key)) return this.cache.get(key);
    const cell = sh.cells.get(`${r},${c}`);
    let val;
    if (!cell) val = null;
    else if (cell.f) {
      if (this.stack.has(key)) return ERR.circ();
      this.stack.add(key);
      try {
        let ast = this.astCache.get(cell.f);
        if (!ast) { ast = parse(cell.f); this.astCache.set(cell.f, ast); }
        val = this.evalNode(ast, sheet);
        if (val && val.range) val = ERR.value();
        if (val === null || val === undefined) val = 0; // a formula pointing at an empty cell shows 0
      } catch (e) {
        val = ERR.name();
      } finally {
        this.stack.delete(key);
      }
    } else val = cell.v === undefined ? null : cell.v;
    this.cache.set(key, val);
    return val;
  }

  rangeValues(rg) {
    const out = [];
    for (let r = Math.min(rg.r1, rg.r2); r <= Math.max(rg.r1, rg.r2); r++)
      for (let c = Math.min(rg.c1, rg.c2); c <= Math.max(rg.c1, rg.c2); c++)
        out.push(this.getValue(rg.sheet, r, c));
    return out;
  }

  evalNode(n, cur) {
    switch (n.k) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'ref': {
        const sheet = this.resolveSheet(n.sheet, cur);
        if (!sheet) return ERR.ref();
        const a = parseA1(n.a);
        if (!n.b) return this.getValue(sheet, a.r, a.c);
        const b = parseA1(n.b);
        return { range: true, sheet, r1: a.r, c1: a.c, r2: b.r, c2: b.c };
      }
      case 'name': {
        const target = this.names[n.v] || this.names[Object.keys(this.names).find((k) => k.toLowerCase() === n.v.toLowerCase())];
        if (!target || target.startsWith('#')) return ERR.name();
        return this.evalNode(parse(target), cur);
      }
      case 'un': {
        const v = this.scalar(this.evalNode(n.e, cur));
        if (n.op === '+') return v;
        const x = toNum(v);
        return isErr(x) ? x : -x;
      }
      case 'pct': {
        const x = toNum(this.scalar(this.evalNode(n.e, cur)));
        return isErr(x) ? x : x / 100;
      }
      case 'bin': return this.evalBin(n, cur);
      case 'fn': return this.evalFn(n, cur);
      default: return ERR.value();
    }
  }

  scalar(v) {
    if (v && v.range) return ERR.value();
    return v;
  }

  evalBin(n, cur) {
    const l = this.scalar(this.evalNode(n.l, cur));
    const r = this.scalar(this.evalNode(n.r, cur));
    if (isErr(l)) return l;
    if (isErr(r)) return r;
    if (n.op === '&') return toStr(l) + toStr(r);
    if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
      const a = l === null ? (typeof r === 'string' ? '' : 0) : l;
      const b = r === null ? (typeof l === 'string' ? '' : 0) : r;
      const cmp = typeof a === 'string' && typeof b === 'string'
        ? a.toLowerCase().localeCompare(b.toLowerCase())
        : (typeof a === 'string' ? 1 : typeof b === 'string' ? -1 : a - b);
      switch (n.op) {
        case '=': return cmp === 0;
        case '<>': return cmp !== 0;
        case '<': return cmp < 0;
        case '>': return cmp > 0;
        case '<=': return cmp <= 0;
        default: return cmp >= 0;
      }
    }
    const a = toNum(l), b = toNum(r);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    switch (n.op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? ERR.div0() : a / b;
      case '^': return Math.pow(a, b);
      default: return ERR.value();
    }
  }

  // collect numbers the way SUM/AVERAGE/MIN/MAX do: text & blanks inside ranges are ignored
  collectNumbers(args, cur) {
    const nums = [];
    for (const a of args) {
      const v = this.evalNode(a, cur);
      if (v && v.range) {
        for (const x of this.rangeValues(v)) {
          if (isErr(x)) return x;
          if (typeof x === 'number') nums.push(x);
        }
      } else {
        if (isErr(v)) return v;
        if (v === null) continue;
        const x = toNum(v);
        if (isErr(x)) return x;
        nums.push(x);
      }
    }
    return nums;
  }

  evalFn(n, cur) {
    const { name, args } = n;
    switch (name) {
      case 'SUM': { const x = this.collectNumbers(args, cur); return isErr(x) ? x : x.reduce((s, v) => s + v, 0); }
      case 'AVERAGE': { const x = this.collectNumbers(args, cur); if (isErr(x)) return x; return x.length ? x.reduce((s, v) => s + v, 0) / x.length : ERR.div0(); }
      case 'MIN': { const x = this.collectNumbers(args, cur); return isErr(x) ? x : (x.length ? Math.min(...x) : 0); }
      case 'MAX': { const x = this.collectNumbers(args, cur); return isErr(x) ? x : (x.length ? Math.max(...x) : 0); }
      case 'COUNT': { const x = this.collectNumbers(args, cur); return isErr(x) ? x : x.length; }
      case 'ABS': { const v = toNum(this.scalar(this.evalNode(args[0], cur))); return isErr(v) ? v : Math.abs(v); }
      case 'ROUND': {
        const v = toNum(this.scalar(this.evalNode(args[0], cur)));
        const d = toNum(this.scalar(this.evalNode(args[1] || { k: 'num', v: 0 }, cur)));
        if (isErr(v)) return v;
        if (isErr(d)) return d;
        const p = Math.pow(10, d);
        return Math.sign(v) * Math.round(Number((Math.abs(v) * p).toPrecision(15))) / p;
      }
      case 'IF': {
        const c = this.scalar(this.evalNode(args[0], cur));
        if (isErr(c)) return c;
        const truthy = typeof c === 'string' ? c.length > 0 : !!c;
        const branch = truthy ? args[1] : args[2];
        return branch ? this.evalNode(branch, cur) : (truthy ? true : false);
      }
      case 'IFERROR': {
        const v = this.evalNode(args[0], cur);
        return isErr(v) ? this.evalNode(args[1], cur) : v;
      }
      default: return ERR.name();
    }
  }
}
