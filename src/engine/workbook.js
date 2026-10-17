// In-memory workbook with an Excel-compatible formula evaluator.
import { addr, parseAddr, parseRange, rangeToString, todaySerial } from './address.js';
import { parseFormula, serialize, mapRefs } from './parser.js';
import { ERR, XlError, isError, errFromString, toNumber, toText, toBool, compare } from './values.js';
import { FUNCTIONS } from './functions.js';

const astCache = new Map();
function getAst(f) {
  let a = astCache.get(f);
  if (!a) {
    a = parseFormula(f);
    astCache.set(f, a);
  }
  return a;
}

const CIRCULAR = new XlError('#CIRC!');

export class Workbook {
  constructor(template) {
    this.styles = template.styles;
    this.theme = template.theme;
    this.sheets = {};
    for (const [name, s] of Object.entries(template.sheets)) {
      const cells = new Map();
      for (const [a, c] of Object.entries(s.cells)) cells.set(a, { ...c });
      this.sheets[name] = {
        ...s,
        name,
        cells,
        cols: { ...s.cols },
        rows: { ...s.rows },
        merges: [...s.merges],
        cf: s.cf.map((r) => ({ ...r })),
        shapes: (s.shapes || []).map((x) => ({ ...x, from: [...x.from], to: [...x.to] })),
      };
    }
    this.names = {};
    for (const [n, ref] of Object.entries(template.names || {})) this.names[n.toUpperCase()] = ref;
    this.cache = new Map();
    this.inProgress = new Set();
    this.today = todaySerial();
  }

  // ---------- cell access ----------
  sheet(name) {
    const s = this.sheets[name];
    if (!s) throw new Error(`No sheet ${name}`);
    return s;
  }

  getCell(sheet, a) {
    return this.sheet(sheet).cells.get(a);
  }

  setCell(sheet, a, data) {
    const s = this.sheet(sheet);
    const prev = s.cells.get(a) || {};
    const next = { s: prev.s, note: prev.note, ...data };
    if (next.f === undefined) delete next.f;
    if (next.v === undefined) delete next.v;
    delete next.cv;
    s.cells.set(a, next);
    const { row, col } = parseAddr(a);
    if (row > s.maxRow) s.maxRow = row;
    if (col > s.maxCol) s.maxCol = col;
    this.invalidate();
  }

  // Raw user input: "=..." formula, numeric text -> number, '' -> blank
  setInput(sheet, a, raw) {
    if (raw === null || raw === undefined || raw === '') {
      this.setCell(sheet, a, {});
      return;
    }
    if (typeof raw === 'number' || typeof raw === 'boolean') {
      this.setCell(sheet, a, { v: raw });
      return;
    }
    const str = String(raw);
    if (str.startsWith('=') && str.length > 1) {
      parseFormula(str); // throws on syntax error
      this.setCell(sheet, a, { f: str });
      return;
    }
    const n = Number(str.replace(/,/g, '').trim());
    if (str.trim() !== '' && !Number.isNaN(n)) this.setCell(sheet, a, { v: n });
    else if (/^\s*-?[\d.]+%\s*$/.test(str)) this.setCell(sheet, a, { v: parseFloat(str) / 100 });
    else this.setCell(sheet, a, { v: str });
  }

  invalidate() {
    this.cache.clear();
  }

  getValue(sheet, a) {
    const key = `${sheet}!${a}`;
    if (this.cache.has(key)) return this.cache.get(key);
    const s = this.sheets[sheet];
    if (!s) return ERR.REF;
    const c = s.cells.get(a);
    if (!c) return null;
    if (c.f === undefined) return c.v === undefined ? null : c.v;
    if (this.inProgress.has(key)) return CIRCULAR;
    this.inProgress.add(key);
    let v;
    try {
      v = this.evalNode(getAst(c.f), { sheet });
      if (v && v.range) v = this.rangeTopLeft(v, sheet, a);
      if (v === null || v === undefined) v = 0;
    } catch (e) {
      console.error(`Error evaluating ${key} ${c.f}`, e);
      v = ERR.VALUE;
    }
    this.inProgress.delete(key);
    this.cache.set(key, v);
    return v;
  }

  // Implicit intersection for a range used as a scalar
  rangeTopLeft(rv, sheet, a) {
    const { row, col } = parseAddr(a);
    if (rv.r1 === rv.r2 && col >= rv.c1 && col <= rv.c2) return this.getValue(rv.sheet, addr(rv.r1, col));
    if (rv.c1 === rv.c2 && row >= rv.r1 && row <= rv.r2) return this.getValue(rv.sheet, addr(row, rv.c1));
    if (rv.r1 === rv.r2 && rv.c1 === rv.c2) return this.getValue(rv.sheet, addr(rv.r1, rv.c1));
    return ERR.VALUE;
  }

  rangeValues(rv) {
    const out = [];
    for (let r = rv.r1; r <= rv.r2; r++) {
      const row = [];
      for (let c = rv.c1; c <= rv.c2; c++) row.push(this.getValue(rv.sheet, addr(r, c)));
      out.push(row);
    }
    return out;
  }

  // ---------- evaluator ----------
  evalNode(n, ctx) {
    switch (n.type) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return errFromString(n.v);
      case 'missing': return null;
      case 'paren': return this.evalNode(n.e, ctx);
      case 'ref': {
        if (n.invalid || n.row < 1 || n.col < 1) return ERR.REF;
        return this.getValue(n.sheet || ctx.sheet, addr(n.row, n.col));
      }
      case 'range': {
        const sheet = n.sheet || ctx.sheet;
        return {
          range: true, sheet,
          r1: Math.min(n.a.row, n.b.row), c1: Math.min(n.a.col, n.b.col),
          r2: Math.max(n.a.row, n.b.row), c2: Math.max(n.a.col, n.b.col),
        };
      }
      case 'name': {
        const ref = this.names[n.name.toUpperCase()];
        if (!ref) return ERR.NAME;
        return this.scalar(this.evalNode(getAst(ref), ctx), ctx);
      }
      case 'unary': {
        const v = toNumber(this.scalar(this.evalNode(n.arg, ctx), ctx));
        if (isError(v)) return v;
        return n.op === '-' ? -v : v;
      }
      case 'pct': {
        const v = toNumber(this.scalar(this.evalNode(n.arg, ctx), ctx));
        return isError(v) ? v : v / 100;
      }
      case 'bin': return this.evalBinary(n, ctx);
      case 'func': {
        const fn = FUNCTIONS[n.name];
        if (!fn) return ERR.NAME;
        return fn(n.args, ctx, this);
      }
      default:
        return ERR.VALUE;
    }
  }

  scalar(v, ctx) {
    if (v && v.range) {
      if (v.r1 === v.r2 && v.c1 === v.c2) return this.getValue(v.sheet, addr(v.r1, v.c1));
      return ERR.VALUE;
    }
    return v;
  }

  evalBinary(n, ctx) {
    const a = this.scalar(this.evalNode(n.l, ctx), ctx);
    const b = this.scalar(this.evalNode(n.r, ctx), ctx);
    if (isError(a)) return a;
    if (isError(b)) return b;
    switch (n.op) {
      case '&': return toText(a) + toText(b);
      case '=': return compare(a, b) === 0;
      case '<>': return compare(a, b) !== 0;
      case '<': return compare(a, b) < 0;
      case '>': return compare(a, b) > 0;
      case '<=': return compare(a, b) <= 0;
      case '>=': return compare(a, b) >= 0;
      default: break;
    }
    const x = toNumber(a);
    const y = toNumber(b);
    if (isError(x)) return x;
    if (isError(y)) return y;
    switch (n.op) {
      case '+': return x + y;
      case '-': return x - y;
      case '*': return x * y;
      case '/': return y === 0 ? ERR.DIV0 : x / y;
      case '^': return x ** y;
      default: return ERR.VALUE;
    }
  }

  // Evaluate an arbitrary formula AST in a sheet context (used for conditional formats)
  evalAst(sheet, ast) {
    const v = this.evalNode(ast, { sheet });
    return v && v.range ? this.scalar(v, { sheet }) : v;
  }

  // ---------- structural edits ----------
  // Insert `count` blank rows before row `before` in `sheetName`, updating every reference
  // in the workbook like Excel does (including HLOOKUP row indexes that point past the insert).
  insertRows(sheetName, before, count) {
    const s = this.sheet(sheetName);
    const moved = new Map();
    for (const [a, c] of s.cells) {
      const { row, col } = parseAddr(a);
      moved.set(row >= before ? addr(row + count, col) : a, c);
    }
    s.cells = moved;
    const rows = {};
    for (const [r, d] of Object.entries(s.rows)) {
      const ri = Number(r);
      rows[ri >= before ? ri + count : ri] = d;
    }
    s.rows = rows;
    s.maxRow += count;
    const shiftRange = (str) => str.split(' ').map((part) => {
      const g = parseRange(part);
      if (g.r1 >= before) g.r1 += count;
      if (g.r2 >= before) g.r2 += count;
      return rangeToString(g);
    }).join(' ');
    s.merges = s.merges.map(shiftRange);
    for (const rule of s.cf) {
      rule.sqref = shiftRange(rule.sqref);
    }
    for (const sh of s.shapes) {
      if (sh.from[2] + 1 >= before) sh.from[2] += count;
      if (sh.to[2] + 1 >= before) sh.to[2] += count;
    }

    const shiftRef = (inSheet) => (r) => {
      if ((r.sheet || inSheet) !== sheetName) return r;
      return r.row >= before ? { ...r, row: r.row + count } : r;
    };
    const transform = (ast, inSheet) => {
      const fix = (n) => {
        if (n.type === 'func') {
          let args = n.args.map(fix);
          if ((n.name === 'HLOOKUP') && n.args[1]?.type === 'range' && n.args[2]?.type === 'num') {
            const rg = n.args[1];
            if ((rg.sheet || inSheet) === sheetName) {
              const top = Math.min(rg.a.row, rg.b.row);
              const target = top + n.args[2].v - 1;
              if (top < before && target >= before) {
                args = [...args];
                args[2] = { type: 'num', v: n.args[2].v + count };
              }
            }
          }
          return { ...n, args };
        }
        if (n.type === 'paren') return { ...n, e: fix(n.e) };
        if (n.type === 'unary' || n.type === 'pct') return { ...n, arg: fix(n.arg) };
        if (n.type === 'bin') return { ...n, l: fix(n.l), r: fix(n.r) };
        return n;
      };
      return mapRefs(fix(ast), shiftRef(inSheet));
    };
    for (const sh of Object.values(this.sheets)) {
      for (const c of sh.cells.values()) {
        if (c.f !== undefined) c.f = '=' + serialize(transform(getAst(c.f), sh.name));
      }
      for (const rule of sh.cf) {
        if (rule.formula && rule.type === 'expression' && sh.name === sheetName) {
          rule.formula = rule.formula.map((f) => serialize(transform(getAst(f), sh.name)));
        }
      }
    }
    for (const [n, ref] of Object.entries(this.names)) {
      this.names[n] = serialize(transform(getAst(ref), null));
    }
    this.invalidate();
  }

  // Copy style (and optionally nothing else) from one row to another, columns c1..c2
  copyRowStyle(sheetName, fromRow, toRow, c1, c2) {
    const s = this.sheet(sheetName);
    for (let c = c1; c <= c2; c++) {
      const src = s.cells.get(addr(fromRow, c));
      if (src && src.s !== undefined) {
        const dst = s.cells.get(addr(toRow, c)) || {};
        s.cells.set(addr(toRow, c), { ...dst, s: src.s });
      }
    }
    if (s.rows[fromRow]) s.rows[toRow] = { ...s.rows[fromRow] };
  }

  // Snapshot every value on a sheet (for export)
  sheetValues(sheetName) {
    const s = this.sheet(sheetName);
    const out = new Map();
    for (const a of s.cells.keys()) out.set(a, this.getValue(sheetName, a));
    return out;
  }
}

export { toBool, toNumber, toText, isError, ERR };
