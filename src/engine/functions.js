// Excel worksheet functions used by the Inclement Weather template.
import { ERR, isError, toNumber, toText, toBool, compare } from './values.js';
import { serialToDate, dateToSerial } from './address.js';
import { formatValue } from './format.js';

// Evaluate args eagerly; ranges stay as range objects.
const evalArgs = (args, ctx, wb) => args.map((a) => wb.evalNode(a, ctx));

// Flatten numbers for SUM/MAX: range cells ignore text/bools/blanks; direct args are coerced.
function collectNumbers(vals, wb) {
  const out = [];
  for (const v of vals) {
    if (v && v.range) {
      for (const row of wb.rangeValues(v)) {
        for (const x of row) {
          if (isError(x)) return x;
          if (typeof x === 'number') out.push(x);
        }
      }
    } else {
      if (v === null) continue;
      const n = toNumber(v);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function collectBools(vals, wb) {
  const out = [];
  for (const v of vals) {
    if (v && v.range) {
      for (const row of wb.rangeValues(v)) {
        for (const x of row) {
          if (isError(x)) return x;
          if (typeof x === 'number' || typeof x === 'boolean') out.push(Boolean(x));
        }
      }
    } else {
      if (v === null) continue;
      const b = toBool(v);
      if (isError(b)) return b;
      out.push(b);
    }
  }
  return out;
}

function scalarArg(a, ctx, wb) {
  return wb.scalar(wb.evalNode(a, ctx), ctx);
}

function numArg(a, ctx, wb) {
  const v = scalarArg(a, ctx, wb);
  if (isError(v)) return v;
  return toNumber(v);
}

function dateParts(a, ctx, wb) {
  const n = numArg(a, ctx, wb);
  if (isError(n)) return n;
  if (n < 0) return ERR.NUM;
  if (Math.floor(n) === 0) return { y: 1900, m: 1, d: 0 };
  const d = serialToDate(n);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}

function matchCriteria(crit) {
  let op = '=';
  let rhs = crit;
  if (typeof crit === 'string') {
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(crit);
    op = m[1] || '=';
    rhs = m[2];
    const n = Number(rhs);
    if (rhs.trim() !== '' && !Number.isNaN(n)) rhs = n;
  }
  return (v) => {
    if (op === '=' && typeof rhs === 'string' && rhs === '') return v === null || v === '';
    if (v === null) return op === '<>' ? true : false;
    if (typeof rhs === 'number' && typeof v !== 'number') return op === '<>';
    if (typeof rhs === 'string' && typeof v !== 'string') return op === '<>';
    const c = compare(v, rhs);
    switch (op) {
      case '=': return c === 0;
      case '<>': return c !== 0;
      case '<': return c < 0;
      case '>': return c > 0;
      case '<=': return c <= 0;
      case '>=': return c >= 0;
      default: return false;
    }
  };
}

export const FUNCTIONS = {
  IF(args, ctx, wb) {
    const c = toBool(scalarArg(args[0], ctx, wb));
    if (isError(c)) return c;
    if (c) return args.length > 1 ? scalarArg(args[1], ctx, wb) ?? 0 : true;
    return args.length > 2 ? scalarArg(args[2], ctx, wb) ?? 0 : false;
  },
  AND(args, ctx, wb) {
    const b = collectBools(evalArgs(args, ctx, wb), wb);
    if (isError(b)) return b;
    if (!b.length) return ERR.VALUE;
    return b.every(Boolean);
  },
  OR(args, ctx, wb) {
    const b = collectBools(evalArgs(args, ctx, wb), wb);
    if (isError(b)) return b;
    if (!b.length) return ERR.VALUE;
    return b.some(Boolean);
  },
  NOT(args, ctx, wb) {
    const b = toBool(scalarArg(args[0], ctx, wb));
    return isError(b) ? b : !b;
  },
  SUM(args, ctx, wb) {
    const n = collectNumbers(evalArgs(args, ctx, wb), wb);
    if (isError(n)) return n;
    return n.reduce((s, x) => s + x, 0);
  },
  MAX(args, ctx, wb) {
    const n = collectNumbers(evalArgs(args, ctx, wb), wb);
    if (isError(n)) return n;
    return n.length ? Math.max(...n) : 0;
  },
  MIN(args, ctx, wb) {
    const n = collectNumbers(evalArgs(args, ctx, wb), wb);
    if (isError(n)) return n;
    return n.length ? Math.min(...n) : 0;
  },
  AVERAGE(args, ctx, wb) {
    const n = collectNumbers(evalArgs(args, ctx, wb), wb);
    if (isError(n)) return n;
    return n.length ? n.reduce((s, x) => s + x, 0) / n.length : ERR.DIV0;
  },
  INT(args, ctx, wb) {
    const n = numArg(args[0], ctx, wb);
    return isError(n) ? n : Math.floor(n);
  },
  ROUND(args, ctx, wb) {
    const n = numArg(args[0], ctx, wb);
    const d = args[1] ? numArg(args[1], ctx, wb) : 0;
    if (isError(n)) return n;
    if (isError(d)) return d;
    const f = 10 ** Math.trunc(d);
    const r = Math.round(Math.abs(n) * f * (1 + 1e-15)) / f;
    return n < 0 ? -r : r;
  },
  ABS(args, ctx, wb) {
    const n = numArg(args[0], ctx, wb);
    return isError(n) ? n : Math.abs(n);
  },
  TODAY(args, ctx, wb) {
    return wb.today;
  },
  DATE(args, ctx, wb) {
    const [y, m, d] = args.map((a) => numArg(a, ctx, wb));
    for (const x of [y, m, d]) if (isError(x)) return x;
    let yy = Math.trunc(y);
    if (yy < 1900) yy += 1900;
    const s = dateToSerial(yy, Math.trunc(m) - 1, Math.trunc(d));
    return s < 0 ? ERR.NUM : s;
  },
  YEAR(args, ctx, wb) {
    const p = dateParts(args[0], ctx, wb);
    return isError(p) ? p : p.y;
  },
  MONTH(args, ctx, wb) {
    const p = dateParts(args[0], ctx, wb);
    return isError(p) ? p : p.m;
  },
  DAY(args, ctx, wb) {
    const p = dateParts(args[0], ctx, wb);
    return isError(p) ? p : p.d;
  },
  CONCATENATE(args, ctx, wb) {
    let s = '';
    for (const a of args) {
      const v = scalarArg(a, ctx, wb);
      if (isError(v)) return v;
      s += toText(v);
    }
    return s;
  },
  TEXT(args, ctx, wb) {
    const v = scalarArg(args[0], ctx, wb);
    const f = scalarArg(args[1], ctx, wb);
    if (isError(v)) return v;
    if (isError(f)) return f;
    const n = typeof v === 'string' ? toNumber(v) : v;
    return formatValue(isError(n) ? v : n ?? 0, toText(f));
  },
  COUNTIF(args, ctx, wb) {
    const rv = wb.evalNode(args[0].type === 'paren' ? args[0].e : args[0], ctx);
    const crit = scalarArg(args[1], ctx, wb);
    if (!rv || !rv.range) return ERR.VALUE;
    const test = matchCriteria(crit);
    let count = 0;
    for (const row of wb.rangeValues(rv)) for (const x of row) if (test(x)) count++;
    return count;
  },
  HLOOKUP(args, ctx, wb) {
    const x = scalarArg(args[0], ctx, wb);
    if (isError(x)) return x;
    const rv = wb.evalNode(args[1], ctx);
    if (!rv || !rv.range) return ERR.VALUE;
    const idx = numArg(args[2], ctx, wb);
    if (isError(idx)) return idx;
    const approx = args.length < 4 ? true : toBool(scalarArg(args[3], ctx, wb));
    const rowIdx = Math.trunc(idx);
    if (rowIdx < 1) return ERR.VALUE;
    if (rowIdx > rv.r2 - rv.r1 + 1) return ERR.REF;
    const width = rv.c2 - rv.c1 + 1;
    let hit = -1;
    for (let i = 0; i < width; i++) {
      const v = wb.getValue(rv.sheet, colAddr(rv.r1, rv.c1 + i));
      if (approx) {
        if (v === null || isError(v)) continue;
        if (typeof v !== typeof x && !(v === null)) continue;
        const c = compare(v, x);
        if (c <= 0) hit = i;
        else break;
      } else if (compare(v, x) === 0 && typeof v === typeof x) {
        hit = i;
        break;
      }
    }
    if (hit < 0) return ERR.NA;
    const out = wb.getValue(rv.sheet, colAddr(rv.r1 + rowIdx - 1, rv.c1 + hit));
    return out === null ? 0 : out;
  },
};

import { addr as colAddr } from './address.js';
