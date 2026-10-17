// Value model: number | string | boolean | null (blank) | {err}
export class XlError {
  constructor(err) { this.err = err; }
  toString() { return this.err; }
}
export const ERR = {
  DIV0: new XlError('#DIV/0!'),
  NA: new XlError('#N/A'),
  NAME: new XlError('#NAME?'),
  NUM: new XlError('#NUM!'),
  REF: new XlError('#REF!'),
  VALUE: new XlError('#VALUE!'),
};
export const isError = (v) => v instanceof XlError;

export function errFromString(s) {
  return Object.values(ERR).find((e) => e.err === s) || new XlError(s);
}

export function toNumber(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const t = String(v).trim();
  if (t === '') return ERR.VALUE;
  const n = Number(t.replace(/,/g, ''));
  if (!Number.isNaN(n)) return n;
  const pm = /^(-?[\d.]+)%$/.exec(t);
  if (pm) return parseFloat(pm[1]) / 100;
  return ERR.VALUE;
}

export function toText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (Number.isInteger(v)) return String(v);
    return String(parseFloat(v.toPrecision(15)));
  }
  return String(v);
}

export function toBool(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return v;
  const u = String(v).toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return ERR.VALUE;
}

const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);

// Excel comparison: returns -1/0/1. Blank adopts the other operand's type.
export function compare(a, b) {
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (ra === 1) {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (ra === 2) return a === b ? 0 : a ? 1 : -1;
  // tolerate float noise like Excel (15 significant digits)
  if (a === b || Math.abs(a - b) <= Math.abs(a) * 1e-15) return 0;
  return a < b ? -1 : 1;
}
