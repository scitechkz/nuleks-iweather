// A1-style address helpers. Rows and columns are 1-based like Excel.

const colCache = [];

export function colToLetters(col) {
  if (colCache[col]) return colCache[col];
  let n = col;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  colCache[col] = s;
  return s;
}

export function lettersToCol(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n;
}

export function addr(row, col) {
  return colToLetters(col) + row;
}

const ADDR_RE = /^\$?([A-Z]{1,3})\$?(\d+)$/;

export function parseAddr(a) {
  const m = ADDR_RE.exec(a.toUpperCase());
  if (!m) throw new Error(`Bad address ${a}`);
  return { row: parseInt(m[2], 10), col: lettersToCol(m[1]) };
}

export function parseRange(r) {
  const [a, b] = r.split(':');
  const p = parseAddr(a);
  const q = b ? parseAddr(b) : p;
  return {
    r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col),
    r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col),
  };
}

export function rangeToString({ r1, c1, r2, c2 }) {
  const a = addr(r1, c1);
  return r1 === r2 && c1 === c2 ? a : `${a}:${addr(r2, c2)}`;
}

// Excel serial date <-> JS (UTC) helpers
const EPOCH = Date.UTC(1899, 11, 30);
export function serialToDate(serial) {
  return new Date(EPOCH + Math.floor(serial) * 86400000);
}
export function dateToSerial(y, m, d) {
  return (Date.UTC(y, m, d) - EPOCH) / 86400000;
}
export function todaySerial() {
  const n = new Date();
  return dateToSerial(n.getFullYear(), n.getMonth(), n.getDate());
}
