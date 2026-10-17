// Excel formula tokenizer, parser (to AST) and serializer.
import { lettersToCol, colToLetters } from './address.js';

const ERRORS = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!'];

function tokenize(src) {
  const tokens = [];
  let i = 0;
  const s = src;
  while (i < s.length) {
    const ch = s[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
    if (ch === '"') {
      let j = i + 1;
      let str = '';
      while (j < s.length) {
        if (s[j] === '"') {
          if (s[j + 1] === '"') { str += '"'; j += 2; continue; }
          break;
        }
        str += s[j++];
      }
      tokens.push({ t: 'str', v: str });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const e = ERRORS.find((x) => s.startsWith(x, i));
      if (!e) throw new Error(`Bad error literal at ${i} in ${src}`);
      tokens.push({ t: 'err', v: e });
      i += e.length;
      continue;
    }
    if (ch === "'") {
      // quoted sheet name: 'Sheet Name'!
      let j = i + 1;
      let name = '';
      while (j < s.length) {
        if (s[j] === "'") {
          if (s[j + 1] === "'") { name += "'"; j += 2; continue; }
          break;
        }
        name += s[j++];
      }
      if (s[j + 1] !== '!') throw new Error(`Expected ! after sheet name in ${src}`);
      tokens.push({ t: 'sheet', v: name });
      i = j + 2;
      continue;
    }
    if (/[0-9.]/.test(ch)) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(s.slice(i));
      // A number immediately followed by ':' and digits could be a row range (1:3) - unsupported, treat as number
      tokens.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    if (/[A-Za-z_$\\]/.test(ch)) {
      const m = /^[A-Za-z_$\\][A-Za-z0-9_.$]*/.exec(s.slice(i));
      const word = m[0];
      i += word.length;
      if (s[i] === '!') {
        tokens.push({ t: 'sheet', v: word });
        i++;
        continue;
      }
      if (s[i] === '(') {
        tokens.push({ t: 'func', v: word.toUpperCase() });
        i++;
        continue;
      }
      if (/^\$?[A-Za-z]{1,3}\$?\d+$/.test(word)) {
        tokens.push({ t: 'cell', v: word });
        continue;
      }
      const up = word.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') { tokens.push({ t: 'bool', v: up === 'TRUE' }); continue; }
      tokens.push({ t: 'name', v: word });
      continue;
    }
    const two = s.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') { tokens.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%:'.includes(ch)) { tokens.push({ t: 'op', v: ch }); i++; continue; }
    if (ch === '(') { tokens.push({ t: '(' }); i++; continue; }
    if (ch === ')') { tokens.push({ t: ')' }); i++; continue; }
    if (ch === ',') { tokens.push({ t: ',' }); i++; continue; }
    throw new Error(`Unexpected character '${ch}' in formula ${src}`);
  }
  return tokens;
}

function cellRef(word, sheet) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(word);
  return {
    type: 'ref', sheet,
    col: lettersToCol(m[2].toUpperCase()), row: parseInt(m[4], 10),
    colAbs: m[1] === '$', rowAbs: m[3] === '$',
  };
}

const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function parseFormula(formula) {
  const src = formula.startsWith('=') ? formula.slice(1) : formula;
  const tokens = tokenize(src);
  let p = 0;
  const peek = () => tokens[p];
  const next = () => tokens[p++];

  function parseExpr(minPrec) {
    let left = parseUnary();
    for (;;) {
      const tk = peek();
      if (!tk || tk.t !== 'op' || !(tk.v in BIN_PREC)) break;
      const prec = BIN_PREC[tk.v];
      if (prec < minPrec) break;
      next();
      const right = parseExpr(tk.v === '^' ? prec : prec + 1);
      left = { type: 'bin', op: tk.v, l: left, r: right };
    }
    return left;
  }

  function parseUnary() {
    const tk = peek();
    if (tk && tk.t === 'op' && (tk.v === '-' || tk.v === '+')) {
      next();
      return { type: 'unary', op: tk.v, arg: parseUnary() };
    }
    return parsePostfix();
  }

  function parsePostfix() {
    let e = parsePrimary();
    while (peek() && peek().t === 'op' && peek().v === '%') {
      next();
      e = { type: 'pct', arg: e };
    }
    return e;
  }

  function parsePrimary() {
    const tk = next();
    if (!tk) throw new Error(`Unexpected end of formula ${formula}`);
    switch (tk.t) {
      case 'num': return { type: 'num', v: tk.v };
      case 'str': return { type: 'str', v: tk.v };
      case 'bool': return { type: 'bool', v: tk.v };
      case 'err': return { type: 'err', v: tk.v };
      case 'sheet': {
        const c = next();
        if (!c || c.t !== 'cell') throw new Error(`Expected cell after sheet in ${formula}`);
        return maybeRange(cellRef(c.v, tk.v));
      }
      case 'cell': return maybeRange(cellRef(tk.v, null));
      case 'name': return { type: 'name', name: tk.v };
      case 'func': {
        const args = [];
        if (peek() && peek().t === ')') { next(); return { type: 'func', name: tk.v, args }; }
        for (;;) {
          if (peek() && (peek().t === ',' || peek().t === ')')) args.push({ type: 'missing' });
          else args.push(parseExpr(1));
          const sep = next();
          if (!sep) throw new Error(`Unclosed function in ${formula}`);
          if (sep.t === ')') break;
          if (sep.t !== ',') throw new Error(`Expected , in ${formula}`);
        }
        return { type: 'func', name: tk.v, args };
      }
      case '(': {
        const e = parseExpr(1);
        const c = next();
        if (!c || c.t !== ')') throw new Error(`Expected ) in ${formula}`);
        return { type: 'paren', e };
      }
      default:
        throw new Error(`Unexpected token ${JSON.stringify(tk)} in ${formula}`);
    }
  }

  function maybeRange(ref) {
    const tk = peek();
    if (tk && tk.t === 'op' && tk.v === ':') {
      next();
      const c = next();
      if (!c || c.t !== 'cell') throw new Error(`Bad range in ${formula}`);
      const end = cellRef(c.v, ref.sheet);
      return { type: 'range', sheet: ref.sheet, a: ref, b: end };
    }
    return ref;
  }

  const ast = parseExpr(1);
  if (p !== tokens.length) throw new Error(`Unexpected trailing tokens in ${formula}`);
  return ast;
}

function quoteSheet(name) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function refStr(r) {
  return `${r.colAbs ? '$' : ''}${colToLetters(r.col)}${r.rowAbs ? '$' : ''}${r.row}`;
}

export function serialize(n) {
  switch (n.type) {
    case 'num': return String(n.v);
    case 'str': return `"${n.v.replace(/"/g, '""')}"`;
    case 'bool': return n.v ? 'TRUE' : 'FALSE';
    case 'err': return n.v;
    case 'missing': return '';
    case 'ref': return (n.sheet ? `${quoteSheet(n.sheet)}!` : '') + (n.invalid ? '#REF!' : refStr(n));
    case 'range': return (n.sheet ? `${quoteSheet(n.sheet)}!` : '') + `${refStr(n.a)}:${refStr(n.b)}`;
    case 'name': return n.name;
    case 'func': return `${n.name}(${n.args.map(serialize).join(',')})`;
    case 'paren': return `(${serialize(n.e)})`;
    case 'unary': return n.op + serialize(n.arg);
    case 'pct': return `${serialize(n.arg)}%`;
    case 'bin': return serialize(n.l) + n.op + serialize(n.r);
    default: throw new Error(`Cannot serialize ${n.type}`);
  }
}

// Deep-map every ref (including range endpoints) through fn(ref, sheetOfRef) -> new ref.
export function mapRefs(n, fn) {
  switch (n.type) {
    case 'ref': return fn(n);
    case 'range': {
      const a = fn({ ...n.a, sheet: n.sheet });
      const b = fn({ ...n.b, sheet: n.sheet });
      return { ...n, a: { ...a, sheet: n.sheet }, b: { ...b, sheet: n.sheet } };
    }
    case 'func': return { ...n, args: n.args.map((x) => mapRefs(x, fn)) };
    case 'paren': return { ...n, e: mapRefs(n.e, fn) };
    case 'unary': case 'pct': return { ...n, arg: mapRefs(n.arg, fn) };
    case 'bin': return { ...n, l: mapRefs(n.l, fn), r: mapRefs(n.r, fn) };
    default: return n;
  }
}

// Shift relative references by (dr, dc): Excel copy/paste semantics.
export function offsetRelative(ast, dr, dc) {
  return mapRefs(ast, (r) => ({
    ...r,
    row: r.rowAbs ? r.row : r.row + dr,
    col: r.colAbs ? r.col : r.col + dc,
  }));
}
