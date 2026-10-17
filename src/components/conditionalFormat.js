// Evaluates a sheet's conditional formatting rules into per-cell overrides.
import { parseFormula, offsetRelative } from '../engine/parser.js';
import { parseRange, addr } from '../engine/address.js';
import { isError, toBool, compare } from '../engine/values.js';

const astCache = new Map();
const ast = (f) => {
  if (!astCache.has(f)) astCache.set(f, parseFormula(f));
  return astCache.get(f);
};

function hexToRgb(h) {
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
}
function mix(a, b, t) {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  return x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function percentile(sorted, p) {
  if (!sorted.length) return 0;
  const k = (sorted.length - 1) * p;
  const lo = Math.floor(k);
  const hi = Math.ceil(k);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (k - lo);
}

function cellsOf(sqref) {
  const out = [];
  for (const part of sqref.split(' ')) {
    const g = parseRange(part);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push({ r, c, g });
  }
  return out;
}

function cfvoValue(cfvo, sorted, wb, sheet) {
  switch (cfvo.type) {
    case 'min': return sorted[0];
    case 'max': return sorted[sorted.length - 1];
    case 'percentile': return percentile(sorted, Number(cfvo.val) / 100);
    case 'percent': return sorted[0] + (sorted[sorted.length - 1] - sorted[0]) * (Number(cfvo.val) / 100);
    case 'num': return Number(cfvo.val);
    case 'formula': return Number(wb.evalAst(sheet, ast(String(cfvo.val))));
    default: return sorted[0];
  }
}

export function computeConditionalFormats(wb, sheet) {
  const s = wb.sheet(sheet);
  const result = new Map(); // addr -> {fill, color, b, priority}
  const rules = [...s.cf].sort((a, b) => a.priority - b.priority);
  const set = (a, prio, patch) => {
    const cur = result.get(a) || {};
    const next = { ...cur };
    for (const [k, v] of Object.entries(patch)) {
      if (v === undefined || v === null) continue;
      if (next[`${k}_p`] === undefined || prio < next[`${k}_p`]) {
        next[k] = v;
        next[`${k}_p`] = prio;
      }
    }
    result.set(a, next);
  };
  for (const rule of rules) {
    const cells = cellsOf(rule.sqref);
    if (rule.type === 'colorScale') {
      const vals = [];
      for (const { r, c } of cells) {
        const v = wb.getValue(sheet, addr(r, c));
        if (typeof v === 'number') vals.push(v);
      }
      if (!vals.length) continue;
      vals.sort((a, b) => a - b);
      const stops = rule.cfvo.map((v) => cfvoValue(v, vals, wb, sheet));
      const colors = rule.colors;
      for (const { r, c } of cells) {
        const v = wb.getValue(sheet, addr(r, c));
        if (typeof v !== 'number') continue;
        let fill;
        if (v <= stops[0]) fill = colors[0];
        else if (v >= stops[stops.length - 1]) fill = colors[colors.length - 1];
        else {
          for (let i = 0; i < stops.length - 1; i++) {
            if (v >= stops[i] && v <= stops[i + 1]) {
              const span = stops[i + 1] - stops[i];
              fill = mix(colors[i], colors[i + 1], span ? (v - stops[i]) / span : 0);
              break;
            }
          }
        }
        set(addr(r, c), rule.priority, { fill });
      }
      continue;
    }
    if (rule.type === 'cellIs') {
      const operands = rule.formula.map((f) => wb.evalAst(sheet, ast(f)));
      for (const { r, c } of cells) {
        const v = wb.getValue(sheet, addr(r, c));
        if (isError(v)) continue;
        const x = v === null ? 0 : v;
        let ok = false;
        const cmp = (o) => compare(x, o);
        switch (rule.operator) {
          case 'equal': ok = cmp(operands[0]) === 0 && !(v === null || v === ''); break;
          case 'notEqual': ok = cmp(operands[0]) !== 0; break;
          case 'greaterThan': ok = typeof x === typeof operands[0] ? cmp(operands[0]) > 0 : false; break;
          case 'greaterThanOrEqual': ok = cmp(operands[0]) >= 0; break;
          case 'lessThan': ok = cmp(operands[0]) < 0; break;
          case 'lessThanOrEqual': ok = cmp(operands[0]) <= 0; break;
          case 'between': ok = cmp(operands[0]) >= 0 && cmp(operands[1]) <= 0; break;
          case 'notBetween': ok = cmp(operands[0]) < 0 || cmp(operands[1]) > 0; break;
          default: ok = false;
        }
        if (ok && rule.dxf) set(addr(r, c), rule.priority, { fill: rule.dxf.fill, color: rule.dxf.color, b: rule.dxf.b });
      }
      continue;
    }
    if (rule.type === 'expression') {
      const base = ast(rule.formula[0]);
      const first = cellsOf(rule.sqref.split(' ')[0])[0];
      for (const { r, c } of cells) {
        const v = wb.evalAst(sheet, offsetRelative(base, r - first.r, c - first.c));
        const b = toBool(v);
        if (b === true && rule.dxf) set(addr(r, c), rule.priority, { fill: rule.dxf.fill, color: rule.dxf.color, b: rule.dxf.b });
      }
    }
  }
  return result;
}
