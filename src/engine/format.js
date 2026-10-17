// Minimal Excel number-format renderer covering the formats used in the template
// (General, 0, 0.0, 0%, 0.0%, date formats, literal suffixes like 0" days").
import { serialToDate } from './address.js';
import { isError } from './values.js';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
  'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function formatGeneral(v) {
  if (Number.isInteger(v)) return String(v);
  const abs = Math.abs(v);
  if (abs !== 0 && (abs >= 1e11 || abs < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'E');
  // Excel General shows up to ~11 characters
  const intDigits = Math.max(1, Math.floor(Math.log10(abs || 1)) + 1);
  const decimals = Math.max(0, 10 - intDigits);
  return String(parseFloat(v.toFixed(decimals)));
}

const formatCache = new Map();

// Split a format section into tokens: literals and placeholders
function compile(fmt) {
  if (formatCache.has(fmt)) return formatCache.get(fmt);
  const sections = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') inQ = !inQ;
    if (ch === '\\' && !inQ) { cur += ch + (fmt[i + 1] ?? ''); i++; continue; }
    if (ch === ';' && !inQ) { sections.push(cur); cur = ''; continue; }
    cur += ch;
  }
  sections.push(cur);
  const compiled = sections.map((sec) => {
    // strip [$-C09] locale / [Red] colour codes
    const clean = sec.replace(/\[[^\]]*\]/g, '');
    const parts = [];
    let isDate = false;
    let i = 0;
    while (i < clean.length) {
      const ch = clean[i];
      if (ch === '"') {
        const j = clean.indexOf('"', i + 1);
        parts.push({ lit: clean.slice(i + 1, j < 0 ? clean.length : j) });
        i = j < 0 ? clean.length : j + 1;
        continue;
      }
      if (ch === '\\') { parts.push({ lit: clean[i + 1] ?? '' }); i += 2; continue; }
      if (ch === '_') { parts.push({ lit: ' ' }); i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      const dm = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM)/i.exec(clean.slice(i));
      if (dm) {
        isDate = true;
        parts.push({ date: dm[0].toLowerCase() });
        i += dm[0].length;
        continue;
      }
      const nm = /^[#0?,.%]+(E[+-]0+)?/i.exec(clean.slice(i));
      if (nm) { parts.push({ num: nm[0] }); i += nm[0].length; continue; }
      if (ch === '@') { parts.push({ text: true }); i++; continue; }
      parts.push({ lit: ch });
      i++;
    }
    return { parts, isDate, raw: clean };
  });
  formatCache.set(fmt, compiled);
  return compiled;
}

function formatNumberPattern(v, pat) {
  let pct = 0;
  for (const ch of pat) if (ch === '%') pct++;
  let x = v * 100 ** pct;
  const p = pat.replace(/%/g, '');
  const dot = p.indexOf('.');
  const intPat = dot >= 0 ? p.slice(0, dot) : p;
  const decPat = dot >= 0 ? p.slice(dot + 1) : '';
  const decimals = decPat.replace(/[^0#?]/g, '').length;
  const minDec = decPat.replace(/[^0]/g, '').length;
  const thousands = intPat.includes(',');
  const minInt = intPat.replace(/[^0]/g, '').length;
  // Excel rounds half away from zero
  const f = 10 ** decimals;
  x = Math.round(Math.abs(x) * f + 1e-9) / f;
  let [ip, dp = ''] = x.toFixed(decimals).split('.');
  if (decimals > minDec) {
    dp = dp.replace(/0+$/, '');
    while (dp.length < minDec) dp += '0';
  }
  if (ip === '0' && minInt === 0) ip = '';
  while (ip.length < minInt) ip = '0' + ip;
  if (thousands) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return ip + (dp.length || (dot >= 0 && decimals > 0 && minDec > 0) ? '.' + dp : '') + '%'.repeat(pct);
}

function formatDate(v, parts) {
  const d = serialToDate(v);
  const y = d.getUTCFullYear();
  const mo = d.getUTCMonth();
  const day = d.getUTCDate();
  const frac = v - Math.floor(v);
  const secs = Math.round(frac * 86400);
  const hh = Math.floor(secs / 3600);
  const mi = Math.floor((secs % 3600) / 60);
  const ss = secs % 60;
  const hasHour = parts.some((p) => p.date && p.date.startsWith('h'));
  let out = '';
  parts.forEach((p, idx) => {
    if (p.lit !== undefined) { out += p.lit; return; }
    if (p.num) { out += p.num; return; }
    if (p.text) return;
    const t = p.date;
    // 'm' after h or before s means minutes
    const prev = parts.slice(0, idx).reverse().find((q) => q.date);
    const nxt = parts.slice(idx + 1).find((q) => q.date);
    const isMinute = (t === 'm' || t === 'mm') && hasHour
      && ((prev && prev.date.startsWith('h')) || (nxt && nxt.date.startsWith('s')));
    switch (t) {
      case 'yyyy': out += y; break;
      case 'yy': out += String(y).slice(-2); break;
      case 'mmmmm': out += MONTHS[mo][0]; break;
      case 'mmmm': out += MONTHS[mo]; break;
      case 'mmm': out += MONTHS[mo].slice(0, 3); break;
      case 'mm': out += isMinute ? String(mi).padStart(2, '0') : String(mo + 1).padStart(2, '0'); break;
      case 'm': out += isMinute ? mi : mo + 1; break;
      case 'dddd': out += DAYS[d.getUTCDay()]; break;
      case 'ddd': out += DAYS[d.getUTCDay()].slice(0, 3); break;
      case 'dd': out += String(day).padStart(2, '0'); break;
      case 'd': out += day; break;
      case 'hh': out += String(hh).padStart(2, '0'); break;
      case 'h': out += hh; break;
      case 'ss': out += String(ss).padStart(2, '0'); break;
      case 's': out += ss; break;
      default: out += t;
    }
  });
  return out;
}

export function formatValue(v, fmt = 'General', { showZeros = true } = {}) {
  if (v === null || v === undefined) return '';
  if (isError(v)) return v.err;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') {
    const secs = compile(fmt || 'General');
    const ts = secs[3];
    if (ts && ts.parts.some((p) => p.text)) {
      return ts.parts.map((p) => (p.text ? v : p.lit ?? '')).join('');
    }
    return v;
  }
  if (typeof v !== 'number' || Number.isNaN(v)) return String(v);
  if (v === 0 && !showZeros) return '';
  if (!fmt || fmt === 'General') return formatGeneral(v);
  const secs = compile(fmt);
  let sec = secs[0];
  let val = v;
  if (secs.length > 1 && v < 0) { sec = secs[1]; val = -v; } else if (secs.length > 2 && v === 0) sec = secs[2];
  if (sec.isDate) return formatDate(val, sec.parts);
  if (sec.raw.trim() === 'General') return (v < 0 && sec !== secs[0] ? '-' : '') + formatGeneral(val);
  let out = '';
  let usedNum = false;
  for (const p of sec.parts) {
    if (p.lit !== undefined) out += p.lit;
    else if (p.num && !usedNum) { out += formatNumberPattern(sec === secs[0] ? Math.abs(val) : val, p.num); usedNum = true; }
    else if (p.text) out += '';
  }
  if (!usedNum) return out;
  if (sec === secs[0] && v < 0) {
    // only add minus if the rounded value isn't zero
    if (/[1-9]/.test(out)) out = '-' + out;
  }
  return out;
}
