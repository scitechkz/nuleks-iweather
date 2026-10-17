// Builds a live Workbook from the template + user state (uploaded BOM csv, categories, edits).
import template from '../template/workbook.json';
import { Workbook } from '../engine/workbook.js';
import { addr, colToLetters, parseRange, rangeToString } from '../engine/address.js';

export const S = {
  BOM: 'BOM csv data',
  CLIMATE: 'Climate',
  WC: 'Weather Calculation',
  CSWS: 'Critical Site Works Summary',
  RAIN: 'Rain+Other Factors',
  CAL: 'Project Site Calendar',
};

const BASE_CATEGORIES = 5;
const RAIN_ROWS = [8, 9, 10, 11]; // <1mm, >1mm, >10mm, >25mm on Rain+Other Factors
const CLIMATE_RAIN_ROWS = [35, 37, 39, 40]; // matching rain-day rows on Climate
const MONTH_COLS = 12;
const WC_FIRST_MONTH_COL = 8; // H
const BLOCK_HEIGHT = 7; // header, 4 rain rows, total, spacer
const WC_FIRST_BLOCK_HEADER = 32;
const WC_IMPACT_TITLE = 67;
const STATISTIC_ELEMENT_ROW = 16;

export const rainCols = (i) => ({ direct: 3 + 4 * i, cons: 4 + 4 * i, total: 5 + 4 * i });

const tcell = (sheet, a) => template.sheets[sheet].cells[a];
const tval = (sheet, a) => tcell(sheet, a)?.v ?? null;

export function defaultCategories() {
  const cats = [];
  for (let i = 0; i < BASE_CATEGORIES; i++) {
    const { direct, cons } = rainCols(i);
    cats.push({
      title: String(tval(S.RAIN, addr(3, direct)) ?? `EXPOSURE ${i + 1}`).trim(),
      // the template pads some names with spaces and a trailing '.'
      label: String(tval(S.WC, addr(WC_FIRST_BLOCK_HEADER + 5 + BLOCK_HEIGHT * i, 1)) ?? '').replace(/\s+\.\s*$/, '').trim(),
      direct: RAIN_ROWS.map((r) => tval(S.RAIN, addr(r, direct)) ?? 0),
      cons: RAIN_ROWS.map((r) => tval(S.RAIN, addr(r, cons)) ?? 0),
    });
  }
  return cats;
}

export function newCategory(n) {
  return {
    title: `EXPOSURE ${n}\nCUSTOM RAIN IMPACTS`,
    label: `RAIN EXPOSURE ${n} : Custom Rain Impacts`,
    direct: [0, 0, 0, 0],
    cons: [0, 0, 0, 0],
  };
}

// Expected statistic labels on the BOM sheet (rows 16+), used to validate uploads.
export function expectedBomLabels() {
  const out = [];
  const s = template.sheets[S.BOM];
  for (let r = STATISTIC_ELEMENT_ROW; r <= s.maxRow; r++) {
    const v = s.cells[addr(r, 1)]?.v;
    if (typeof v === 'string') out.push({ row: r, label: v });
  }
  return out;
}

// Comparison key for statistic labels: ignores the "for years ..." period, unit spelling and punctuation.
const statKey = (s) => normaliseLabel(s).toLowerCase().replace(/\s+for years.*$/, '')
  .replace(/degrees c/g, '').replace(/[^a-z0-9]+/g, ' ').trim();

// Degree sign as UTF-8, mis-decoded Windows-1252 ("Â°") or replacement char.
const DEG = '(?:Â?°|�)';
const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

// Rewrite a BoM statistic label in the wording used by the "BOM csv data" sheet:
//  "Mean maximum temperature (°C) for years ..."      -> "(Degrees C)"
//  "Date of Highest temperature  °C  for years ..."   -> "Date of Highest temperature for years ..."
//  "Mean number of days >= 30 °C for years ..."        -> ">= 30 Degrees C"
export function normaliseLabel(label) {
  let s = String(label ?? '');
  if (/^\s*Date of /i.test(s)) {
    s = s.replace(new RegExp(`\\s*\\(?${DEG}\\s*C\\)?\\s*(?=\\s+for years)`, 'i'), '');
    s = s.replace(/\s*\(?Degrees C\)?\s*(?=\s+for years)/i, '');
  }
  s = s.replace(new RegExp(`\\(\\s*${DEG}\\s*C\\s*\\)`, 'g'), '(Degrees C)');
  s = s.replace(new RegExp(`${DEG}\\s*C\\b`, 'g'), 'Degrees C');
  return s.replace(/ {2,}(?=for years)/, ' ');
}

// Re-shape the uploaded BoM csv so it matches the "BOM csv data" sheet exactly:
// sheet-style labels, BoM's mislabelled Decile 9 rainfall row, real dates, no trailing empty columns.
export function normaliseBomRows(rows) {
  const out = rows.map((r) => {
    const row = [...r];
    while (row.length && (row[row.length - 1] === '' || row[row.length - 1] == null)) row.pop();
    if (row.length) row[0] = normaliseLabel(row[0]);
    return row;
  });
  // BoM exports label the Decile 9 rainfall row as a second "Decile 5 (median)" row.
  for (let i = 1; i < out.length; i++) {
    const a = String(out[i - 1][0] ?? '');
    const b = String(out[i][0] ?? '');
    if (/^Decile 5 \(median\) monthly rainfall/i.test(a) && /^Decile 5 \(median\) monthly rainfall/i.test(b)) {
      out[i][0] = b.replace(/^Decile 5 \(median\)/i, 'Decile 9');
    }
  }
  return out;
}

// Convert parsed CSV rows into a sparse {row: [values]} map aligned so that
// "Statistic Element" lands on row 16, exactly as in the template.
export function alignBomCsv(rawRows) {
  const rows = normaliseBomRows(rawRows);
  const idx = rows.findIndex((r) => String(r[0] ?? '').trim().toLowerCase() === 'statistic element');
  if (idx < 0) throw new Error('Could not find the "Statistic Element" header row in the CSV file.');
  const offset = STATISTIC_ELEMENT_ROW - (idx + 1);
  const grid = {};
  rows.forEach((r, i) => {
    const sheetRow = i + 1 + offset;
    if (sheetRow < 5) return; // rows 1-4 hold the template instructions
    grid[sheetRow] = r.map((v, ci) => (ci === 0 ? (v === '' ? null : v) : parseCsvValue(v)));
  });
  // validation against template statistic labels
  const warnings = [];
  for (const { row, label } of expectedBomLabels()) {
    const got = grid[row]?.[0];
    if (statKey(got) !== statKey(label)) {
      warnings.push({ row, expected: label.trim(), got: got == null ? '(missing)' : String(got).trim() });
    }
  }
  return { grid, warnings, statisticRow: idx + 1 };
}

function parseCsvValue(v) {
  if (v === undefined || v === null) return null;
  const s = String(v);
  if (s === '') return null;
  const t = s.trim();
  if (t !== '' && /^-?\d*\.?\d+(e[+-]?\d+)?$/i.test(t)) return Number(t);
  const serial = (y, m, d) => (Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000;
  let dm = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (dm) return serial(+dm[3], +dm[2] - 1, +dm[1]);
  // BoM style "27-Jan-26" / "27-Jan-2026"; two-digit years up to the current year are 20xx, the rest 19xx.
  dm = /^(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*[-\s](\d{2}|\d{4})$/.exec(t);
  if (dm && MONTHS[dm[2].toLowerCase()] !== undefined) {
    let y = +dm[3];
    if (dm[3].length === 2) y += y <= new Date().getFullYear() % 100 ? 2000 : 1900;
    return serial(y, MONTHS[dm[2].toLowerCase()], +dm[1]);
  }
  return s;
}

function applyBom(wb, grid) {
  const s = wb.sheet(S.BOM);
  for (const [a, c] of [...s.cells]) {
    const row = parseInt(a.replace(/^[A-Z]+/, ''), 10);
    if (row >= 5) {
      if (c.s === undefined) s.cells.delete(a);
      else s.cells.set(a, { s: c.s });
    }
  }
  for (const [r, vals] of Object.entries(grid)) {
    vals.forEach((v, ci) => {
      if (v === null) return;
      const a = addr(Number(r), ci + 1);
      const prev = s.cells.get(a) || {};
      s.cells.set(a, { s: prev.s, v });
      if (Number(r) > s.maxRow) s.maxRow = Number(r);
      if (ci + 1 > s.maxCol) s.maxCol = ci + 1;
    });
  }
  wb.invalidate();
}

function cloneCell(wb, sheet, from, to, extra = {}) {
  const s = wb.sheet(sheet);
  const src = s.cells.get(from);
  s.cells.set(to, { s: src?.s, ...extra });
}

function addRainFactorColumns(wb, i) {
  const s = wb.sheet(S.RAIN);
  const src = rainCols(BASE_CATEGORIES - 1);
  const dst = rainCols(i);
  const shift = dst.direct - src.direct;
  for (let r = 3; r <= 11; r++) {
    for (let c = src.direct; c <= src.total + 1; c++) {
      const from = addr(r, c);
      const to = addr(r, c + shift);
      const cell = s.cells.get(from);
      if (!cell) continue;
      const copy = { s: cell.s };
      if (r === 6 || r === 7) copy.v = cell.v;
      s.cells.set(to, copy);
    }
  }
  for (let c = src.direct; c <= src.total + 1; c++) {
    if (s.cols[c]) s.cols[c + shift] = { ...s.cols[c] };
  }
  s.merges.push(rangeToString({ r1: 3, c1: dst.direct, r2: 5, c2: dst.total }));
  RAIN_ROWS.forEach((r) => {
    s.cells.set(addr(r, dst.total), {
      s: s.cells.get(addr(r, src.total))?.s,
      f: `=SUM(${addr(r, dst.direct)}:${addr(r, dst.cons)})`,
    });
  });
  for (const rule of [...s.cf]) {
    const g = parseRange(rule.sqref.split(' ')[0]);
    if (rule.sqref.includes(' ') || g.r1 !== 8 || g.c1 < src.direct || g.c2 > src.total) continue;
    s.cf.push({ ...rule, sqref: rangeToString({ ...g, c1: g.c1 + shift, c2: g.c2 + shift }) });
  }
  s.maxCol = Math.max(s.maxCol, dst.total + 1);
  // Rain summary table row (B..F) for this level
  const summaryRow = 20 + i;
  wb.insertRows(S.RAIN, summaryRow, 1);
  wb.copyRowStyle(S.RAIN, summaryRow - 1, summaryRow, 1, 6);
  s.cells.set(addr(summaryRow, 2), { s: s.cells.get(addr(summaryRow - 1, 2))?.s, v: i + 1 });
  RAIN_ROWS.forEach((r, k) => {
    s.cells.set(addr(summaryRow, 3 + k), {
      s: s.cells.get(addr(summaryRow - 1, 3 + k))?.s,
      f: `=${colToLetters(dst.total)}${r}`,
    });
  });
}

function addWeatherCalcBlock(wb, i) {
  const s = wb.sheet(S.WC);
  const header = WC_FIRST_BLOCK_HEADER + BLOCK_HEIGHT * i;
  const srcHeader = WC_FIRST_BLOCK_HEADER + BLOCK_HEIGHT * (BASE_CATEGORIES - 1);
  wb.insertRows(S.WC, header, BLOCK_HEIGHT);
  const lastCol = WC_FIRST_MONTH_COL + MONTH_COLS - 1;
  for (let k = 0; k < BLOCK_HEIGHT; k++) wb.copyRowStyle(S.WC, srcHeader + k, header + k, 1, lastCol + 3);
  const { direct, cons } = rainCols(i);
  const rain = (c, r) => `'${S.RAIN}'!${colToLetters(c)}${r}`;
  ['Rainfall', 'Direct', 'Consequential', 'Total'].forEach((t, k) => {
    cloneCell(wb, S.WC, addr(srcHeader, k + 1), addr(header, k + 1), { v: t });
  });
  RAIN_ROWS.forEach((rr, k) => {
    const r = header + 1 + k;
    const sr = srcHeader + 1 + k;
    cloneCell(wb, S.WC, addr(sr, 1), addr(r, 1), { v: s.cells.get(addr(sr, 1))?.v });
    cloneCell(wb, S.WC, addr(sr, 2), addr(r, 2), { f: `=${rain(direct, rr)}` });
    cloneCell(wb, S.WC, addr(sr, 3), addr(r, 3), { f: `=${rain(cons, rr)}` });
    cloneCell(wb, S.WC, addr(sr, 4), addr(r, 4), { f: `=SUM(B${r}:C${r})` });
    cloneCell(wb, S.WC, addr(sr, 6), addr(r, 6), { f: `=SUM(H${r}:S${r})` });
    for (let m = 0; m < MONTH_COLS; m++) {
      const col = colToLetters(WC_FIRST_MONTH_COL + m);
      const clim = `Climate!${colToLetters(2 + m)}${CLIMATE_RAIN_ROWS[k]}`;
      const hours = k < 3 ? '*HoursPerDayRatio' : '';
      cloneCell(wb, S.WC, addr(sr, WC_FIRST_MONTH_COL + m), addr(r, WC_FIRST_MONTH_COL + m),
        { f: `=+${clim}*$D${r}${hours}*${col}$3` });
    }
  });
  const total = header + 5;
  cloneCell(wb, S.WC, addr(srcHeader + 5, 1), addr(total, 1), { v: '' });
  cloneCell(wb, S.WC, addr(srcHeader + 5, 6), addr(total, 6), { f: `=SUM(H${total}:S${total})` });
  for (let m = 0; m < MONTH_COLS; m++) {
    const col = colToLetters(WC_FIRST_MONTH_COL + m);
    cloneCell(wb, S.WC, addr(srcHeader + 5, WC_FIRST_MONTH_COL + m), addr(total, WC_FIRST_MONTH_COL + m),
      { f: `=SUM(${col}${header + 1}:${col}${header + 4})` });
  }
  // merges + conditional formats of the template block
  const dr = header - srcHeader;
  for (const m of [...s.merges]) {
    const g = parseRange(m);
    if (g.r1 >= srcHeader && g.r2 < srcHeader + BLOCK_HEIGHT) {
      s.merges.push(rangeToString({ ...g, r1: g.r1 + dr, r2: g.r2 + dr }));
    }
  }
  for (const rule of [...s.cf]) {
    if (rule.sqref.includes(' ')) continue;
    const g = parseRange(rule.sqref);
    if (g.r1 >= srcHeader && g.r2 < srcHeader + BLOCK_HEIGHT) {
      s.cf.push({ ...rule, sqref: rangeToString({ ...g, r1: g.r1 + dr, r2: g.r2 + dr }) });
    }
  }
  // Impact table row for this exposure level
  const extra = i - BASE_CATEGORIES + 1;
  const impactRow = WC_IMPACT_TITLE + BLOCK_HEIGHT * extra + 2 + i;
  wb.insertRows(S.WC, impactRow, 1);
  wb.copyRowStyle(S.WC, impactRow - 1, impactRow, 1, lastCol + 3);
  cloneCell(wb, S.WC, addr(impactRow - 1, 4), addr(impactRow, 4), { f: `=SUM(H${impactRow}:S${impactRow})` });
  cloneCell(wb, S.WC, addr(impactRow - 1, 5), addr(impactRow, 5), { f: `=D${impactRow}/'${S.CAL}'!$U$37` });
  cloneCell(wb, S.WC, addr(impactRow - 1, 6), addr(impactRow, 6), { v: i + 1 });
  s.merges.push(`F${impactRow}:G${impactRow}`);
  for (let m = 0; m < MONTH_COLS; m++) {
    const col = colToLetters(WC_FIRST_MONTH_COL + m);
    cloneCell(wb, S.WC, addr(impactRow - 1, WC_FIRST_MONTH_COL + m), addr(impactRow, WC_FIRST_MONTH_COL + m),
      { f: `=${col}${total}` });
  }
}

// Layout of the Weather Calculation sheet for a given number of categories.
export function weatherLayout(count) {
  const extra = Math.max(0, count - BASE_CATEGORIES);
  const title = WC_IMPACT_TITLE + BLOCK_HEIGHT * extra;
  const header = title + 1;
  const first = header + 1;
  return {
    title, header, first,
    exposureRows: Array.from({ length: count }, (_, i) => first + i),
    blockTotal: (i) => WC_FIRST_BLOCK_HEADER + BLOCK_HEIGHT * i + 5,
    last: first + count - 1 + 8, // + Heat, Cold, Wind, Client, Subbie, IR, Plant, Other
  };
}

export function applyCategoryValues(wb, layout, cats) {
  const rainSheet = wb.sheet(S.RAIN);
  const wc = wb.sheet(S.WC);
  cats.forEach((cat, i) => {
    const { direct, cons } = rainCols(i);
    rainSheet.cells.set(addr(3, direct), { ...rainSheet.cells.get(addr(3, direct)), v: cat.title });
    RAIN_ROWS.forEach((r, k) => {
      rainSheet.cells.set(addr(r, direct), { ...rainSheet.cells.get(addr(r, direct)), v: Number(cat.direct[k]) || 0 });
      rainSheet.cells.set(addr(r, cons), { ...rainSheet.cells.get(addr(r, cons)), v: Number(cat.cons[k]) || 0 });
    });
    const totalRow = layout.blockTotal(i);
    wc.cells.set(addr(totalRow, 1), { ...wc.cells.get(addr(totalRow, 1)), v: cat.label });
    // Category name shown alongside its row in the Weather Analysis Impact Table
    const impactRow = layout.exposureRows[i];
    wc.cells.set(addr(impactRow, 2), { s: wc.cells.get(addr(impactRow, 2))?.s, f: `=A${totalRow}` });
  });
  wb.invalidate();
}

// Cells holding formulas can never be changed by users, and users can't enter formulas.
export function isLocked(wb, sheet, a) {
  return wb.getCell(sheet, a)?.f !== undefined;
}
export function isAllowedInput(wb, sheet, a, raw) {
  return !isLocked(wb, sheet, a) && !String(raw ?? '').trim().startsWith('=');
}

// Template content that must not appear in the app.
function applyTemplateCleanups(wb) {
  // Instruction notes in columns U and V of Weather Calculation, and its numbered marker shapes
  const wc = wb.sheet(S.WC);
  for (const a of [...wc.cells.keys()]) {
    if (/^[UV]\d+$/.test(a)) wc.cells.delete(a);
  }
  wc.merges = wc.merges.filter((m) => !/^[UV]\d/.test(m));
  wc.shapes = [];
  wc.maxCol = Math.min(wc.maxCol, 19);
  // No planner name on the Critical Site Works Summary, and no spreadsheet instructions
  const cs = wb.sheet(S.CSWS);
  cs.cells.set('B5', { s: cs.cells.get('B5')?.s });
  removeCells(cs, (a) => a === 'DB2');
  // Project Site Calendar: spreadsheet instructions (columns W-AB) and tab references
  const cal = wb.sheet(S.CAL);
  removeCells(cal, (a) => /^(W|X|Y|Z|AA|AB)\d+$/.test(a) || a === 'AE20' || a === 'AE33');
  for (let c = 23; c <= 29; c++) cal.cols[c] = { ...(cal.cols[c] || {}), hidden: true };
  const b29 = cal.cells.get('B29');
  if (b29?.v) cal.cells.set('B29', { ...b29, v: String(b29.v).replace(/\s*\(copy\/paste[^)]*\)/i, '') });
}

function removeCells(sheet, test) {
  for (const a of [...sheet.cells.keys()]) if (test(a)) sheet.cells.delete(a);
  sheet.merges = sheet.merges.filter((m) => !test(m.split(':')[0]));
}

export function buildWorkbook({ bomGrid, categories, overrides }) {
  const wb = new Workbook(template);
  applyTemplateCleanups(wb);
  // The uploaded csv is the only climate data source: start with an empty BOM csv data sheet.
  applyBom(wb, bomGrid || {});
  const cats = categories && categories.length ? categories : defaultCategories();
  for (let i = BASE_CATEGORIES; i < cats.length; i++) {
    addRainFactorColumns(wb, i);
    addWeatherCalcBlock(wb, i);
  }
  const layout = weatherLayout(cats.length);
  applyCategoryValues(wb, layout, cats);
  for (const [key, raw] of Object.entries(overrides || {})) {
    const [sheet, a] = key.split('|');
    if (!wb.sheets[sheet] || !isAllowedInput(wb, sheet, a, raw)) continue;
    try {
      wb.setInput(sheet, a, raw);
    } catch (e) {
      wb.setCell(sheet, a, { v: '#ERROR: ' + e.message });
    }
  }
  wb.invalidate();
  return { wb, layout };
}
