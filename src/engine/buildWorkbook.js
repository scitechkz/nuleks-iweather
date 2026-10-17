import layout from '../data/layout.json';
import { Workbook } from './formulaEngine.js';
import { isoToSerial } from './format.js';

function sheetFromLayout(spec, overrides = {}) {
  const cells = new Map();
  for (const c of spec.cells) {
    if (c.f === undefined && c.v === undefined) continue;
    const entry = {};
    if (c.f) entry.f = c.f;
    else entry.v = c.dt ? isoToSerial(c.v) : c.v;
    cells.set(`${c.r},${c.c}`, entry);
  }
  return { cells };
}

function applyOverrides(sheet, overrides) {
  for (const [a1, val] of Object.entries(overrides)) {
    const m = /^([A-Z]+)(\d+)$/.exec(a1);
    if (!m) continue;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    sheet.cells.set(`${m[2]},${col}`, { v: val });
  }
}

/** rows: 2-D array from parseBomCsv (or []) */
export function buildWorkbook(dumpRows, overrides = {}) {
  const dump = { cells: new Map() };
  (dumpRows || []).forEach((row, ri) =>
    row.forEach((v, ci) => { if (v !== null && v !== undefined) dump.cells.set(`${ri + 1},${ci + 1}`, { v }); }),
  );
  const extract = sheetFromLayout(layout['BOM Extract']);
  const analysis = sheetFromLayout(layout['Weather Analysis']);
  applyOverrides(analysis, overrides);
  return new Workbook(
    { 'BOM Dump': dump, 'BOM Extract': extract, 'Weather Analysis': analysis },
    layout.definedNames,
  );
}

export { layout };
