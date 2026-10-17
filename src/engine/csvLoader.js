import Papa from 'papaparse';
import labels from '../data/bomDumpLabels.json';

/** Decode a BOM CSV: UTF-8 first; fall back to Windows-1252 when the bytes are not valid UTF-8 (e.g. a lone 0xB0 degree sign). */
export async function readFileText(file) {
  const buf = await file.arrayBuffer();
  let text = new TextDecoder('utf-8', { fatal: false }).decode(buf);
  if (text.includes('\uFFFD')) text = new TextDecoder('windows-1252').decode(buf);
  return text.replace(/^\uFEFF/, '');
}

const clean = (s) => s.replace(/Â°/g, '°').replace(/\u00A0/g, ' ');
const norm = (s) =>
  clean(String(s)).replace(/\s+for years.*$/i, '').replace(/\s+/g, ' ').trim().toLowerCase();

function toCell(raw) {
  if (raw === undefined || raw === null || raw === '') return null;
  const s = clean(String(raw));
  if (/^-?\d+(\.\d+)?$/.test(s.trim()) && s.trim() === s) return Number(s); // numbers; leave ' ' placeholders as text like Excel does
  return s;
}

/**
 * Convert CSV text into the same row/column grid as the "BOM Dump" sheet:
 * CSV line n  ->  sheet row n  (blank lines are kept so row numbers line up).
 * If the "Statistic Element" header is not on row 13, the grid is shifted so it is.
 */
export function parseBomCsv(text, fileName = 'BOM Dump.csv') {
  const parsed = Papa.parse(text, { skipEmptyLines: false });
  let rows = parsed.data.map((r) => r.map(toCell));
  // trailing blank line produced by a final newline
  while (rows.length && rows[rows.length - 1].every((c) => c === null)) rows.pop();

  const headerIdx = rows.findIndex((r) => r[0] && norm(r[0]) === labels.headerText);
  if (headerIdx === -1) {
    throw new Error('This file does not look like a BOM climate-statistics CSV: no "Statistic Element" header row was found.');
  }
  const warnings = [];
  const shift = labels.headerRow - 1 - headerIdx;
  if (shift > 0) rows = [...Array(shift).fill([]), ...rows];
  else if (shift < 0) rows = rows.slice(-shift);
  if (shift !== 0) warnings.push(`The header row was found on line ${headerIdx + 1}; rows were realigned so it sits on row ${labels.headerRow}, as in the workbook.`);

  // sanity check that every row referenced by the BOM Extract formulas holds the expected statistic
  for (const [row, expected] of Object.entries(labels.rows)) {
    const found = rows[Number(row) - 1]?.[0];
    if (!found || !norm(found).startsWith(expected.slice(0, 28))) {
      warnings.push(`BOM Dump row ${row}: expected "${expected}" but found "${found ? norm(found) : '(empty)'}". Values on BOM Extract may be misaligned.`);
    }
  }

  const station = rows[1]?.[0] || '';
  return { fileName, rows, warnings, station };
}
