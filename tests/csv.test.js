import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import Papa from 'papaparse';
import template from '../src/template/workbook.json';
import { alignBomCsv, buildWorkbook, S } from '../src/model/build.js';

const rows = Papa.parse(fs.readFileSync('samples/laverton_bom.csv', 'utf8'), { skipEmptyLines: false }).data;

describe('BoM csv (current export format)', () => {
  const { grid, warnings } = alignBomCsv(rows);
  it('matches the BOM csv data sheet layout with no warnings', () => {
    expect(warnings).toEqual([]);
    expect(grid[5][0]).toBe("Monthly Climate Statistics for 'LAVERTON RAAF' [087031]");
    expect(grid[8][0]).toBe('087031 LAVERTON RAAF');
    expect(grid[16][0]).toBe('Statistic Element');
    expect(grid[17][0]).toBe('Mean maximum temperature (Degrees C) for years 1943 to 2026 ');
    expect(grid[19][0]).toBe('Date of Highest temperature for years 1943 to 2026 ');
    expect(grid[24][0]).toBe('Mean number of days >= 30 Degrees C for years 1943 to 2026 ');
    expect(grid[47][0]).toBe('Decile 9 monthly rainfall (mm) for years 1800 to 3000 ');
    expect(grid[73][0]).toMatch(/^Mean 3pm wind speed/);
    expect(grid[17].length).toBe(17); // A..Q, no trailing empty columns
  });
  it('labels in column A match the template row by row', () => {
    for (let r = 16; r <= 73; r++) {
      const exp = template.sheets[S.BOM].cells[`A${r}`]?.v;
      const norm = (x) => String(x).replace(/for years.*$/, '').replace(/\s+/g, ' ').trim().toLowerCase();
      expect(norm(grid[r][0]), `row ${r}`).toBe(norm(exp));
    }
  });
  it('converts BoM dates and keeps values in place', () => {
    const d = new Date(Date.UTC(1899, 11, 30) + grid[19][1] * 86400000);
    expect(d.toISOString().slice(0, 10)).toBe('2026-01-27');
    const d2 = new Date(Date.UTC(1899, 11, 30) + grid[19][2] * 86400000);
    expect(d2.toISOString().slice(0, 10)).toBe('2009-02-07');
    expect(grid[17][1]).toBe(25.8);
    expect(grid[47][1]).toBe(89.4);
  });
  it('drives the Climate sheet', () => {
    const { wb } = buildWorkbook({ bomGrid: grid });
    expect(wb.getValue(S.CLIMATE, 'A1')).toBe("Monthly Climate Statistics for 'LAVERTON RAAF' [087031]");
    expect(wb.getValue(S.CLIMATE, 'B4')).toBe(25.8);   // mean max temp Jan
    expect(wb.getValue(S.CLIMATE, 'B16')).toBe(5.2);   // lowest temp Jan
    expect(wb.getValue(S.CLIMATE, 'B28')).toBe(89.4);  // decile 9 rainfall Jan
    expect(wb.getValue(S.CLIMATE, 'B29')).toBe(40.8);  // mean rainfall Jan
    expect(wb.getValue(S.CLIMATE, 'B55')).toBe(15.9);  // 9am wind Jan
    expect(wb.getValue(S.CLIMATE, 'B62')).toBe(22.9);  // 3pm wind Jan
    expect(wb.getValue(S.CSWS, 'H37')).toBe('Weather Data from BoM Weather Station 087031 LAVERTON RAAF');
  });
});
