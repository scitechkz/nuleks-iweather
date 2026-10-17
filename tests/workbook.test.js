import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import { parseBomCsv } from '../src/engine/csvLoader.js';
import { buildWorkbook } from '../src/engine/buildWorkbook.js';
import { isErr } from '../src/engine/formulaEngine.js';
import expected from './excel_cached_values.json';

const csv = fs.readFileSync(new URL('../public/sample_BOM_Dump.csv', import.meta.url), 'utf8');

describe('formula engine vs. Excel cached values', () => {
  const { rows, warnings } = parseBomCsv(csv);
  const wb = buildWorkbook(rows);

  it('CSV aligns with the workbook BOM Dump layout (no warnings)', () => {
    expect(warnings).toEqual([]);
  });

  for (const sheet of ['BOM Extract', 'Weather Analysis']) {
    it(`every formula on "${sheet}" matches Excel`, () => {
      const bad = [];
      let n = 0;
      for (const [key, exp] of Object.entries(expected[sheet])) {
        const [r, c] = key.split(',').map(Number);
        const got = wb.getValue(sheet, r, c);
        n++;
        if (isErr(got)) { bad.push(`${key}: ${got.code}`); continue; }
        if (typeof exp === 'number') {
          if (typeof got !== 'number' || Math.abs(got - exp) > 1e-9 * Math.max(1, Math.abs(exp))) bad.push(`${key}: got ${got}, Excel ${exp}`);
        } else if ((exp ?? 0) !== got) bad.push(`${key}: got ${JSON.stringify(got)}, Excel ${JSON.stringify(exp)}`);
      }
      expect(bad).toEqual([]);
      expect(n).toBeGreaterThan(50);
    });
  }

  it('editing a direct impact flows through to totals', () => {
    const wb2 = buildWorkbook(rows, { E7: 0.8 });           // hot weather 40% -> 80%
    const base = wb.getValue('Weather Analysis', 7, 6);      // F7
    const doubled = wb2.getValue('Weather Analysis', 7, 6);
    expect(doubled).toBeCloseTo(base * 2, 9);
    const wb3 = buildWorkbook(rows, { B31: 1 });             // B31 0.5 -> 1  => D31 1.5 -> 2
    expect(wb3.getValue('Weather Analysis', 31, 4)).toBe(2);
  });
});

describe('CSV loader robustness', () => {
  it('realigns rows when extra lines precede the header', () => {
    const shifted = '\n\n' + csv;
    const { rows, warnings } = parseBomCsv(shifted);
    expect(String(rows[12][0]).toLowerCase()).toContain('statistic element');
    expect(warnings.some((w) => w.includes('realigned'))).toBe(true);
    expect(warnings.filter((w) => w.startsWith('BOM Dump row'))).toEqual([]);
  });

  it('warns when a statistic row is missing/misplaced', () => {
    const lines = csv.split('\n');
    lines.splice(21, 1); // drop the ">= 35 °C" row
    const { warnings } = parseBomCsv(lines.join('\n'));
    expect(warnings.some((w) => w.startsWith('BOM Dump row'))).toBe(true);
  });

  it('rejects a file without the Statistic Element header', () => {
    expect(() => parseBomCsv('a,b,c\n1,2,3')).toThrow(/Statistic Element/);
  });

  it('handles quoted fields and the mis-decoded degree sign', () => {
    const { rows } = parseBomCsv(csv.replace('Mean maximum temperature (°C) for years 1943 to 2026 ', '"Mean maximum temperature (Â°C) for years 1943 to 2026 "'));
    expect(rows[13][0]).toContain('(°C)');
    expect(rows[13][1]).toBe(25.8);
  });
});
