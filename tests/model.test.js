import { describe, it, expect } from 'vitest';
import template from '../src/template/workbook.json';
import { buildWorkbook as build, defaultCategories, newCategory, S } from '../src/model/build.js';

// The original workbook's BoM data (Moorabbin) - the app itself starts with no data.
const bomGrid = {};
for (const [a, c] of Object.entries(template.sheets['BOM csv data'].cells)) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  if (+m[2] < 5 || c.v === undefined) continue;
  (bomGrid[m[2]] ||= [])[m[1].charCodeAt(0) - 65] = c.v;
}
const buildWorkbook = (state) => build({ bomGrid, ...state });

describe('categories', () => {
  it('default build matches template', () => {
    const { wb } = buildWorkbook({});
    expect(wb.getValue(S.CSWS, 'D51')).toBeCloseTo(78.01550576036867, 9);
    expect(wb.getValue(S.WC, 'B69')).toBe('RAIN EXPOSURE 1 : Bulk + Inground Works');
  });
  it('starts empty, without instructions, planner name or formula edits', () => {
    const { wb } = build({ overrides: { [`${S.CSWS}|D51`]: '5', [`${S.CAL}|I32`]: '=1+1', [`${S.CAL}|J32`]: '2' } });
    expect(wb.getValue(S.CLIMATE, 'B4')).toBe(0);
    expect(wb.getCell(S.BOM, 'A17')?.v).toBeUndefined();
    expect([...wb.sheet(S.WC).cells.keys()].filter((a) => /^[UV]\d/.test(a))).toEqual([]);
    expect(wb.getValue(S.CSWS, 'B5')).toBeNull();
    expect(wb.getCell(S.CSWS, 'D51').f).toBe('=SUM(D42:D50)');
    expect(wb.getValue(S.CAL, 'I32')).toBe(0);
    expect(wb.getValue(S.CAL, 'J32')).toBe(2);
  });
  it('adding categories keeps other lookups intact and feeds new rows', () => {
    const cats = [...defaultCategories(), newCategory(6), newCategory(7)];
    cats[5].direct = [0, 1, 1, 1];
    cats[6].direct = [0, 2, 2, 2];
    const { wb, layout } = buildWorkbook({ categories: cats });
    expect(layout.title).toBe(81);
    expect(wb.getValue(S.WC, `E${layout.title}`)).toBe('Weather Analysis Impact Table');
    expect(wb.getValue(S.WC, `F${layout.exposureRows[6]}`)).toBe(7);
    expect(wb.getValue(S.WC, `F${layout.exposureRows[6] + 1}`)).toBe('Heat');
    // unchanged totals
    expect(wb.getValue(S.CSWS, 'D43')).toBeCloseTo(4.498617511520737, 9);
    expect(wb.getValue(S.CSWS, 'D45')).toBeCloseTo(8, 9);
    expect(wb.getValue(S.CSWS, 'D51')).toBeCloseTo(78.01550576036867, 9);
    console.log(wb.getCell(S.CSWS, 'I43').f, wb.getCell(S.CSWS, 'I42').f);
    // new exposure 7 = 2x exposure 6
    const r6 = wb.getValue(S.WC, `D${layout.exposureRows[5]}`);
    const r7 = wb.getValue(S.WC, `D${layout.exposureRows[6]}`);
    expect(r6).toBeGreaterThan(0);
    expect(r7).toBeCloseTo(2 * r6, 9);
    expect(wb.getValue(S.RAIN, 'C26')).toBe(0);
    expect(wb.getValue(S.RAIN, 'D26')).toBe(2);
  });
  it('hot weather factor flows to Critical Site Summary', () => {
    const { wb } = buildWorkbook({ overrides: { [`${S.WC}|E5`]: '0.8' } });
    expect(wb.getValue(S.CSWS, 'D43')).toBeCloseTo(2 * 4.498617511520737, 9);
    expect(wb.getValue(S.CSWS, 'C43')).toBe(0.8);
  });
});
