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

describe('additional rain exposure rating rows', () => {
  it('no extra rows leaves the template unchanged', () => {
    const { wb } = buildWorkbook({ extraRatingRows: 0 });
    expect(wb.getCell(S.CSWS, 'I65')).toBeUndefined();
    expect(wb.getValue(S.CSWS, 'D51')).toBeCloseTo(78.01550576036867, 9);
  });
  it('an extra row adds its wet-weather days to the delay allowance', () => {
    const base = buildWorkbook({});
    const blank = buildWorkbook({ extraRatingRows: 1 });
    expect(blank.wb.getValue(S.CSWS, 'A65')).toBe('Additional Rain Exposure Rating 2');
    expect(blank.wb.getValue(S.CSWS, 'D42')).toBeCloseTo(base.wb.getValue(S.CSWS, 'D42'), 9);
    // Rating row 65 the same as row 11 in one period doubles that period's wet weather days
    const { wb } = buildWorkbook({ extraRatingRows: 1, overrides: { [`${S.CSWS}|J65`]: '3' } });
    expect(wb.getValue(S.CSWS, 'J42')).toBeCloseTo(2 * base.wb.getValue(S.CSWS, 'J42'), 9);
    expect(wb.getValue(S.CSWS, 'D51') - base.wb.getValue(S.CSWS, 'D51')).toBeCloseTo(base.wb.getValue(S.CSWS, 'J42'), 9);
  });
  it('extra rows keep working with added categories', () => {
    const cats = [...defaultCategories(), newCategory(6)];
    cats[5].direct = [0, 1, 1, 1];
    const { wb } = buildWorkbook({ categories: cats, extraRatingRows: 1, overrides: { [`${S.CSWS}|J65`]: '6' } });
    expect(wb.getValue(S.CSWS, 'J65')).toBeNull(); // ratings stay limited to 1-5
    const ok = buildWorkbook({ categories: cats, extraRatingRows: 1, overrides: { [`${S.CSWS}|J65`]: '5' } });
    expect(ok.wb.getCell(S.CSWS, 'J42').f).toMatch(/J65/);
    expect(typeof ok.wb.getValue(S.CSWS, 'J42')).toBe('number');
  });
});
