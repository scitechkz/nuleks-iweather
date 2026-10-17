import { describe, it, expect } from 'vitest';
import template from '../src/template/workbook.json';
import { Workbook } from '../src/engine/workbook.js';
import { isError } from '../src/engine/values.js';

function mismatches(wb) {
  const bad = [];
  let n = 0;
  for (const [name, s] of Object.entries(wb.sheets)) {
    for (const [a, c] of s.cells) {
      if (c.f === undefined || c.f.includes('TODAY')) continue;
      n++;
      const exp = template.sheets[name].cells[a]?.cv ?? null;
      let got = wb.getValue(name, a);
      if (isError(got)) got = got.err;
      if (got === '' && exp === null) continue;
      if (typeof exp === 'number' && typeof got === 'number' && Math.abs(exp - got) < 1e-9 * Math.max(1, Math.abs(exp))) continue;
      if (exp === got) continue;
      if ((exp === null && got === 0)) continue;
      bad.push(`${name}!${a} ${c.f} expected=${JSON.stringify(exp)} got=${JSON.stringify(got)}`);
    }
  }
  return { bad, n };
}

describe('engine matches Excel cached values', () => {
  it('evaluates every formula like Excel', () => {
    const wb = new Workbook(template);
    const { bad, n } = mismatches(wb);
    console.log(`checked ${n} formulas, ${bad.length} mismatches`);
    expect(bad.slice(0, 30)).toEqual([]);
  });
});
