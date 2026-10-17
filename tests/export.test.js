import { it, expect } from 'vitest';
import fs from 'node:fs';
import { buildWorkbook, S } from '../src/model/build.js';
import { exportSheet } from '../src/export/exportSheet.js';

it('exports Weather Calculation', async () => {
  const { wb } = buildWorkbook({});
  const buf = await exportSheet(wb, S.WC, 'x.xlsx');
  fs.mkdirSync('tests/out', { recursive: true });
  fs.writeFileSync('tests/out/weather.xlsx', Buffer.from(buf));
  expect(buf.byteLength).toBeGreaterThan(5000);
});
