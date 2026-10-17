// Export one sheet of the live workbook to .xlsx with the template's formatting and
// the current calculated values.
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { parseAddr } from '../engine/address.js';
import { isError } from '../engine/values.js';

const argb = (hex) => (hex ? { argb: `FF${hex}` } : undefined);

function excelStyle(st) {
  if (!st) return {};
  const f = st.font || {};
  const out = {
    font: {
      name: f.name || 'Calibri', size: f.sz || 11, bold: f.b || undefined, italic: f.i || undefined,
      underline: f.u || undefined, strike: f.strike || undefined, color: argb(f.color),
    },
    alignment: {
      horizontal: st.align?.h || undefined, vertical: st.align?.v || undefined,
      wrapText: st.align?.wrap || undefined, indent: st.align?.indent || undefined,
      textRotation: st.align?.rot || undefined, shrinkToFit: st.align?.shrink || undefined,
    },
    numFmt: st.numFmt || 'General',
  };
  if (st.fill) out.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(st.fill), bgColor: argb(st.fill) };
  const b = st.border || {};
  const side = (k) => (b[k] ? { style: b[k].style, color: argb(b[k].color) } : undefined);
  if (Object.keys(b).length) out.border = { left: side('l'), right: side('r'), top: side('t'), bottom: side('b') };
  return out;
}

function cfRule(rule) {
  if (rule.type === 'colorScale') {
    return {
      type: 'colorScale', priority: rule.priority,
      cfvo: rule.cfvo.map((v) => (v.val === null || v.val === undefined ? { type: v.type } : { type: v.type, value: Number(v.val) })),
      color: rule.colors.map((c) => argb(c)),
    };
  }
  const style = {};
  if (rule.dxf?.fill) style.fill = { type: 'pattern', pattern: 'solid', bgColor: argb(rule.dxf.fill) };
  if (rule.dxf?.color || rule.dxf?.b) style.font = { color: argb(rule.dxf.color), bold: rule.dxf.b || undefined };
  if (rule.type === 'cellIs') {
    return { type: 'cellIs', operator: rule.operator, formulae: rule.formula, style, priority: rule.priority };
  }
  if (rule.type === 'expression') return { type: 'expression', formulae: rule.formula, style, priority: rule.priority };
  return null;
}

export async function exportSheet(wb, sheetName, fileName) {
  const s = wb.sheet(sheetName);
  const book = new ExcelJS.Workbook();
  book.creator = 'Inclement Weather Analysis';
  const ws = book.addWorksheet(sheetName.slice(0, 31), {
    views: [{ zoomScale: s.zoom || 100, showGridLines: s.showGrid }],
    properties: { defaultRowHeight: s.defaultRowHeight || 15 },
  });
  for (let c = 1; c <= s.maxCol; c++) {
    const d = s.cols[c];
    const col = ws.getColumn(c);
    col.width = d?.w ?? s.defaultColWidth ?? 8.43;
    if (d?.hidden) col.hidden = true;
  }
  for (const [r, d] of Object.entries(s.rows)) {
    const row = ws.getRow(Number(r));
    if (d.h) row.height = d.h;
    if (d.hidden) row.hidden = true;
  }
  for (const [a, cell] of s.cells) {
    const { row, col } = parseAddr(a);
    const xc = ws.getCell(row, col);
    let v = wb.getValue(sheetName, a);
    if (isError(v)) v = { error: v.err };
    else if (cell.f !== undefined && v === '') v = null;
    if (v !== null && v !== undefined) xc.value = v;
    const st = wb.styles[cell.s];
    if (st) Object.assign(xc, excelStyle(st));
    if (cell.note) xc.note = cell.note;
  }
  for (const m of s.merges) {
    try {
      ws.mergeCells(m);
    } catch {
      /* overlapping merge in source - ignore */
    }
  }
  for (const rule of s.cf) {
    const r = cfRule(rule);
    if (r) ws.addConditionalFormatting({ ref: rule.sqref, rules: [r] });
  }
  let buf = await book.xlsx.writeBuffer();
  if (!s.showZeros) {
    // ExcelJS has no showZeros view option; patch the sheet XML so zero values stay hidden like the template.
    const zip = await JSZip.loadAsync(buf);
    const path = 'xl/worksheets/sheet1.xml';
    const xml = await zip.file(path).async('string');
    zip.file(path, xml.replace('<sheetView ', '<sheetView showZeros="0" '));
    buf = await zip.generateAsync({ type: 'arraybuffer' });
  }
  if (typeof document === 'undefined') return buf;
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  return buf;
}
