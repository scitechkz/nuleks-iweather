import { useMemo, useRef, useState, useEffect } from 'react';
import { addr, colToLetters, parseRange } from '../engine/address.js';
import { formatValue } from '../engine/format.js';
import { isError } from '../engine/values.js';
import { computeConditionalFormats } from './conditionalFormat.js';

const colPx = (w) => Math.round((w ?? 8.43) * 7 + 5);
const rowPx = (h) => Math.round(((h ?? 15) * 4) / 3);

const BORDER = {
  thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed',
  medium: '2px solid', mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed',
  dashDotDot: '1px dotted', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed',
  thick: '3px solid', double: '3px double',
};

function cssFor(style, cf, value, showGrid) {
  const css = {};
  if (!style) return css;
  const f = style.font || {};
  css.fontFamily = `"${f.name || 'Calibri'}", Calibri, Carlito, Arial, sans-serif`;
  css.fontSize = `${f.sz || 11}pt`;
  if (f.b || cf?.b) css.fontWeight = 700;
  if (f.i) css.fontStyle = 'italic';
  const deco = [];
  if (f.u) deco.push('underline');
  if (f.strike) deco.push('line-through');
  if (deco.length) css.textDecoration = deco.join(' ');
  const color = cf?.color || f.color;
  if (color) css.color = `#${color}`;
  const fill = cf?.fill || style.fill;
  if (fill) css.background = `#${fill}`;
  const b = style.border || {};
  const side = (k, prop) => {
    if (b[k]) css[prop] = `${BORDER[b[k].style] || '1px solid'} #${b[k].color || '000000'}`;
    else if (showGrid && !fill) css[prop] = '1px solid var(--grid)';
  };
  side('l', 'borderLeft'); side('r', 'borderRight'); side('t', 'borderTop'); side('b', 'borderBottom');
  const a = style.align || {};
  let h = a.h;
  if (!h || h === 'general') {
    h = typeof value === 'number' ? 'right' : typeof value === 'boolean' || isError(value) ? 'center' : 'left';
  }
  if (h === 'centerContinuous') h = 'center';
  if (h === 'fill' || h === 'justify' || h === 'distributed') h = 'left';
  css.textAlign = h;
  css.verticalAlign = a.v === 'center' ? 'middle' : a.v === 'top' ? 'top' : 'bottom';
  css.whiteSpace = a.wrap ? 'pre-wrap' : 'pre';
  if (a.indent) css[h === 'right' ? 'paddingRight' : 'paddingLeft'] = `${a.indent * 9 + 2}px`;
  return css;
}

/**
 * Renders an Excel sheet (or part of one) with the template's styles.
 * props:
 *  wb, sheet, visibleRows?: (row)=>bool, cols?: [c1,c2], rows?: [r1,r2]
 *  editable?: bool | (addr)=>bool, onEdit(addr, raw)
 *  showHidden?: bool  - reveal rows/cols hidden in the workbook
 *  headers?: bool     - show column letters / row numbers (off by default)
 */
export default function SheetGrid({
  wb, sheet, version, rows, cols, visibleRows, editable = false, onEdit, showHidden = false,
  headers = false, shapes = false, zoom,
}) {
  const s = wb.sheet(sheet);
  const r1 = rows?.[0] ?? 1;
  const r2 = rows?.[1] ?? s.maxRow;
  const c1 = cols?.[0] ?? 1;
  const c2 = cols?.[1] ?? s.maxCol;
  const [sel, setSel] = useState(null);
  const [editing, setEditing] = useState(null); // {a, text}
  const inputRef = useRef(null);

  const cf = useMemo(() => computeConditionalFormats(wb, sheet), [wb, sheet, version]);

  const rowList = [];
  for (let r = r1; r <= r2; r++) {
    const hidden = (s.rows[r]?.hidden && !showHidden) || (visibleRows && !visibleRows(r));
    rowList.push({ r, hidden });
  }
  const colList = [];
  for (let c = c1; c <= c2; c++) {
    const d = s.cols[c];
    if (d?.hidden && !showHidden) continue;
    colList.push({ c, w: colPx(d?.w ?? s.defaultColWidth) });
  }
  const colSet = new Set(colList.map((x) => x.c));
  const hiddenRow = new Set(rowList.filter((x) => x.hidden).map((x) => x.r));

  // merges: top-left -> span, covered -> skip
  const { spans, covered } = useMemo(() => {
    const sp = new Map();
    const cov = new Set();
    for (const m of s.merges) {
      const g = parseRange(m);
      let rs = 0;
      for (let r = g.r1; r <= g.r2; r++) if (!hiddenRow.has(r) && r >= r1 && r <= r2) rs++;
      let cs = 0;
      for (let c = g.c1; c <= g.c2; c++) if (colSet.has(c)) cs++;
      // anchor = first visible cell of the merge
      let ar = g.r1;
      while (ar <= g.r2 && (hiddenRow.has(ar) || ar < r1)) ar++;
      let ac = g.c1;
      while (ac <= g.c2 && !colSet.has(ac)) ac++;
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) cov.add(addr(r, c));
      if (ar <= g.r2 && ac <= g.c2 && rs && cs) {
        cov.delete(addr(ar, ac));
        sp.set(addr(ar, ac), { rowSpan: rs, colSpan: cs, src: addr(g.r1, g.c1) });
      }
    }
    return { spans: sp, covered: cov };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [s.merges, r1, r2, c1, c2, showHidden, visibleRows, version]);

  // Formula cells are always read-only; only plain values can be edited.
  const canEdit = (a) => wb.getCell(sheet, a)?.f === undefined
    && (typeof editable === 'function' ? editable(a) : editable);

  useEffect(() => {
    if (editing && inputRef.current) inputRef.current.focus();
  }, [editing]);

  const rawOf = (a) => {
    const cell = wb.getCell(sheet, a);
    if (!cell) return '';
    if (cell.f !== undefined) return '';
    if (cell.v === undefined || cell.v === null) return '';
    const st = wb.styles[cell.s];
    if (typeof cell.v === 'number' && st && /[dmy]/.test(st.numFmt.replace(/"[^"]*"/g, '')) && st.numFmt !== 'General') {
      return formatValue(cell.v, 'dd/mm/yyyy');
    }
    return String(cell.v);
  };

  const commit = (a, text) => {
    setEditing(null);
    if (text === rawOf(a) || text.trim().startsWith('=') || !canEdit(a)) return;
    let raw = text;
    const dm = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text.trim());
    if (dm) raw = String((Date.UTC(+dm[3], +dm[2] - 1, +dm[1]) - Date.UTC(1899, 11, 30)) / 86400000);
    onEdit?.(a, raw);
  };

  const showZeros = s.showZeros;
  const showGrid = false; // no spreadsheet gridlines - only the designed borders
  const totalWidth = colList.reduce((t, x) => t + x.w, 0) + (headers ? 40 : 0);

  // pixel offsets for shapes
  const shapeEls = [];
  if (shapes && s.shapes?.length) {
    const xOf = (c0, off) => {
      let x = headers ? 40 : 0;
      for (const { c, w } of colList) {
        if (c - 1 >= c0) break;
        x += w;
      }
      return x + off / 9525;
    };
    const yOf = (r0, off) => {
      let y = headers ? 20 : 0;
      for (const { r, hidden } of rowList) {
        if (r - 1 >= r0) break;
        if (!hidden) y += rowPx(s.rows[r]?.h ?? s.defaultRowHeight);
      }
      return y + off / 9525;
    };
    s.shapes.forEach((sh, i) => {
      if (hiddenRow.has(sh.from[2] + 1) && hiddenRow.has(sh.to[2] + 1)) return;
      if (sh.from[2] + 1 < r1 || sh.to[2] + 1 > r2) return;
      const x1 = xOf(sh.from[0], sh.from[1]);
      const y1 = yOf(sh.from[2], sh.from[3]);
      const x2 = xOf(sh.to[0], sh.to[1]);
      const y2 = yOf(sh.to[2], sh.to[3]);
      const w = Math.max(2, x2 - x1);
      const h = Math.max(2, y2 - y1);
      let text = sh.text;
      if (sh.link) {
        const v = wb.getValue(sheet, sh.link);
        text = formatValue(v, wb.styles[wb.getCell(sheet, sh.link)?.s]?.numFmt || 'General');
      }
      if (sh.geom === 'rect') {
        shapeEls.push(
          <div key={i} className="xl-shape" style={{ left: x1, top: y1, width: w, height: h }}>{text}</div>,
        );
      } else if (sh.geom === 'leftRightArrow') {
        shapeEls.push(
          <svg key={i} className="xl-arrow" style={{ left: x1, top: y1 }} width={w} height={h} viewBox={`0 0 ${w} ${h}`}>
            <polygon fill="#5B9BD5" stroke="#41719C"
              points={`0,${h / 2} ${h},0 ${h},${h / 4} ${w - h},${h / 4} ${w - h},0 ${w},${h / 2} ${w - h},${h} ${w - h},${(3 * h) / 4} ${h},${(3 * h) / 4} ${h},${h}`} />
          </svg>,
        );
      }
    });
  }

  return (
    <div className="sheet-wrap">
      <div className="sheet-scroll">
        <div className="sheet-zoom" style={{ zoom: (zoom ?? s.zoom ?? 100) / 100 }}>
          <div style={{ position: 'relative', width: totalWidth }}>
            <table className="xl" style={{ width: totalWidth }}>
              <colgroup>
                {headers && <col style={{ width: 40 }} />}
                {colList.map(({ c, w }) => <col key={c} style={{ width: w }} />)}
              </colgroup>
              {headers && (
                <thead>
                  <tr style={{ height: 20 }}>
                    <th className="xl-corner" />
                    {colList.map(({ c }) => <th key={c} className="xl-colhead">{colToLetters(c)}</th>)}
                  </tr>
                </thead>
              )}
              <tbody>
                {rowList.filter((x) => !x.hidden).map(({ r }) => {
                  const h = rowPx(s.rows[r]?.h ?? s.defaultRowHeight);
                  return (
                    <tr key={r} style={{ height: h }}>
                      {headers && <th className="xl-rowhead">{r}</th>}
                      {colList.map(({ c }, ci) => {
                        const a = addr(r, c);
                        if (covered.has(a)) return null;
                        const span = spans.get(a);
                        const src = span ? span.src : a;
                        const cell = s.cells.get(src);
                        const st = cell ? wb.styles[cell.s] : null;
                        const v = cell ? wb.getValue(sheet, src) : null;
                        const text = st ? formatValue(v, st.numFmt, { showZeros }) : formatValue(v, 'General', { showZeros });
                        const css = cssFor(st || wb.styles[0], cf.get(src), v, showGrid);
                        if (!st && showGrid) {
                          css.borderRight = '1px solid var(--grid)';
                          css.borderBottom = '1px solid var(--grid)';
                        }
                        // Let unwrapped text spill into empty neighbours like Excel.
                        let overflow = 'hidden';
                        if (text && !(st?.align?.wrap) && !span) {
                          const next = colList[ci + 1];
                          const nextEmpty = !next || wb.getValue(sheet, addr(r, next.c)) === null;
                          if (css.textAlign === 'left' && nextEmpty && typeof v === 'string') overflow = 'visible';
                        }
                        const isSel = sel === a;
                        const isEdit = editing?.a === a;
                        const ed = canEdit(src);
                        return (
                          <td
                            key={c}
                            data-addr={src}
                            rowSpan={span?.rowSpan}
                            colSpan={span?.colSpan}
                            style={{ ...css, overflow, maxWidth: overflow === 'hidden' ? 0 : undefined }}
                            className={`${isSel ? 'xl-sel' : ''} ${ed ? 'xl-editable' : ''}`}
                            title={cell?.note || undefined}
                            onClick={() => setSel(src)}
                            onDoubleClick={() => ed && setEditing({ a: src, text: rawOf(src) })}
                            onKeyDown={(e) => {
                              if (!ed || isEdit) return;
                              if (e.key === 'Enter' || e.key === 'F2') setEditing({ a: src, text: rawOf(src) });
                              else if (e.key === 'Delete' || e.key === 'Backspace') onEdit?.(src, '');
                              else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) setEditing({ a: src, text: e.key });
                            }}
                            tabIndex={ed ? 0 : -1}
                          >
                            {isEdit ? (
                              <input
                                ref={inputRef}
                                className="xl-input"
                                defaultValue={editing.text}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') commit(src, e.currentTarget.value);
                                  if (e.key === 'Escape') setEditing(null);
                                }}
                                onBlur={(e) => commit(src, e.currentTarget.value)}
                              />
                            ) : text}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {shapeEls}
          </div>
        </div>
      </div>
    </div>
  );
}
