import React, { useMemo, useState } from 'react';
import { formatValue } from '../engine/format.js';
import { isErr } from '../engine/formulaEngine.js';
import { numToCol } from '../engine/formulaEngine.js';

const colPx = (w) => Math.round(w * 7 + 5);          // Excel column width -> px (Calibri 11)
const rowPx = (pt) => Math.round((pt * 96) / 72 * 100) / 100; // points -> px
const BORDER_W = { hair: 1, thin: 1, dotted: 1, dashed: 1, medium: 2, thick: 3, double: 3 };
const BORDER_STYLE = { hair: 'dotted', dotted: 'dotted', dashed: 'dashed', double: 'double' };

function cssBorder(b) {
  if (!b) return undefined;
  return `${BORDER_W[b.s] || 1}px ${BORDER_STYLE[b.s] || 'solid'} ${b.c}`;
}

/** Parse what the user typed. Percent-formatted cells follow Excel's rule: 10 and 0.1 both mean 10%. */
function parseInput(text, isPercent) {
  const t = text.trim();
  if (t === '') return null;
  const hasPct = t.endsWith('%');
  const n = Number(hasPct ? t.slice(0, -1).trim() : t);
  if (!isFinite(n) || n < 0) return null;
  if (!isPercent) return n;
  return hasPct || n >= 1 ? n / 100 : n;
}

function EditableInput({ value, isPercent, onCommit, title }) {
  const shown = isPercent ? `${Math.round(value * 1e6) / 1e4}%` : String(value);
  const [text, setText] = useState(null); // null = not editing
  const commit = () => {
    if (text === null) return;
    const v = parseInput(text, isPercent);
    setText(null);
    if (v !== null && v !== value) onCommit(v);
  };
  return (
    <input
      className="cell-input"
      inputMode="decimal"
      title={title}
      value={text ?? shown}
      onFocus={(e) => { setText(isPercent ? String(Math.round(value * 1e6) / 1e4) : String(value)); requestAnimationFrame(() => e.target.select()); }}
      onChange={(e) => setText(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === 'Enter') e.currentTarget.blur();
        if (e.key === 'Escape') { setText(null); e.currentTarget.blur(); }
      }}
    />
  );
}

/**
 * Renders a sheet exactly as laid out in the Excel workbook (merged cells, column widths,
 * row heights, fills, fonts, borders, number formats) with live formula results.
 */
export default function SheetView({ spec, sheetName, workbook, editable = new Set(), onEdit }) {
  const model = useMemo(() => {
    const cellMap = new Map();
    for (const c of spec.cells) cellMap.set(`${c.r},${c.c}`, c);
    const anchors = new Map();
    const covered = new Set();
    for (const m of spec.merges) {
      const [a, b] = m.split(':');
      const pa = /^([A-Z]+)(\d+)$/.exec(a), pb = /^([A-Z]+)(\d+)$/.exec(b);
      const toN = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      const r1 = +pa[2], c1 = toN(pa[1]), r2 = +pb[2], c2 = toN(pb[1]);
      anchors.set(`${r1},${c1}`, { r2, c2, rowSpan: r2 - r1 + 1, colSpan: c2 - c1 + 1 });
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered.add(`${r},${c}`);
    }
    // cumulative top offsets for frozen (sticky) rows
    const tops = [0];
    for (let r = 1; r <= spec.freezeRows; r++) tops.push(tops[r - 1] + rowPx(spec.rows[r]));
    return { cellMap, anchors, covered, tops };
  }, [spec]);

  const styleAt = (r, c) => {
    const cell = model.cellMap.get(`${r},${c}`);
    return cell ? spec.styles[cell.s] : undefined;
  };

  const rows = [];
  for (let r = 1; r <= spec.maxRow; r++) {
    const tds = [];
    for (let c = 1; c <= spec.maxCol; c++) {
      const key = `${r},${c}`;
      if (model.covered.has(key)) continue;
      const cell = model.cellMap.get(key);
      const st = (cell && spec.styles[cell.s]) || {};
      const merge = model.anchors.get(key);
      const a1 = `${numToCol(c)}${r}`;

      // borders (merged areas take their outer edges from the far cells)
      const b = { ...(st.b || {}) };
      if (merge) {
        const rightSt = styleAt(r, merge.c2)?.b?.right;
        const bottomSt = styleAt(merge.r2, c)?.b?.bottom || styleAt(merge.r2, merge.c2)?.b?.bottom;
        if (rightSt) b.right = rightSt;
        if (bottomSt) b.bottom = bottomSt;
      }

      // value
      let display = '';
      let raw = null;
      if (cell && (cell.f !== undefined || cell.v !== undefined)) {
        raw = workbook.getValue(sheetName, r, c);
        display = formatValue(raw, st.nf);
      }

      const font = st.f || {};
      const al = st.a || {};
      const isNum = typeof raw === 'number';
      // Excel lets unwrapped left-aligned text run over empty neighbouring cells; do the same
      const nextCell = model.cellMap.get(`${r},${c + (merge ? merge.colSpan : 1)}`);
      const nextEmpty = !nextCell || (nextCell.f === undefined && nextCell.v === undefined);
      const spill =
        typeof raw === 'string' && display.trim() !== '' && !al.w && (!al.h || al.h === 'left' || al.h === 'general') && nextEmpty &&
        c + (merge ? merge.colSpan : 1) <= spec.maxCol;
      const style = {
        fontFamily: font.n ? `${font.n}, Calibri, Arial, sans-serif` : undefined,
        fontSize: font.z ? `${Math.round((font.z * 96) / 72 * 100) / 100}px` : undefined,
        fontWeight: font.b ? 700 : undefined,
        fontStyle: font.i ? 'italic' : undefined,
        textDecoration: font.u ? 'underline' : undefined,
        color: font.c,
        background: st.bg,
        textAlign: al.h && al.h !== 'general' ? (al.h === 'centerContinuous' ? 'center' : al.h) : isErr(raw) ? 'center' : isNum ? 'right' : 'left',
        verticalAlign: al.v === 'center' ? 'middle' : al.v === 'top' ? 'top' : 'bottom',
        whiteSpace: al.w ? 'pre-wrap' : 'nowrap',
        overflow: spill ? 'visible' : undefined,
        paddingLeft: al.i ? `${al.i * 9 + 2}px` : undefined,
        borderLeft: cssBorder(b.left),
        borderRight: cssBorder(b.right),
        borderTop: cssBorder(b.top),
        borderBottom: cssBorder(b.bottom),
      };
      if (r <= spec.freezeRows) {
        style.position = 'sticky';
        style.top = model.tops[r - 1];
        style.zIndex = spill ? 4 : 3; // spilled text must stay above the sticky neighbours it runs over
        if (!style.background) style.background = '#fff';
      }
      const isEditable = editable.has(a1) && typeof raw === 'number';
      tds.push(
        <td
          key={key}
          rowSpan={merge?.rowSpan}
          colSpan={merge?.colSpan}
          style={style}
          className={isEditable ? 'editable' : undefined}
        >
          {isEditable ? (
            <EditableInput
              value={raw}
              isPercent={/%$/.test(st.nf || '')}
              title="Editable input – recalculates the whole sheet"
              onCommit={(v) => onEdit(a1, v)}
            />
          ) : spill ? (
            <span className="spill">{display}</span>
          ) : (
            display
          )}
        </td>,
      );
    }
    rows.push(<tr key={r} style={{ height: rowPx(spec.rows[r]) }}>{tds}</tr>);
  }

  const totalW = Array.from({ length: spec.maxCol }, (_, i) => colPx(spec.cols[i + 1])).reduce((a, b) => a + b, 0);

  return (
    <div className="sheet-scroll">
      <table
        className={`sheet ${spec.showGrid ? 'grid' : ''}`}
        style={{ width: totalW, zoom: spec.zoom / 100 }}
      >
        <colgroup>
          {Array.from({ length: spec.maxCol }, (_, i) => (
            <col key={i} style={{ width: colPx(spec.cols[i + 1]) }} />
          ))}
        </colgroup>
        <tbody>{rows}</tbody>
      </table>
    </div>
  );
}
