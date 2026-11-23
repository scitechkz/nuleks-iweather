import { useState } from 'react';
import { addr } from '../engine/address.js';
import { formatValue } from '../engine/format.js';
import {
  S, inputError, RATING_MIN, RATING_MAX, RATING_ROW, RATING_FIRST_COL, RATING_LAST_COL,
  extraRatingRow, MAX_EXTRA_RATING_ROWS,
} from '../model/build.js';

// Critical Site Works Summary: the Period row (row 9), the Rain Exposure Rating row (row 11) and any
// additional rating rows (row 65 onwards) across every period column I..CZ, in the same table style
// as the other pages. Every rating row adds to the Wet Weather delay allowance.
const RATINGS = Array.from({ length: RATING_MAX - RATING_MIN + 1 }, (_, i) => String(RATING_MIN + i));
const SERIES = ['var(--series-1)', 'var(--series-2)', 'var(--series-rain)'];

function readPeriods(wb) {
  const periods = [];
  for (let c = RATING_FIRST_COL; c <= RATING_LAST_COL; c++) {
    const serial = wb.getValue(S.CSWS, addr(9, c));
    if (typeof serial === 'number') periods.push({ c, label: formatValue(serial, 'mmm-yy') });
  }
  return periods;
}

function readRow(wb, overrides, row, periods) {
  return periods.map(({ c }) => {
    const a = addr(row, c);
    const rating = wb.getValue(S.CSWS, a);
    return {
      a,
      rating: rating === null || rating === undefined ? '' : String(rating),
      edited: `${S.CSWS}|${a}` in overrides,
    };
  });
}

// Activity rows 12-22 (labels in column A; the rows between them are spacers): read-only values
// shown with the workbook's own number formats.
const ACTIVITY_FIRST_ROW = 12;
const ACTIVITY_LAST_ROW = 22;
const ACTIVITY_COLS = [
  { col: 3, label: 'W.D' },
  { col: 4, label: 'Net Start' },
  { col: 5, label: 'Net Finish' },
  { col: 6, label: 'Gross %' },
  { col: 7, label: 'Net %' },
];

function show(wb, a) {
  const v = wb.getValue(S.CSWS, a);
  if (v === null || v === undefined || v === '') return '';
  return formatValue(v, wb.styles[wb.getCell(S.CSWS, a)?.s]?.numFmt || 'General');
}

function readActivities(wb, periods) {
  const out = [];
  for (let r = ACTIVITY_FIRST_ROW; r <= ACTIVITY_LAST_ROW; r++) {
    const label = String(wb.getValue(S.CSWS, addr(r, 1)) ?? '').trim();
    if (!label) continue;
    // Total impact days: the sum of the period values (I..CZ), as in the Gross % formula's numerator.
    const total = periods.reduce((sum, { c }) => {
      const v = wb.getValue(S.CSWS, addr(r, c));
      return typeof v === 'number' ? sum + v : sum;
    }, 0);
    out.push({
      row: r,
      label,
      summary: ACTIVITY_COLS.map(({ col }) => show(wb, addr(r, col))),
      total: formatValue(total, '0.0'),
      periods: periods.map(({ c }) => show(wb, addr(r, c))),
    });
  }
  return out;
}

export default function CriticalSiteSummary({ store }) {
  const { wb, state } = store;
  const [error, setError] = useState(null);
  const extra = state.extraRatingRows || 0;
  const periods = readPeriods(wb);
  const rows = [
    { row: RATING_ROW, label: 'Rain Exposure Rating' },
    ...Array.from({ length: extra }, (_, k) => ({ row: extraRatingRow(k), label: `Rain Exposure Rating ${k + 2}` })),
  ].map((r) => ({ ...r, cells: readRow(wb, state.overrides, r.row, periods) }));
  const edited = rows.flatMap((r) => r.cells.filter((p) => p.edited).map((p) => p.a));
  const activities = readActivities(wb, periods);

  const commit = (a, raw) => {
    const msg = inputError(S.CSWS, a, raw);
    if (msg) { setError(msg); return; }
    setError(null);
    store.setOverride(S.CSWS, a, raw);
  };

  const removeRow = () => {
    const last = rows[rows.length - 1];
    if (last.cells.some((p) => p.rating !== '')
      && !window.confirm(`Remove “${last.label}” and its ratings?`)) return;
    setError(null);
    store.removeRatingRow();
  };

  return (
    <div className="page">
      <header className="page-head row">
        <div>
          <h2>Critical Site Works Summary</h2>
          <p className="muted">Enter the Rain Exposure Rating for each period — delay allowances and working days recalculate automatically from the weather impact results.</p>
        </div>
        <div className="btn-row">
          <button type="button" className="btn ghost" disabled={!edited.length}
            onClick={() => {
              if (window.confirm(`Discard ${edited.length} change(s) on this page?`)) { setError(null); store.resetSheet(S.CSWS, edited); }
            }}>
            Reset changes{edited.length ? ` (${edited.length})` : ''}
          </button>
        </div>
      </header>
      {error && <div className="alert error" role="alert">{error}</div>}
      <section className="card impact-card impact-fit">
        <header className="impact-head">
          <h3>Rain exposure ratings</h3>
          <p>Choose a rating from {RATING_MIN} to {RATING_MAX} for each period, or — for none. Use + to add another rating row; the wet-weather days lost from every row are added to the delay allowance. Scroll sideways to see all periods.</p>
        </header>
        <div className="table-scroll">
          <table className="impact-table cal-table csws-table csws-compact">
            <thead>
              <tr>
                <th scope="row">Period</th>
                {periods.map((p) => <th key={p.c} scope="col">{p.label}</th>)}
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={r.row}>
                  <th scope="row">
                    <span className="impact-swatch" style={{ background: SERIES[i % SERIES.length] }} aria-hidden="true" />
                    {r.label}
                  </th>
                  {r.cells.map((p, j) => (
                    <td key={p.a} className={p.edited ? 'is-edited' : undefined}>
                      <select className="cal-input csws-select" value={p.rating}
                        aria-label={`${r.label} for ${periods[j].label}`}
                        onChange={(e) => commit(p.a, e.target.value)}>
                        <option value="">—</option>
                        {RATINGS.map((v) => <option key={v} value={v}>{v}</option>)}
                        {p.rating !== '' && !RATINGS.includes(p.rating) && <option value={p.rating}>{p.rating}</option>}
                      </select>
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="btn-row csws-row-actions">
          <button type="button" className="btn ghost" onClick={store.addRatingRow}
            disabled={extra >= MAX_EXTRA_RATING_ROWS} aria-label="Add a rating row" title="Add a rating row">
            + Add row
          </button>
          {extra > 0 && (
            <button type="button" className="btn ghost" onClick={removeRow}
              aria-label="Remove the last rating row" title="Remove the last rating row">
              − Remove row
            </button>
          )}
        </div>
        <p className="impact-foot">Edited periods are highlighted.</p>
      </section>
      <section className="card impact-card impact-fit">
        <header className="impact-head">
          <h3>Critical activities</h3>
          <p>Working days and weather impact for each activity, the total impact days, and the working days lost in each period the activity is on site. These are calculated and can't be edited here.</p>
        </header>
        <div className="table-scroll">
          <table className="impact-table cal-table csws-table csws-compact csws-activities">
            <thead>
              <tr>
                <th scope="col">Activity</th>
                {ACTIVITY_COLS.map((h) => <th key={h.col} scope="col" className="csws-sum">{h.label}</th>)}
                <th scope="col" className="csws-sum csws-total" title="Total impact days for the activity">Total</th>
                {periods.map((p) => <th key={p.c} scope="col">{p.label}</th>)}
              </tr>
            </thead>
            <tbody>
              {activities.map((act) => (
                <tr key={act.row}>
                  <th scope="row">{act.label}</th>
                  {act.summary.map((v, j) => <td key={ACTIVITY_COLS[j].col} className="csws-sum">{v}</td>)}
                  <td className="csws-sum csws-total">{act.total}</td>
                  {act.periods.map((v, j) => <td key={periods[j].c} className={v ? 'csws-active' : undefined}>{v}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}
