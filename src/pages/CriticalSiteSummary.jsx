import { useState } from 'react';
import { addr } from '../engine/address.js';
import { formatValue } from '../engine/format.js';
import { S, inputError, RATING_MIN, RATING_MAX } from '../model/build.js';

// Critical Site Works Summary: the Period row (row 9) and the Rain Exposure Rating row (row 11)
// across every period column I..CZ, in the same table style as the other pages.
const FIRST_COL = 9; // I
const LAST_COL = 104; // CZ
const RATINGS = Array.from({ length: RATING_MAX - RATING_MIN + 1 }, (_, i) => String(RATING_MIN + i));

function readPeriods(wb, overrides) {
  const periods = [];
  for (let c = FIRST_COL; c <= LAST_COL; c++) {
    const serial = wb.getValue(S.CSWS, addr(9, c));
    if (typeof serial !== 'number') continue;
    const a = addr(11, c);
    const rating = wb.getValue(S.CSWS, a);
    periods.push({
      a,
      label: formatValue(serial, 'mmm-yy'),
      rating: rating === null || rating === undefined ? '' : String(rating),
      edited: `${S.CSWS}|${a}` in overrides,
    });
  }
  return periods;
}

export default function CriticalSiteSummary({ store }) {
  const { wb, state } = store;
  const [error, setError] = useState(null);
  const periods = readPeriods(wb, state.overrides);
  const edited = periods.filter((p) => p.edited).map((p) => p.a);

  const commit = (a, raw) => {
    const msg = inputError(S.CSWS, a, raw);
    if (msg) { setError(msg); return; }
    setError(null);
    store.setOverride(S.CSWS, a, raw);
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
          <p>Choose a rating from {RATING_MIN} to {RATING_MAX} for each period, or — for none. Scroll sideways to see all periods.</p>
        </header>
        <div className="table-scroll">
          <table className="impact-table cal-table csws-table csws-compact">
            <thead>
              <tr>
                <th scope="row">Period</th>
                {periods.map((p) => <th key={p.a} scope="col">{p.label}</th>)}
              </tr>
            </thead>
            <tbody>
              <tr>
                <th scope="row">
                  <span className="impact-swatch" style={{ background: 'var(--series-1)' }} aria-hidden="true" />
                  Rain Exposure Rating
                </th>
                {periods.map((p) => (
                  <td key={p.a} className={p.edited ? 'is-edited' : undefined}>
                    <select className="cal-input csws-select" value={p.rating}
                      aria-label={`Rain Exposure Rating for ${p.label}`}
                      onChange={(e) => commit(p.a, e.target.value)}>
                      <option value="">—</option>
                      {RATINGS.map((r) => <option key={r} value={r}>{r}</option>)}
                      {p.rating !== '' && !RATINGS.includes(p.rating) && <option value={p.rating}>{p.rating}</option>}
                    </select>
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
        <p className="impact-foot">Edited periods are highlighted.</p>
      </section>
    </div>
  );
}
