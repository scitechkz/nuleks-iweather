import { useEffect, useState } from 'react';
import { addr } from '../engine/address.js';
import { S, inputError } from '../model/build.js';

// Project Site Calendar, reduced to the editable Calendar Days row (row 4, Jan..Dec in I..T; U4 is the annual sum).
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const ROW = 4;
const FIRST_COL = 9; // I
const DAY_ADDRS = MONTHS.map((_, m) => addr(ROW, FIRST_COL + m));

// Number cell that commits on blur or Enter; Esc restores the current value.
function DaysInput({ value, label, onCommit }) {
  const [text, setText] = useState(value);
  useEffect(() => setText(value), [value]);
  const commit = () => { if (String(text).trim() !== String(value)) onCommit(String(text).trim()); };
  return (
    <input className="cal-input" type="number" inputMode="numeric" min={1} max={31} step={1}
      aria-label={label} value={text}
      onChange={(e) => setText(e.target.value)} onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === 'Enter') e.currentTarget.blur();
        if (e.key === 'Escape') { setText(value); e.currentTarget.blur(); }
      }} />
  );
}

export default function ProjectCalendar({ store }) {
  const { wb, state } = store;
  const [error, setError] = useState(null);
  const days = DAY_ADDRS.map((a) => wb.getValue(S.CAL, a));
  const total = wb.getValue(S.CAL, addr(ROW, FIRST_COL + 12));
  const edited = DAY_ADDRS.filter((a) => `${S.CAL}|${a}` in state.overrides);

  const commit = (a, raw) => {
    const msg = inputError(S.CAL, a, raw);
    if (msg) { setError(msg); return; }
    setError(null);
    store.setOverride(S.CAL, a, raw);
  };

  return (
    <div className="page">
      <header className="page-head row">
        <div>
          <h2>Project Site Calendar</h2>
          <p className="muted">Enter the number of calendar days in each month — working days and the ratios used throughout the analysis recalculate automatically.</p>
        </div>
        <div className="btn-row">
          <button type="button" className="btn ghost" disabled={!edited.length}
            onClick={() => {
              if (window.confirm(`Discard ${edited.length} change(s) on this page?`)) { setError(null); store.resetSheet(S.CAL, edited); }
            }}>
            Reset changes{edited.length ? ` (${edited.length})` : ''}
          </button>
        </div>
      </header>
      {error && <div className="alert error" role="alert">{error}</div>}
      <section className="card impact-card impact-fit">
        <header className="impact-head">
          <h3>Calendar days</h3>
          <p>Days in each month of the project calendar. Whole numbers from 1 to 31 — use 29 for February in a leap year.</p>
        </header>
        <div className="table-scroll">
          <table className="impact-table cal-table cal-compact">
            <thead>
              <tr>
                <th scope="col">Month</th>
                {MONTHS.map((m) => <th key={m} scope="col">{m}</th>)}
                <th scope="col" className="impact-total">Annual</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <th scope="row">
                  <span className="impact-swatch" style={{ background: 'var(--series-1)' }} aria-hidden="true" />
                  Calendar days
                </th>
                {DAY_ADDRS.map((a, m) => (
                  <td key={a} className={`${S.CAL}|${a}` in state.overrides ? 'is-edited' : undefined}>
                    <DaysInput value={days[m] ?? ''} label={`Calendar days in ${MONTHS[m]}`} onCommit={(raw) => commit(a, raw)} />
                  </td>
                ))}
                <td className="impact-total">{typeof total === 'number' ? total : '–'}</td>
              </tr>
            </tbody>
          </table>
        </div>
        <p className="impact-foot">Click a month to change it (Enter to confirm, Esc to cancel). Edited months are highlighted.</p>
      </section>
    </div>
  );
}
