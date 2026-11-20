import { useState } from 'react';
import { addr } from '../engine/address.js';
import { S } from '../model/build.js';

// Weather Analysis Impact Table (Weather Calculation sheet): column D annual days, E share of the working year,
// F key (rating number or factor name), H..S months Jan..Dec.
const COL = { annual: 4, share: 5, key: 6, month0: 8 };
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const RAIN = 'var(--series-1)';
const OTHER = 'var(--series-2)';
const FACTOR_NAMES = {
  Heat: 'Heat', Cold: 'Cold', Wind: 'Wind', Client: 'Client / authority', Subbie: 'Subcontractor',
  IR: 'Industrial relations', Plant: 'Plant', Other: 'Other',
};

const num = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : null);
const fmt = (v) => (v === null || v === undefined ? '–' : (Math.round(v * 10) / 10).toLocaleString(undefined, { maximumFractionDigits: 1 }));
const shown = (v) => fmt(v) !== '0' && fmt(v) !== '–'; // values that round to 0 read as zero
const catName = (label, i) => String(label || '').replace(/^\s*RAIN EXPOSURE\s*\d+\s*:\s*/i, '').trim() || `Rating ${i + 1}`;

function readRow(wb, r) {
  return {
    annual: num(wb.getValue(S.WC, addr(r, COL.annual))),
    share: num(wb.getValue(S.WC, addr(r, COL.share))),
    months: MONTHS.map((_, m) => num(wb.getValue(S.WC, addr(r, COL.month0 + m)))),
  };
}

function impactRows(store) {
  const { wb, layout, state } = store;
  const rain = layout.exposureRows.map((r, i) => ({
    id: `r${i + 1}`, rating: i + 1, name: catName(state.categories[i]?.label, i), ...readRow(wb, r),
  }));
  const others = [];
  for (let r = layout.first + layout.exposureRows.length; r <= layout.last; r++) {
    const key = String(wb.getValue(S.WC, addr(r, COL.key)) ?? '').trim();
    if (key) others.push({ id: key, name: FACTOR_NAMES[key] || key, ...readRow(wb, r) });
  }
  return { rain, others };
}

function ImpactTable({ rain, others }) {
  const max = Math.max(0.0001, ...[...rain, ...others].flatMap((r) => r.months.map((v) => v ?? 0)));
  const tint = (v, color) => (v > 0 && shown(v)
    ? { background: `color-mix(in srgb, ${color} ${Math.round(6 + 30 * (v / max))}%, transparent)` }
    : undefined);
  const row = (r, color) => (
    <tr key={r.id}>
      <th scope="row">
        <span className="impact-swatch" style={{ background: color }} aria-hidden="true" />
        {r.rating ? <><span className="impact-rating">{r.rating}</span>{r.name}</> : r.name}
      </th>
      <td className={`impact-total ${shown(r.annual) ? '' : 'is-zero'}`}>{fmt(r.annual)}</td>
      <td className="impact-share">{r.share === null ? '–' : `${fmt(r.share * 100)}%`}</td>
      {r.months.map((v, m) => (
        <td key={m} style={tint(v, color)} className={shown(v) ? undefined : 'is-zero'}>{fmt(v)}</td>
      ))}
    </tr>
  );
  return (
    <section className="card impact-card impact-fit">
      <header className="impact-head">
        <h3>Weather analysis impact table</h3>
        <p>Working days lost per month and per year, and the share of the working year. Darker cells lose more days.</p>
      </header>
      <div className="table-scroll">
        <table className="impact-table impact-compact">
          <thead>
            <tr>
              <th scope="col">Factor</th>
              <th scope="col" className="impact-total">Annual</th>
              <th scope="col" className="impact-share">% of year</th>
              {MONTHS.map((m) => <th key={m} scope="col">{m}</th>)}
            </tr>
          </thead>
          <tbody>
            <tr className="impact-group"><th colSpan={15} scope="colgroup">Rain exposure ratings</th></tr>
            {rain.map((r) => row(r, RAIN))}
            <tr className="impact-group"><th colSpan={15} scope="colgroup">Other delay factors</th></tr>
            {others.map((r) => row(r, OTHER))}
          </tbody>
        </table>
      </div>
      <p className="impact-foot">Rain rows are alternatives — each period uses one rating. Other delay factors apply on top of the rain impact.</p>
    </section>
  );
}

export default function WeatherImpact({ store }) {
  const { wb } = store;
  const [busy, setBusy] = useState(false);
  const { rain, others } = impactRows(store);
  const download = async () => {
    setBusy(true);
    try {
      const { exportSheet } = await import('../export/exportSheet.js');
      await exportSheet(wb, S.WC, 'Weather Calculation.xlsx');
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="page">
      <header className="page-head row">
        <div>
          <h2>Weather impact</h2>
          <p className="muted">Working days lost per month for each rain exposure category and other delay factors. Download the weather calculation for the full breakdown.</p>
        </div>
        <button type="button" className="btn primary" onClick={download} disabled={busy}>
          {busy ? 'Preparing…' : 'Download weather calculation'}
        </button>
      </header>
      <ImpactTable rain={rain} others={others} />
    </div>
  );
}
