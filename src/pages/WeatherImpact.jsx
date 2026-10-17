import { useState } from 'react';
import SheetGrid from '../components/SheetGrid.jsx';
import { S } from '../model/build.js';

export default function WeatherImpact({ store, version }) {
  const { wb, layout } = store;
  const [busy, setBusy] = useState(false);
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
      <div className="card sheet-card">
        <SheetGrid wb={wb} sheet={S.WC} version={version} rows={[layout.title, layout.last]} cols={[2, 19]} />
      </div>
    </div>
  );
}
