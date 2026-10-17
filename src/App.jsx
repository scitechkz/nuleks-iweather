import { useState } from 'react';
import { useStore } from './state/useStore.js';
import { S } from './model/build.js';
import CsvUpload, { useCsvPicker } from './components/CsvUpload.jsx';
import ClimateCharts from './pages/ClimateCharts.jsx';
import ModifyVariables from './pages/ModifyVariables.jsx';
import WeatherImpact from './pages/WeatherImpact.jsx';
import SheetPage from './pages/SheetPage.jsx';
import SheetGrid from './components/SheetGrid.jsx';

const NAV = [
  { id: 'climate', label: 'See climate charts' },
  { id: 'variables', label: 'Modify variables' },
  { id: 'impact', label: 'View weather impact' },
  { id: 'critical', label: 'Critical Site Summary' },
  { id: 'calendar', label: 'Project Calendar' },
];

// Shown on data-driven pages until a BoM csv has been uploaded.
function UploadFirst({ store, what }) {
  const picker = useCsvPicker(store);
  return (
    <div className="page">
      <section className="card empty">
        {picker.input}
        <div className="empty-icon" aria-hidden="true">☁</div>
        <h2>Upload your BoM climate statistics</h2>
        <p className="muted">{what} appear once a BoM monthly climate statistics csv has been uploaded. Download it from the Bureau of Meteorology
          (Climate statistics for Australian locations → choose the nearest station → “All years of data” csv) and upload it here.</p>
        <button type="button" className="btn primary" onClick={picker.open}>Upload BoM csv</button>
        {picker.error && <div className="alert error" role="alert">{picker.error}</div>}
      </section>
    </div>
  );
}

function initialTab() {
  try {
    const t = localStorage.getItem('iwa-tab');
    if (NAV.some((n) => n.id === t)) return t;
  } catch {
    /* ignore */
  }
  return 'climate';
}

export default function App() {
  const store = useStore();
  const [tab, setTabRaw] = useState(initialTab);
  const [showBom, setShowBom] = useState(false);
  const setTab = (t) => {
    setTabRaw(t);
    try { localStorage.setItem('iwa-tab', t); } catch { /* ignore */ }
  };
  // any state change produces a new object -> use it as the render version for memoised grids
  const version = store.state;
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark" aria-hidden="true">☂</span>
          <div>
            <h1>Inclement Weather Analysis</h1>
            <div className="muted small">{store.hasData ? String(store.wb.getValue(S.CSWS, 'H37') || '') : 'No climate data uploaded yet'}</div>
          </div>
        </div>
        <CsvUpload store={store} viewing={showBom} onView={() => setShowBom((v) => !v)} />
      </header>
      <nav className="tabs" aria-label="Sections">
        {NAV.map((n) => (
          <button key={n.id} type="button" className={`tab ${tab === n.id ? 'active' : ''}`}
            aria-current={tab === n.id ? 'page' : undefined} onClick={() => setTab(n.id)}>
            {n.label}
          </button>
        ))}
        <button type="button" className="tab reset" onClick={() => {
          if (window.confirm('Reset all variables, categories, calendar and summary edits to their defaults? The uploaded csv is kept.')) store.resetAll();
        }}>Reset all</button>
      </nav>
      <main>
        {showBom && store.hasData && (
          <section className="page bom-panel">
            <header className="page-head">
              <h2>BOM csv data</h2>
              <p className="muted">The uploaded csv as used by the analysis: station details first, then the “Statistic Element” header and one row per monthly statistic.</p>
            </header>
            <div className="card sheet-card">
              <SheetGrid wb={store.wb} sheet={S.BOM} version={version} rows={[5, store.wb.sheet(S.BOM).maxRow]} cols={[1, 17]} />
            </div>
          </section>
        )}
        {tab === 'climate' && (store.hasData ? <ClimateCharts wb={store.wb} /> : <UploadFirst store={store} what="Climate charts" />)}
        {tab === 'variables' && <ModifyVariables store={store} />}
        {tab === 'impact' && (store.hasData
          ? <WeatherImpact store={store} version={version} />
          : <UploadFirst store={store} what="Weather impact results" />)}
        {tab === 'critical' && !store.hasData && <UploadFirst store={store} what="The critical site summary and its delay allowances" />}
        {tab === 'critical' && store.hasData && (
          <SheetPage store={store} version={version} sheet={S.CSWS} title="Critical Site Works Summary" hiddenToggle shapes
            intro="Enter the project title, start dates and Rain Exposure Ratings — delay allowances and working days recalculate automatically from the weather impact results." />
        )}
        {tab === 'calendar' && (
          <SheetPage store={store} version={version} sheet={S.CAL} title="Project Site Calendar"
            intro="Enter non-work days and work hours for the project. Section B sets the working days per month and the work-hours ratio used throughout the analysis." />
        )}
      </main>
    </div>
  );
}
