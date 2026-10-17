import { useState } from 'react';
import SheetGrid from '../components/SheetGrid.jsx';

// Full, editable mirror of a workbook sheet.
export default function SheetPage({ store, version, sheet, title, intro, hiddenToggle, shapes }) {
  const { wb } = store;
  const [showHidden, setShowHidden] = useState(false);
  const edited = Object.keys(store.state.overrides).filter((k) => k.startsWith(`${sheet}|`)).length;
  return (
    <div className="page">
      <header className="page-head row">
        <div>
          <h2>{title}</h2>
          <p className="muted">{intro}</p>
        </div>
        <div className="btn-row">
          {hiddenToggle && (
            <button type="button" className="btn ghost" onClick={() => setShowHidden((v) => !v)} aria-pressed={showHidden}>
              {showHidden ? 'Hide activity rows' : 'Show activity rows'}
            </button>
          )}
          <button type="button" className="btn ghost" disabled={!edited}
            onClick={() => { if (window.confirm(`Discard ${edited} change(s) on this page?`)) store.resetSheet(sheet); }}>
            Reset changes{edited ? ` (${edited})` : ''}
          </button>
        </div>
      </header>
      <p className="muted small hint">Click an input cell and type to change it (Enter to confirm, Esc to cancel). Calculated values are locked.</p>
      <div className="card sheet-card">
        <SheetGrid wb={wb} sheet={sheet} version={version} editable showHidden={showHidden} shapes={shapes}
          onEdit={(a, raw) => store.setOverride(sheet, a, raw)} />
      </div>
    </div>
  );
}
