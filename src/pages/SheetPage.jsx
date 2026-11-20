import { useState } from 'react';
import SheetGrid from '../components/SheetGrid.jsx';
import { inputError } from '../model/build.js';

// Editable mirror of a workbook sheet. Optional view props restrict it to part of the sheet:
// visibleRows, cols, editable (bool | addr=>bool), rowLabels, editableAddr (addr=>bool: which edits "Reset changes" counts and clears).
export default function SheetPage({
  store, version, sheet, title, intro, hiddenToggle, shapes, visibleRows, cols, editable = true, rowLabels, editableAddr, plain, options,
}) {
  const { wb } = store;
  const [showHidden, setShowHidden] = useState(false);
  const [error, setError] = useState(null);
  const editedAddrs = Object.keys(store.state.overrides)
    .filter((k) => k.startsWith(`${sheet}|`))
    .map((k) => k.slice(sheet.length + 1))
    .filter((a) => !editableAddr || editableAddr(a));
  const edited = editedAddrs.length;
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
            onClick={() => { if (window.confirm(`Discard ${edited} change(s) on this page?`)) store.resetSheet(sheet, editableAddr ? editedAddrs : undefined); }}>
            Reset changes{edited ? ` (${edited})` : ''}
          </button>
        </div>
      </header>
      <p className="muted small hint">Click an input cell and type to change it (Enter to confirm, Esc to cancel). Calculated values are locked.</p>
      {error && <div className="alert error" role="alert">{error}</div>}
      <div className="card sheet-card">
        <SheetGrid wb={wb} sheet={sheet} version={version} editable={editable} showHidden={showHidden} shapes={shapes}
          visibleRows={visibleRows} cols={cols} rowLabels={rowLabels} plain={plain} options={options}
          onEdit={(a, raw) => {
            const msg = inputError(sheet, a, raw);
            if (msg) { setError(msg); return; }
            setError(null);
            store.setOverride(sheet, a, raw);
          }} />
      </div>
    </div>
  );
}
