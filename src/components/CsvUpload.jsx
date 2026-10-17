import { useRef, useState } from 'react';
import Papa from 'papaparse';

// Upload control for the BoM monthly climate statistics csv.
export function useCsvPicker(store) {
  const ref = useRef(null);
  const [error, setError] = useState(null);
  const onFile = (file) => {
    setError(null);
    Papa.parse(file, {
      skipEmptyLines: false,
      complete: (res) => {
        try {
          store.loadCsv(res.data, file.name);
        } catch (e) {
          setError(e.message);
        }
      },
      error: (e) => setError(e.message),
    });
  };
  const input = (
    <input ref={ref} type="file" accept=".csv,text/csv" hidden
      onChange={(e) => { if (e.target.files?.[0]) onFile(e.target.files[0]); e.target.value = ''; }} />
  );
  return { input, open: () => ref.current?.click(), error };
}

export default function CsvUpload({ store, onView, viewing }) {
  const picker = useCsvPicker(store);
  const meta = store.state.bomMeta;
  return (
    <div className="csv">
      {picker.input}
      <button type="button" className="btn primary" onClick={picker.open}>
        {meta ? 'Upload new BoM csv' : 'Upload BoM csv'}
      </button>
      {meta && (
        <button type="button" className="btn ghost" onClick={onView} aria-pressed={viewing}>
          {viewing ? 'Hide BOM csv data' : 'View BOM csv data'}
        </button>
      )}
      {meta && <span className="csv-status">Using <strong>{meta.fileName}</strong></span>}
      {picker.error && <div className="alert error" role="alert">{picker.error}</div>}
      {meta?.warnings?.length > 0 && (
        <details className="alert warn">
          <summary>{meta.warnings.length} statistic row(s) don’t match the expected BoM layout — check these values.</summary>
          <ul>
            {meta.warnings.map((w) => (
              <li key={w.row}>Row {w.row}: expected “{w.expected}”, found “{w.got}”</li>
            ))}
          </ul>
        </details>
      )}
    </div>
  );
}
