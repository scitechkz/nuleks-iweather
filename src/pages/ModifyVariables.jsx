import { useEffect, useState } from 'react';
import { addr } from '../engine/address.js';
import { formatValue } from '../engine/format.js';
import { S, rainCols } from '../model/build.js';

const RAIN_LABELS = ['<1mm', '>1mm', '>10mm', '>25mm'];
const RAIN_ROWS = [8, 9, 10, 11];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const show = (v, f = '0.00') => (typeof v === 'number' ? formatValue(v, f) : v ?? '');

// Text/number input that commits on blur or Enter (keeps recalculation snappy).
function Field({ value, onCommit, type = 'number', step = '0.05', label, multiline, className = '' }) {
  const [text, setText] = useState(value ?? '');
  useEffect(() => setText(value ?? ''), [value]);
  const commit = () => {
    if (String(text) !== String(value ?? '')) onCommit(type === 'number' && text !== '' ? Number(text) : text);
  };
  if (multiline) {
    return (
      <textarea className={`field ${className}`} aria-label={label} value={text} rows={3}
        onChange={(e) => setText(e.target.value)} onBlur={commit} />
    );
  }
  return (
    <input className={`field ${className}`} aria-label={label} type={type} step={step} value={text}
      onChange={(e) => setText(e.target.value)} onBlur={commit}
      onKeyDown={(e) => { if (e.key === 'Enter') e.currentTarget.blur(); }} />
  );
}

function CategoryCard({ wb, cat, index, onChange, removable, onRemove }) {
  const { total } = rainCols(index);
  return (
    <article className="card cat-card">
      <div className="cat-head">
        <span className="cat-badge">{index + 1}</span>
        {removable && (
          <button type="button" className="link-btn danger" onClick={onRemove}>Remove</button>
        )}
      </div>
      <label className="lbl">Category heading
        <Field type="text" multiline value={cat.title} label={`Category ${index + 1} heading`}
          onCommit={(v) => onChange({ title: v })} />
      </label>
      <label className="lbl">Name in the weather impact table
        <Field type="text" value={cat.label} label={`Category ${index + 1} name`}
          onCommit={(v) => onChange({ label: v })} />
      </label>
      <table className="var-table">
        <thead>
          <tr><th>Daily rainfall</th><th>Direct impact</th><th>Consequential delay</th><th>Total impact</th></tr>
        </thead>
        <tbody>
          {RAIN_LABELS.map((l, k) => (
            <tr key={l}>
              <th>{l}</th>
              <td><Field value={cat.direct[k]} label={`${l} direct impact`}
                onCommit={(v) => onChange({ direct: cat.direct.map((x, j) => (j === k ? Number(v) || 0 : x)) })} /></td>
              <td><Field value={cat.cons[k]} label={`${l} consequential delay`}
                onCommit={(v) => onChange({ cons: cat.cons.map((x, j) => (j === k ? Number(v) || 0 : x)) })} /></td>
              <td className="calc">{show(wb.getValue(S.RAIN, addr(RAIN_ROWS[k], total)))}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </article>
  );
}

export default function ModifyVariables({ store }) {
  const { wb, state, layout } = store;
  const cats = state.categories;
  const ov = (sheet, a) => {
    const c = wb.getCell(sheet, a);
    return c?.v ?? '';
  };
  const set = (sheet, a) => (v) => store.setOverride(sheet, a, v === '' ? '' : String(v));

  const manualRows = [7, 8, 9, 10, 11, 12];
  const thresholds = [
    { name: 'Heat (Mean max temp °C)', low: 'C153', high: 'E153', text: ['Low below', 'High above'] },
    { name: 'Cold (Lowest temp °C)', low: 'C173', high: 'E173', text: ['Low above', 'High below'] },
    { name: 'Rain (Mean rainfall mm)', low: 'C194', high: 'E194', text: ['Low below', 'High above'] },
    { name: 'Wind (Mean 3pm km/h)', low: 'C211', high: 'E211', text: ['Low below', 'High above'] },
  ];

  return (
    <div className="page">
      <header className="page-head">
        <h2>Modify variables</h2>
        <p className="muted">Changes here update the weather impact results and the critical site summary immediately.</p>
      </header>

      <section className="section">
        <div className="section-head">
          <h3>Rain impact factors &amp; categories</h3>
          <div className="btn-row">
            <button type="button" className="btn" onClick={store.addCategory}>+ Add category</button>
            <button type="button" className="btn ghost" onClick={store.resetCategories}>Reset to template</button>
          </div>
        </div>
        <p className="muted small">Total impact = Direct impact + Consequential delay (days lost per rain day). New categories get their own row in the weather impact table and can be chosen as a Rain Exposure Rating on the critical site summary.</p>
        <div className="cat-grid">
          {cats.map((c, i) => (
            <CategoryCard key={i} wb={wb} cat={c} index={i}
              onChange={(patch) => store.updateCategory(i, patch)}
              removable={i >= 5 && i === cats.length - 1} onRemove={store.removeCategory} />
          ))}
        </div>

        <h4 className="sub">Rain summary table</h4>
        <div className="table-scroll">
          <table className="data-table">
            <thead>
              <tr><th>Exposure level</th><th>Category</th>{['<1mm', '>1mm', '>10mm', '>25mm'].map((h) => <th key={h}>{h}</th>)}</tr>
            </thead>
            <tbody>
              {cats.map((c, i) => (
                <tr key={i}>
                  <th>{wb.getValue(S.RAIN, addr(20 + i, 2))}</th>
                  <td className="left">{wb.getValue(S.WC, addr(layout.blockTotal(i), 1))}</td>
                  {[3, 4, 5, 6].map((col) => <td key={col}>{show(wb.getValue(S.RAIN, addr(20 + i, col)))}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="section">
        <h3>Hot &amp; cold weather direct impact</h3>
        <div className="hc-grid">
          {[{ a: 'E5', row: 5, name: 'Hot weather', hint: 'Days lost per average day over 37 °C' },
            { a: 'E6', row: 6, name: 'Cold weather', hint: 'Days lost per average day below 2 °C' }].map((x) => (
            <div key={x.a} className="card hc-card">
              <div className="hc-name">{x.name}</div>
              <div className="muted small">{x.hint}</div>
              <Field value={ov(S.WC, x.a)} label={`${x.name} direct impact`} onCommit={set(S.WC, x.a)} className="big" />
              <div className="muted small">Annual total: <strong>{show(wb.getValue(S.WC, `F${x.row}`))}</strong> working days</div>
            </div>
          ))}
        </div>
      </section>

      <section className="section">
        <h3>Other monthly delay allowances <span className="muted">(working days)</span></h3>
        <div className="table-scroll">
          <table className="var-table wide">
            <thead>
              <tr><th>Allowance</th>{MONTHS.map((m) => <th key={m}>{m}</th>)}<th>Total</th></tr>
            </thead>
            <tbody>
              {manualRows.map((r) => (
                <tr key={r}>
                  <th className="left">{String(wb.getValue(S.WC, `A${r}`)).replace(/^Manual Input\s*:\s*/, '').trim()}</th>
                  {MONTHS.map((m, k) => {
                    const a = addr(r, 8 + k);
                    return <td key={m}><Field value={ov(S.WC, a)} label={`${m} row ${r}`} onCommit={set(S.WC, a)} step="0.5" /></td>;
                  })}
                  <td className="calc">{show(wb.getValue(S.WC, `F${r}`))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="section">
        <h3>Risk rating thresholds <span className="muted">(Low / Med / High)</span></h3>
        <div className="hc-grid">
          {thresholds.map((t) => (
            <div key={t.name} className="card hc-card">
              <div className="hc-name">{t.name}</div>
              <label className="lbl inline">{t.text[0]}
                <Field value={ov(S.CLIMATE, t.low)} label={`${t.name} ${t.text[0]}`} onCommit={set(S.CLIMATE, t.low)} step="1" />
              </label>
              <label className="lbl inline">{t.text[1]}
                <Field value={ov(S.CLIMATE, t.high)} label={`${t.name} ${t.text[1]}`} onCommit={set(S.CLIMATE, t.high)} step="1" />
              </label>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}
