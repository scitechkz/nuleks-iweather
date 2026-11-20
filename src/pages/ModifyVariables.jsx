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

// Card shell shared by every section: title, one-line description, optional actions.
function Panel({ title, sub, actions, children, fit, className = '' }) {
  return (
    <section className={`card impact-card mv-panel ${fit ? 'impact-fit' : ''} ${className}`}>
      <header className="mv-panel-head">
        <div className="impact-head">
          <h3>{title}</h3>
          {sub && <p>{sub}</p>}
        </div>
        {actions && <div className="btn-row">{actions}</div>}
      </header>
      {children}
    </section>
  );
}

function CategoryCard({ wb, cat, index, onChange, removable, onRemove }) {
  const { total } = rainCols(index);
  return (
    <article className="mv-cat">
      <div className="mv-cat-head">
        <span className="mv-cat-badge">{index + 1}</span>
        <span className="mv-cat-title">Rain exposure {index + 1}</span>
        {removable && (
          <button type="button" className="link-btn danger" onClick={onRemove}>Remove</button>
        )}
      </div>
      <label className="mv-lbl">Category heading
        <Field type="text" multiline value={cat.title} label={`Category ${index + 1} heading`}
          onCommit={(v) => onChange({ title: v })} className="cal-input mv-text" />
      </label>
      <label className="mv-lbl">Name in the weather impact table
        <Field type="text" value={cat.label} label={`Category ${index + 1} name`}
          onCommit={(v) => onChange({ label: v })} className="cal-input mv-text" />
      </label>
      <table className="impact-table impact-compact mv-table">
        <thead>
          <tr>
            <th scope="col">Daily rain</th>
            <th scope="col">Direct</th>
            <th scope="col">Consequential</th>
            <th scope="col" className="impact-total">Total</th>
          </tr>
        </thead>
        <tbody>
          {RAIN_LABELS.map((l, k) => (
            <tr key={l}>
              <th scope="row">{l}</th>
              <td><Field value={cat.direct[k]} label={`${l} direct impact`} className="cal-input mv-num"
                onCommit={(v) => onChange({ direct: cat.direct.map((x, j) => (j === k ? Number(v) || 0 : x)) })} /></td>
              <td><Field value={cat.cons[k]} label={`${l} consequential delay`} className="cal-input mv-num"
                onCommit={(v) => onChange({ cons: cat.cons.map((x, j) => (j === k ? Number(v) || 0 : x)) })} /></td>
              <td className="impact-total">{show(wb.getValue(S.RAIN, addr(RAIN_ROWS[k], total)))}</td>
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
    { name: 'Heat', unit: 'Mean max temp °C', low: 'C153', high: 'E153', text: ['Low below', 'High above'] },
    { name: 'Cold', unit: 'Lowest temp °C', low: 'C173', high: 'E173', text: ['Low above', 'High below'] },
    { name: 'Rain', unit: 'Mean rainfall mm', low: 'C194', high: 'E194', text: ['Low below', 'High above'] },
    { name: 'Wind', unit: 'Mean 3pm km/h', low: 'C211', high: 'E211', text: ['Low below', 'High above'] },
  ];

  return (
    <div className="page">
      <header className="page-head">
        <h2>Modify variables</h2>
        <p className="muted">Changes here update the weather impact results and the critical site summary immediately.</p>
      </header>

      <Panel title="Rain impact factors & categories"
        sub="Total impact = Direct impact + Consequential delay (days lost per rain day). New categories get their own row in the weather impact table and can be chosen as a Rain Exposure Rating on the critical site summary."
        actions={(
          <>
            <button type="button" className="btn" onClick={store.addCategory}>+ Add category</button>
            <button type="button" className="btn ghost" onClick={store.resetCategories}>Reset to template</button>
          </>
        )}>
        <div className="mv-cat-grid">
          {cats.map((c, i) => (
            <CategoryCard key={i} wb={wb} cat={c} index={i}
              onChange={(patch) => store.updateCategory(i, patch)}
              removable={i >= 5 && i === cats.length - 1} onRemove={store.removeCategory} />
          ))}
        </div>
      </Panel>

      <Panel fit title="Rain summary table" sub="Total working days lost per rain day for each exposure level.">
        <div className="table-scroll">
          <table className="impact-table impact-compact">
            <thead>
              <tr>
                <th scope="col">Exposure level</th>
                <th scope="col" className="mv-left">Category</th>
                {RAIN_LABELS.map((h) => <th key={h} scope="col">{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {cats.map((c, i) => (
                <tr key={i}>
                  <th scope="row">
                    <span className="impact-swatch" style={{ background: 'var(--series-1)' }} aria-hidden="true" />
                    {wb.getValue(S.RAIN, addr(20 + i, 2))}
                  </th>
                  <td className="mv-left">{wb.getValue(S.WC, addr(layout.blockTotal(i), 1))}</td>
                  {[3, 4, 5, 6].map((col) => <td key={col}>{show(wb.getValue(S.RAIN, addr(20 + i, col)))}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Panel>

      <Panel fit title="Hot & cold weather direct impact" sub="Working days lost per average day beyond each temperature limit.">
        <table className="impact-table impact-compact">
          <thead>
            <tr>
              <th scope="col">Factor</th>
              <th scope="col">Days lost per day</th>
              <th scope="col" className="impact-total">Annual total</th>
            </tr>
          </thead>
          <tbody>
            {[{ a: 'E5', row: 5, name: 'Hot weather', hint: 'over 37 °C' },
              { a: 'E6', row: 6, name: 'Cold weather', hint: 'below 2 °C' }].map((x) => (
              <tr key={x.a}>
                <th scope="row">
                  <span className="impact-swatch" style={{ background: 'var(--series-2)' }} aria-hidden="true" />
                  {x.name} <span className="mv-hint">{x.hint}</span>
                </th>
                <td><Field value={ov(S.WC, x.a)} label={`${x.name} direct impact`} onCommit={set(S.WC, x.a)} className="cal-input mv-num" /></td>
                <td className="impact-total">{show(wb.getValue(S.WC, `F${x.row}`))} <span className="mv-hint">days</span></td>
              </tr>
            ))}
          </tbody>
        </table>
      </Panel>

      <Panel fit title="Other monthly delay allowances" sub="Working days allowed each month for each delay factor.">
        <div className="table-scroll">
          <table className="impact-table cal-table cal-compact">
            <thead>
              <tr>
                <th scope="col">Allowance</th>
                {MONTHS.map((m) => <th key={m} scope="col">{m}</th>)}
                <th scope="col" className="impact-total">Total</th>
              </tr>
            </thead>
            <tbody>
              {manualRows.map((r) => {
                const name = String(wb.getValue(S.WC, `A${r}`)).replace(/^Manual Input\s*:\s*/, '').trim();
                return (
                  <tr key={r}>
                    <th scope="row">
                      <span className="impact-swatch" style={{ background: 'var(--series-2)' }} aria-hidden="true" />
                      {name}
                    </th>
                    {MONTHS.map((m, k) => {
                      const a = addr(r, 8 + k);
                      return (
                        <td key={m}>
                          <Field value={ov(S.WC, a)} label={`${name} ${m}`} onCommit={set(S.WC, a)} step="0.5" className="cal-input" />
                        </td>
                      );
                    })}
                    <td className="impact-total">{show(wb.getValue(S.WC, `F${r}`))}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Panel>

      <Panel fit title="Risk rating thresholds" sub="Limits that set each month's Low / Med / High risk rating.">
        <table className="impact-table impact-compact">
          <thead>
            <tr>
              <th scope="col">Factor</th>
              <th scope="col">Low</th>
              <th scope="col">High</th>
            </tr>
          </thead>
          <tbody>
            {thresholds.map((t) => (
              <tr key={t.name}>
                <th scope="row">
                  <span className="impact-swatch" style={{ background: 'var(--series-1)' }} aria-hidden="true" />
                  {t.name} <span className="mv-hint">{t.unit}</span>
                </th>
                <td>
                  <label className="mv-inline"><span className="mv-hint">{t.text[0]}</span>
                    <Field value={ov(S.CLIMATE, t.low)} label={`${t.name} ${t.text[0]}`} onCommit={set(S.CLIMATE, t.low)} step="1" className="cal-input mv-num" />
                  </label>
                </td>
                <td>
                  <label className="mv-inline"><span className="mv-hint">{t.text[1]}</span>
                    <Field value={ov(S.CLIMATE, t.high)} label={`${t.name} ${t.text[1]}`} onCommit={set(S.CLIMATE, t.high)} step="1" className="cal-input mv-num" />
                  </label>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Panel>
    </div>
  );
}
