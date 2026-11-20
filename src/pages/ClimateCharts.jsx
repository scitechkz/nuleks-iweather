import { useState } from 'react';
import {
  ResponsiveContainer, BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, LabelList,
} from 'recharts';
import { addr } from '../engine/address.js';
import { S } from '../model/build.js';

// The four charts on the Climate sheet, in sheet order, with their Excel series references.
// Colours are CSS tokens (see styles.css) so light and dark mode each get their own validated steps.
// `stat` picks the headline figure; `mark` picks which bar gets a direct value label.
const CHARTS = [
  {
    title: 'Mean maximum temperature', unit: '°C', type: 'bar', stat: 'max', mark: 'max',
    series: [{ row: 4, color: 'var(--series-hot)' }],
  },
  {
    title: 'Lowest temperature', unit: '°C', type: 'bar', stat: 'min', mark: 'min',
    series: [{ row: 16, color: 'var(--series-cold)' }],
  },
  {
    title: 'Mean monthly rainfall', unit: 'mm', type: 'bar', stat: 'sum', mark: 'max',
    series: [{ row: 29, color: 'var(--series-rain)' }],
  },
  {
    title: 'Mean wind speed', unit: 'km/h', type: 'line', yMin: 6, stat: 'max',
    series: [{ row: 55, color: 'var(--series-1)' }, { row: 62, color: 'var(--series-2)' }],
  },
];

const num = (v) => (typeof v === 'number' ? v : null);
const fmt = (v) => (v === null || v === undefined ? '–' : (Math.round(v * 10) / 10).toLocaleString(undefined, { maximumFractionDigits: 1 }));

// "Mean 9am wind speed (km/h) for years 1990 to 2020" -> "9am"; other labels lose the period and units.
function shortName(name) {
  const t = /\b(\d{1,2}\s?[ap]m)\b/i.exec(name);
  if (t) return t[1].replace(/\s/g, '').toLowerCase();
  return name.replace(/\s+for years.*$/i, '').replace(/\s*\([^)]*\)\s*$/, '').trim();
}

// Excel axis behaviour: bar charts include zero; the wind chart has a fixed minimum of 6 (major unit 5).
function yAxisProps(chart, data) {
  const vals = data.flatMap((d) => chart.series.map((_, k) => d[`s${k}`])).filter((v) => typeof v === 'number');
  if (chart.yMin !== undefined) {
    const max = Math.max(chart.yMin + 5, ...vals);
    const ticks = [];
    for (let t = chart.yMin; t < max + 5; t += 5) ticks.push(t);
    return { domain: [chart.yMin, ticks[ticks.length - 1]], ticks, allowDataOverflow: true };
  }
  const lo = Math.min(0, ...vals);
  const hi = Math.max(0, ...vals);
  const raw = (hi - lo || 1) / 5;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((x) => x >= raw);
  const min = Math.floor(lo / step) * step;
  const max = Math.ceil((hi * 1.08) / step) * step; // headroom for the peak label
  const ticks = [];
  for (let t = min; t <= max + step / 2; t += step) ticks.push(Number(t.toFixed(6)));
  return { domain: [min, max], ticks };
}

// Bar with a 4px rounded data end (top for positive values, bottom for negative), square at the baseline.
function RoundedBar({ x, y, width, height, fill, value }) {
  if (!width || height === undefined || value === null || value === undefined) return null;
  const top = Math.min(y, y + height);
  const h = Math.abs(height);
  const r = Math.min(4, h, width / 2);
  const neg = Number(Array.isArray(value) ? value[1] : value) < 0;
  const x2 = x + width;
  const b = top + h;
  const d = neg
    ? `M${x},${top} H${x2} V${b - r} Q${x2},${b} ${x2 - r},${b} H${x + r} Q${x},${b} ${x},${b - r} Z`
    : `M${x},${b} V${top + r} Q${x},${top} ${x + r},${top} H${x2 - r} Q${x2},${top} ${x2},${top + r} V${b} Z`;
  return <path d={d} style={{ fill }} />;
}

function ChartTooltip({ active, payload, label, names, colors, unit }) {
  if (!active || !payload?.length) return null;
  return (
    <div className="viz-tip">
      <div className="viz-tip-head">{label}</div>
      {payload.map((p) => {
        const k = Number(String(p.dataKey).slice(1));
        return (
          <div key={p.dataKey} className="viz-tip-row">
            <span className="viz-swatch" style={{ background: colors[k] }} aria-hidden="true" />
            <span className="viz-tip-name">{names[k]}</span>
            <span className="viz-tip-val">{fmt(p.value)} <span className="viz-unit">{unit}</span></span>
          </div>
        );
      })}
    </div>
  );
}

function headline(chart, data, names) {
  const pts = data.flatMap((d) => chart.series.map((_, k) => ({ v: d[`s${k}`], m: d.month, k })))
    .filter((p) => typeof p.v === 'number');
  if (!pts.length) return null;
  if (chart.stat === 'sum') {
    const wettest = pts.reduce((a, b) => (b.v > a.v ? b : a));
    return { label: 'Annual total', value: pts.reduce((t, p) => t + p.v, 0), note: `Wettest: ${wettest.m}` };
  }
  const pick = pts.reduce((a, b) => ((chart.stat === 'min' ? b.v < a.v : b.v > a.v) ? b : a));
  const label = chart.stat === 'min' ? 'Coldest month' : chart.type === 'line' ? 'Windiest' : 'Hottest month';
  const who = chart.series.length > 1 ? ` · ${names[pick.k]}` : '';
  return { label, value: pick.v, note: `${pick.m}${who}` };
}

function ChartCard({ wb, chart }) {
  const [table, setTable] = useState(false);
  const months = Array.from({ length: 12 }, (_, i) => wb.getValue(S.CLIMATE, addr(2, 2 + i)));
  const fullNames = chart.series.map((s) => String(wb.getValue(S.CLIMATE, addr(s.row, 1)) ?? '').trim());
  const names = fullNames.map((n) => (chart.series.length > 1 ? shortName(n) : chart.title));
  const colors = chart.series.map((s) => s.color);
  const data = months.map((m, i) => {
    const d = { month: m };
    chart.series.forEach((s, k) => { d[`s${k}`] = num(wb.getValue(S.CLIMATE, addr(s.row, 2 + i))); });
    return d;
  });
  const period = (/for years\s+(.*)$/i.exec(fullNames[0] || '') || [])[1];
  const stat = headline(chart, data, names);

  // Single direct label on the extreme bar (selective labelling, not a number on every bar).
  const vals0 = data.map((d) => d.s0).filter((v) => typeof v === 'number');
  const markVal = chart.mark && vals0.length ? (chart.mark === 'min' ? Math.min(...vals0) : Math.max(...vals0)) : null;
  const markIdx = data.findIndex((d) => d.s0 === markVal);

  const axisTick = { fill: 'var(--chart-muted)', fontSize: 12 };
  const margin = { top: 20, right: chart.type === 'line' ? 44 : 8, left: 0, bottom: 0 };
  const common = (
    <>
      <CartesianGrid vertical={false} stroke="var(--chart-grid)" />
      <XAxis dataKey="month" tick={axisTick} tickLine={false} axisLine={{ stroke: 'var(--chart-axis)' }} tickMargin={8} />
      <YAxis tick={axisTick} tickLine={false} axisLine={false} width={40} {...yAxisProps(chart, data)} />
      <Tooltip
        cursor={chart.type === 'bar' ? { fill: 'var(--chart-cursor)' } : { stroke: 'var(--chart-axis)', strokeWidth: 1 }}
        content={<ChartTooltip names={names} colors={colors} unit={chart.unit} />}
        isAnimationActive={false}
      />
    </>
  );

  return (
    <section className="card chart-card" aria-label={`${chart.title} chart`}>
      <header className="chart-head">
        <div>
          <h3 className="chart-title">{chart.title}</h3>
          <p className="chart-sub">Monthly average · {chart.unit}{period ? ` · ${period}` : ''}</p>
        </div>
        {stat && (
          <div className="chart-stat">
            <div className="chart-stat-label">{stat.label}</div>
            <div className="chart-stat-value">{fmt(stat.value)}<span className="viz-unit"> {chart.unit}</span></div>
            <div className="chart-stat-note">{stat.note}</div>
          </div>
        )}
      </header>
      {chart.series.length > 1 && (
        <ul className="viz-legend" aria-label="Legend">
          {names.map((n, k) => (
            <li key={n}><span className="viz-swatch line" style={{ background: colors[k] }} aria-hidden="true" />{n}</li>
          ))}
        </ul>
      )}
      <div className="chart-box">
        <ResponsiveContainer width="100%" height={260}>
          {chart.type === 'bar' ? (
            <BarChart data={data} barCategoryGap="28%" margin={margin}>
              {common}
              {chart.series.map((s, k) => (
                <Bar key={k} dataKey={`s${k}`} fill={s.color} shape={<RoundedBar />} isAnimationActive={false}>
                  {k === 0 && markIdx >= 0 && (
                    <LabelList dataKey="s0" content={({ x, y, width, height, index, value }) => {
                      if (index !== markIdx) return null;
                      const below = value < 0;
                      const ty = below ? Math.max(y, y + height) + 15 : Math.min(y, y + height) - 6;
                      return (
                        <text x={x + width / 2} y={ty} textAnchor="middle" className="viz-label">{fmt(value)}</text>
                      );
                    }} />
                  )}
                </Bar>
              ))}
            </BarChart>
          ) : (
            <LineChart data={data} margin={margin}>
              {common}
              {chart.series.map((s, k) => (
                <Line key={k} type="monotone" dataKey={`s${k}`} stroke={s.color} strokeWidth={2}
                  dot={{ r: 4, fill: s.color, stroke: 'var(--surface)', strokeWidth: 2 }}
                  activeDot={{ r: 6, fill: s.color, stroke: 'var(--surface)', strokeWidth: 2 }}
                  isAnimationActive={false}>
                  <LabelList dataKey={`s${k}`} content={({ x, y, index }) => (index === data.length - 1
                    ? <text x={x + 10} y={y + 4} className="viz-label">{names[k]}</text>
                    : null)} />
                </Line>
              ))}
            </LineChart>
          )}
        </ResponsiveContainer>
      </div>
      <div className="chart-foot" style={{ justifyContent: 'flex-end' }}>
        <button type="button" className="link-btn" onClick={() => setTable((t) => !t)} aria-expanded={table}>
          {table ? 'Hide data table' : 'Show data table'}
        </button>
      </div>
      {table && (
        <div className="table-scroll">
          <table className="data-table">
            <thead>
              <tr><th>{chart.unit}</th>{months.map((m) => <th key={m}>{m}</th>)}</tr>
            </thead>
            <tbody>
              {chart.series.map((s, k) => (
                <tr key={k}>
                  <th>{names[k]}</th>
                  {data.map((d) => <td key={d.month}>{fmt(d[`s${k}`])}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

export default function ClimateCharts({ wb }) {
  const station = wb.getValue(S.CLIMATE, 'A1');
  return (
    <div className="page">
      <header className="page-head">
        <h2>Climate charts</h2>
        <p className="muted">{typeof station === 'string' && station ? `${station} — ` : ''}long-term monthly averages from the uploaded BoM climate statistics.</p>
      </header>
      <div className="chart-grid">
        {CHARTS.map((c) => <ChartCard key={c.title} wb={wb} chart={c} />)}
      </div>
    </div>
  );
}
