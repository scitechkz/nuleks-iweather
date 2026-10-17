import { useState } from 'react';
import {
  ResponsiveContainer, BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
} from 'recharts';
import { addr } from '../engine/address.js';
import { S } from '../model/build.js';

// The four charts on the Climate sheet, in sheet order, with their Excel series references.
const CHARTS = [
  { title: 'Mean maximum temperature (°C)', type: 'bar', series: [{ row: 4, color: '#FF0000' }] },
  { title: 'Lowest Temperature Celcius', type: 'bar', series: [{ row: 16, color: '#5B9BD5' }] },
  { title: 'Mean Monthly Rainfall (mm)', type: 'bar', series: [{ row: 29, color: '#5B9BD5' }] },
  {
    title: 'Mean Wind Speeds km/h', type: 'line', yMin: 6,
    series: [{ row: 55, color: '#5B9BD5', named: true }, { row: 62, color: '#ED7D31', named: true }],
  },
];

const num = (v) => (typeof v === 'number' ? v : null);
const fmt = (v) => (v === null || v === undefined ? '–' : Number(v.toFixed(2)).toString());

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
  const raw = (hi - lo || 1) / 6;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((x) => x >= raw);
  const min = Math.floor(lo / step) * step;
  const max = Math.ceil(hi / step) * step;
  const ticks = [];
  for (let t = min; t <= max + step / 2; t += step) ticks.push(Number(t.toFixed(6)));
  return { domain: [min, max], ticks };
}

function ChartCard({ wb, chart }) {
  const [table, setTable] = useState(false);
  const months = Array.from({ length: 12 }, (_, i) => wb.getValue(S.CLIMATE, addr(2, 2 + i)));
  const names = chart.series.map((s) => String(wb.getValue(S.CLIMATE, addr(s.row, 1)) ?? '').trim());
  const data = months.map((m, i) => {
    const d = { month: m };
    chart.series.forEach((s, k) => { d[`s${k}`] = num(wb.getValue(S.CLIMATE, addr(s.row, 2 + i))); });
    return d;
  });
  const station = String(wb.getValue(S.BOM, 'A8') ?? '').trim();
  const axisTick = { fill: 'var(--chart-ink)', fontSize: 12 };
  const common = (
    <>
      <CartesianGrid vertical={false} stroke="var(--chart-grid)" />
      <XAxis dataKey="month" tick={axisTick} tickLine={false} axisLine={{ stroke: 'var(--chart-axis)' }} />
      <YAxis tick={axisTick} tickLine={false} axisLine={false} width={44} {...yAxisProps(chart, data)} />
      <Tooltip
        cursor={chart.type === 'bar' ? { fill: 'var(--chart-cursor)' } : { stroke: 'var(--chart-axis)' }}
        formatter={(v, key) => [fmt(v), names[Number(String(key).slice(1))] || chart.title]}
        contentStyle={{ background: 'var(--surface)', border: '1px solid var(--line)', color: 'var(--ink)', fontSize: 12 }}
      />
    </>
  );
  return (
    <section className="card chart-card">
      <h3 className="chart-title">{chart.title}</h3>
      <div className="chart-box">
        <ResponsiveContainer width="100%" height={300}>
          {chart.type === 'bar' ? (
            <BarChart data={data} barCategoryGap="38%" margin={{ top: 8, right: 12, left: 0, bottom: 4 }}>
              {common}
              {chart.series.map((s, k) => <Bar key={k} dataKey={`s${k}`} fill={s.color} isAnimationActive={false} />)}
            </BarChart>
          ) : (
            <LineChart data={data} margin={{ top: 8, right: 12, left: 0, bottom: 4 }}>
              {common}
              <Legend verticalAlign="bottom" formatter={(v) => <span style={{ color: 'var(--ink-2)' }}>{names[Number(v.slice(1))]}</span>} />
              {chart.series.map((s, k) => (
                <Line key={k} type="linear" dataKey={`s${k}`} stroke={s.color} strokeWidth={2.25}
                  dot={{ r: 4, fill: s.color, stroke: s.color }} activeDot={{ r: 6 }} isAnimationActive={false} />
              ))}
            </LineChart>
          )}
        </ResponsiveContainer>
      </div>
      <div className="chart-foot">
        <span className="muted">{station}</span>
        <button type="button" className="link-btn" onClick={() => setTable((t) => !t)}>
          {table ? 'Hide data' : 'Show data'}
        </button>
      </div>
      {table && (
        <div className="table-scroll">
          <table className="data-table">
            <thead>
              <tr><th />{months.map((m) => <th key={m}>{m}</th>)}</tr>
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
        <p className="muted">{typeof station === 'string' && station ? station : ''} — long-term monthly averages from the uploaded BoM climate statistics.</p>
      </header>
      <div className="chart-grid">
        {CHARTS.map((c) => <ChartCard key={c.title} wb={wb} chart={c} />)}
      </div>
    </div>
  );
}
