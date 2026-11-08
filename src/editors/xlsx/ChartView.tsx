// Vẽ biểu đồ bằng SVG (cột, thanh ngang, đường, vùng, tròn) từ dữ liệu đã tính.
import type { ChartData, ChartKind } from './charts';
import { PALETTE } from './charts';

const fmt = (n: number) => {
  const a = Math.abs(n);
  if (a >= 1e9) return `${+(n / 1e9).toFixed(1)} tỷ`;
  if (a >= 1e6) return `${+(n / 1e6).toFixed(1)} tr`;
  if (a >= 1e4) return `${+(n / 1e3).toFixed(0)}k`;
  return String(+n.toFixed(2));
};

/** Khoảng chia trục "đẹp": 1, 2, 5 × 10^k. */
function niceScale(min: number, max: number, ticks = 5) {
  if (min === max) { max = min + 1; if (min > 0) min = 0; }
  const raw = (max - min) / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const out: number[] = [];
  for (let v = lo; v <= hi + step / 2; v += step) out.push(+v.toFixed(10));
  return { lo, hi, ticks: out };
}

export default function ChartView({ kind, title, data, w, h }: { kind: ChartKind; title?: string; data: ChartData; w: number; h: number }) {
  const titleH = title ? 26 : 8;
  const isPie = kind === 'pie';
  const legend = isPie ? data.cats.map((c, i) => ({ name: c, color: PALETTE[i % PALETTE.length] }))
    : data.series.map((s, i) => ({ name: s.name, color: PALETTE[i % PALETTE.length] }));
  const legendH = isPie ? 0 : 22;
  const body = { x: 0, y: titleH, w, h: h - titleH - legendH };

  let plot: JSX.Element[] = [];
  if (isPie) {
    const vals = (data.series[0]?.values ?? []).map((v) => Math.max(0, v));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const legendW = Math.min(140, w * 0.35);
    const r = Math.max(10, Math.min(body.w - legendW - 20, body.h - 16) / 2);
    const cx = (body.w - legendW) / 2, cy = body.y + body.h / 2;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2;
      const color = PALETTE[i % PALETTE.length];
      if (v / total >= 0.9999) plot.push(<circle key={i} cx={cx} cy={cy} r={r} fill={color} />);
      else if (v > 0) {
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p = (a: number) => `${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`;
        plot.push(<path key={i} d={`M${cx},${cy}L${p(a0)}A${r},${r} 0 ${large} 1 ${p(a1)}Z`} fill={color} stroke="#fff" strokeWidth={1} />);
        if (v / total > 0.05) {
          const m = (a0 + a1) / 2;
          plot.push(<text key={`t${i}`} x={cx + r * 0.62 * Math.cos(m)} y={cy + r * 0.62 * Math.sin(m)} fill="#fff" fontSize={11} textAnchor="middle" dominantBaseline="middle">
            {Math.round((v / total) * 100)}%</text>);
        }
      }
      a0 = a1;
    });
    plot.push(...legend.map((l, i) => (
      <g key={`l${i}`} transform={`translate(${body.w - legendW + 4},${cy - (legend.length * 16) / 2 + i * 16})`}>
        <rect width={10} height={10} y={-1} fill={l.color} />
        <text x={14} y={8} fontSize={11} fill="#444">{l.name.length > 18 ? `${l.name.slice(0, 17)}…` : l.name}</text>
      </g>
    )));
  } else {
    const all = data.series.flatMap((s) => s.values);
    const sc = niceScale(Math.min(0, ...all), Math.max(0, ...all));
    const horiz = kind === 'bar';
    const labelW = horiz ? Math.min(90, Math.max(24, ...data.cats.map((c) => c.length * 6.5))) : Math.max(...sc.ticks.map((t) => fmt(t).length * 6.5), 16);
    const pad = { l: labelW + 8, r: 12, t: 6, b: horiz ? 20 : 20 };
    const px = body.x + pad.l, py = body.y + pad.t, pw = Math.max(10, body.w - pad.l - pad.r), ph = Math.max(10, body.h - pad.t - pad.b);
    const n = Math.max(1, data.cats.length);
    const valPos = (v: number) => (v - sc.lo) / (sc.hi - sc.lo || 1);
    // Trục giá trị và lưới.
    for (const t of sc.ticks) {
      const f = valPos(t);
      if (horiz) {
        const x = px + f * pw;
        plot.push(<line key={`g${t}`} x1={x} x2={x} y1={py} y2={py + ph} stroke="#e3e3e3" />);
        plot.push(<text key={`v${t}`} x={x} y={py + ph + 14} fontSize={10} fill="#666" textAnchor="middle">{fmt(t)}</text>);
      } else {
        const y = py + ph - f * ph;
        plot.push(<line key={`g${t}`} x1={px} x2={px + pw} y1={y} y2={y} stroke="#e3e3e3" />);
        plot.push(<text key={`v${t}`} x={px - 6} y={y + 3.5} fontSize={10} fill="#666" textAnchor="end">{fmt(t)}</text>);
      }
    }
    const band = (horiz ? ph : pw) / n;
    const every = Math.max(1, Math.ceil((n * 40) / (horiz ? ph * 2.2 : pw)));
    data.cats.forEach((c, i) => {
      if (i % every) return;
      const mid = (horiz ? py : px) + band * (i + 0.5);
      const label = c.length > 14 ? `${c.slice(0, 13)}…` : c;
      plot.push(horiz
        ? <text key={`c${i}`} x={px - 6} y={mid + 3.5} fontSize={10} fill="#666" textAnchor="end">{label}</text>
        : <text key={`c${i}`} x={mid} y={py + ph + 14} fontSize={10} fill="#666" textAnchor="middle">{label}</text>);
    });
    const zero = valPos(Math.max(sc.lo, Math.min(0, sc.hi)));
    if (kind === 'col' || kind === 'bar') {
      const k = data.series.length || 1;
      const inner = band * 0.7, bw = inner / k;
      data.series.forEach((s, si) => s.values.forEach((v, i) => {
        const f = valPos(v);
        const a = Math.min(f, zero), len = Math.abs(f - zero);
        const off = (horiz ? py : px) + band * i + (band - inner) / 2 + bw * (horiz ? k - 1 - si : si);
        plot.push(horiz
          ? <rect key={`b${si}-${i}`} x={px + a * pw} y={off} width={len * pw} height={Math.max(1, bw - 1)} fill={PALETTE[si % PALETTE.length]} />
          : <rect key={`b${si}-${i}`} x={off} y={py + ph - (a + len) * ph} width={Math.max(1, bw - 1)} height={len * ph} fill={PALETTE[si % PALETTE.length]} />);
      }));
    } else {
      data.series.forEach((s, si) => {
        const color = PALETTE[si % PALETTE.length];
        const pts = s.values.map((v, i) => [px + band * (i + 0.5), py + ph - valPos(v) * ph] as const);
        if (!pts.length) return;
        const line = pts.map(([x, y], i) => `${i ? 'L' : 'M'}${x},${y}`).join('');
        if (kind === 'area') {
          const base = py + ph - zero * ph;
          plot.push(<path key={`a${si}`} d={`${line}L${pts[pts.length - 1][0]},${base}L${pts[0][0]},${base}Z`} fill={color} fillOpacity={0.55} stroke={color} />);
        } else {
          plot.push(<path key={`l${si}`} d={line} fill="none" stroke={color} strokeWidth={2.2} strokeLinejoin="round" />);
          pts.forEach(([x, y], i) => plot.push(<circle key={`p${si}-${i}`} cx={x} cy={y} r={3} fill={color} />));
        }
      });
    }
    // Trục danh mục.
    plot.push(horiz
      ? <line key="axis" x1={px + zero * pw} x2={px + zero * pw} y1={py} y2={py + ph} stroke="#999" />
      : <line key="axis" x1={px} x2={px + pw} y1={py + ph - zero * ph} y2={py + ph - zero * ph} stroke="#999" />);
    // Chú thích dưới biểu đồ.
    let lx = 0;
    const items = legend.map((l, i) => {
      const name = l.name.length > 16 ? `${l.name.slice(0, 15)}…` : l.name;
      const g = <g key={`lg${i}`} transform={`translate(${lx},0)`}><rect width={10} height={10} y={-9} fill={l.color} /><text x={14} fontSize={11} fill="#444">{name}</text></g>;
      lx += 22 + name.length * 6.5;
      return g;
    });
    plot.push(<g key="legend" transform={`translate(${Math.max(4, (w - lx) / 2)},${h - 7})`}>{items}</g>);
  }

  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ display: 'block', fontFamily: 'system-ui, sans-serif' }}>
      <rect width={w} height={h} fill="#fff" />
      {title && <text x={w / 2} y={18} fontSize={14} fill="#333" textAnchor="middle">{title}</text>}
      {plot}
    </svg>
  );
}
