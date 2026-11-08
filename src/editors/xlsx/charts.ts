// Biểu đồ trong bảng tính: mô hình, lấy dữ liệu từ ô, và đọc/ghi phần biểu đồ (DrawingML) của tệp .xlsx.
// ExcelJS không hỗ trợ biểu đồ nên ta xử lý trực tiếp các tệp XML trong gói .xlsx bằng JSZip.
import JSZip from 'jszip';
import type { Book, Engine, Merge, Sheet } from './sheetModel';
import { addr, colIndex, formatValue, key } from './sheetModel';

export type ChartKind = 'col' | 'bar' | 'line' | 'area' | 'pie';

export const CHART_KINDS: [ChartKind, string][] = [
  ['col', 'Cột'], ['bar', 'Thanh ngang'], ['line', 'Đường'], ['area', 'Vùng'], ['pie', 'Tròn'],
];

export interface ChartSeries {
  /** Tham chiếu ô chứa tên chuỗi, ví dụ "Sheet1!$B$1". */
  name?: string;
  /** Tham chiếu vùng nhãn danh mục, ví dụ "Sheet1!$A$2:$A$5". */
  cat?: string;
  /** Tham chiếu vùng giá trị. */
  val: string;
}

export interface Chart {
  id: string;
  kind: ChartKind;
  title?: string;
  series: ChartSeries[];
  /** Vị trí theo ô: góc trên trái (r0, c0) đến góc dưới phải (r1, c1), không bao gồm hàng/cột r1/c1. */
  anchor: Merge;
}

export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

const quoteSheet = (name: string) => (/^[A-Za-z_][\w.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);

/** Tham chiếu tuyệt đối có tên trang tính: 'Doanh thu'!$A$1:$A$5. */
export function absRef(sheet: string, r0: number, c0: number, r1 = r0, c1 = c0): string {
  const a = (r: number, c: number) => addr(r, c).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  return `${quoteSheet(sheet)}!${a(r0, c0)}${r1 !== r0 || c1 !== c0 ? `:${a(r1, c1)}` : ''}`;
}

/** "'Doanh thu'!$A$1:$B$3" → tên trang và vùng. */
export function parseRef(ref: string): { sheet?: string; m: Merge } | null {
  const x = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.trim());
  if (!x) return null;
  const sheet = x[1]?.replace(/''/g, "'") ?? x[2];
  const r0 = Number(x[4]) - 1, c0 = colIndex(x[3].toUpperCase());
  const r1 = x[6] ? Number(x[6]) - 1 : r0, c1 = x[5] ? colIndex(x[5].toUpperCase()) : c0;
  return { sheet, m: { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) } };
}

function refCells(book: Book, ref: string | undefined, si: number): [number, number, number][] {
  if (!ref) return [];
  const p = parseRef(ref);
  if (!p) return [];
  const idx = p.sheet ? book.sheets.findIndex((s) => s.name.toLowerCase() === p.sheet!.toLowerCase()) : si;
  if (idx < 0) return [];
  const out: [number, number, number][] = [];
  for (let r = p.m.r0; r <= p.m.r1; r++) for (let c = p.m.c0; c <= p.m.c1; c++) out.push([idx, r, c]);
  return out;
}

export interface ChartData { cats: string[]; series: { name: string; values: number[] }[] }

/** Lấy dữ liệu hiện tại của biểu đồ từ các ô (đã tính công thức). */
export function chartData(chart: Chart, book: Book, engine: Engine, si: number): ChartData {
  const shown = (s: number, r: number, c: number) => {
    const cell = book.sheets[s]?.cells[key(r, c)];
    return formatValue(engine.valueAt(s, r, c), cell?.s?.numFmt);
  };
  let cats: string[] = [];
  const series = chart.series.map((ser, i) => {
    const vals = refCells(book, ser.val, si).map(([s, r, c]) => {
      const v = engine.valueAt(s, r, c);
      return typeof v === 'number' && Number.isFinite(v) ? v : 0;
    });
    if (!cats.length && ser.cat) cats = refCells(book, ser.cat, si).map(([s, r, c]) => shown(s, r, c));
    const nameCell = refCells(book, ser.name, si)[0];
    const name = nameCell ? shown(...nameCell) : ser.name && !parseRef(ser.name) ? ser.name : `Chuỗi ${i + 1}`;
    return { name, values: vals };
  });
  const n = Math.max(0, ...series.map((s) => s.values.length));
  if (cats.length < n) cats = Array.from({ length: n }, (_, i) => cats[i] ?? String(i + 1));
  return { cats: cats.slice(0, n), series };
}

/**
 * Tạo biểu đồ từ vùng chọn như Excel: hàng đầu là tên chuỗi nếu là chữ, cột đầu là nhãn nếu là chữ;
 * mỗi cột là một chuỗi khi vùng cao hơn rộng, ngược lại mỗi hàng là một chuỗi.
 */
export function chartFromRange(sheet: Sheet, engine: Engine, si: number, m: Merge, kind: ChartKind, id: string): Chart | null {
  const val = (r: number, c: number) => engine.valueAt(si, r, c);
  const isText = (v: unknown) => v !== null && v !== '' && typeof v !== 'number';
  const isEmpty = (v: unknown) => v === null || v === '';
  const { r0, c0, r1, c1 } = m;
  let headRow = r0 < r1, headCol = c0 < c1;
  for (let c = c0 + 1; c <= c1 && headRow; c++) if (!isText(val(r0, c)) && !isEmpty(val(r0, c))) headRow = false;
  if (headRow && ![...Array(c1 - c0)].some((_, i) => isText(val(r0, c0 + 1 + i)))) headRow = false;
  for (let r = r0 + (headRow ? 1 : 0); r <= r1 && headCol; r++) if (!isText(val(r, c0)) && !isEmpty(val(r, c0))) headCol = false;
  if (headCol && ![...Array(r1 - r0 + 1)].some((_, i) => isText(val(r0 + i, c0)))) headCol = false;
  // Một hàng hoặc một cột toàn chữ ở góc: ô góc thuộc cả hai, không phải dữ liệu.
  const dr0 = r0 + (headRow ? 1 : 0), dc0 = c0 + (headCol ? 1 : 0);
  if (dr0 > r1 || dc0 > c1) return null;
  const byCol = r1 - dr0 >= c1 - dc0;
  const series: ChartSeries[] = [];
  const name = sheet.name;
  if (byCol) {
    for (let c = dc0; c <= c1; c++) series.push({
      name: headRow ? absRef(name, r0, c) : undefined,
      cat: headCol ? absRef(name, dr0, c0, r1, c0) : undefined,
      val: absRef(name, dr0, c, r1, c),
    });
  } else {
    for (let r = dr0; r <= r1; r++) series.push({
      name: headCol ? absRef(name, r, c0) : undefined,
      cat: headRow ? absRef(name, r0, dc0, r0, c1) : undefined,
      val: absRef(name, r, dc0, r, c1),
    });
  }
  if (kind === 'pie') series.splice(1);
  // Một chuỗi có tên: dùng tên đó làm tiêu đề như Excel.
  const only = series.length === 1 && series[0].name ? parseRef(series[0].name) : null;
  const title = only ? formatValue(val(only.m.r0, only.m.c0)) : '';
  return { id, kind, title: title || undefined, series, anchor: { r0, c0, r1, c1 } };
}

// ---------------- Đọc/ghi XML ----------------

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** Đường dẫn tuyệt đối trong gói từ thư mục gốc và đích tương đối trong .rels. */
function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

const relsPath = (p: string) => { const i = p.lastIndexOf('/'); return `${p.slice(0, i)}/_rels/${p.slice(i + 1)}.rels`; };

function parseRels(xml: string | undefined): Map<string, { target: string; type: string }> {
  const out = new Map<string, { target: string; type: string }>();
  for (const m of (xml ?? '').matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
    const type = /\bType="([^"]+)"/.exec(m[0])?.[1] ?? '';
    if (id && target) out.set(id, { target: unesc(target), type });
  }
  return out;
}

/** Danh sách trang tính theo thứ tự: tên và đường dẫn tệp XML. */
async function sheetParts(zip: JSZip): Promise<{ name: string; path: string }[]> {
  const wb = await zip.file('xl/workbook.xml')?.async('string');
  const rels = parseRels(await zip.file('xl/_rels/workbook.xml.rels')?.async('string'));
  const out: { name: string; path: string }[] = [];
  for (const m of (wb ?? '').matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(m[0])?.[1];
    const rid = /\br:id="([^"]+)"/.exec(m[0])?.[1] ?? /\b\w+:id="([^"]+)"/.exec(m[0])?.[1];
    const rel = rid ? rels.get(rid) : undefined;
    if (name !== undefined && rel) out.push({ name: unesc(name), path: resolvePath('xl/workbook.xml', rel.target) });
  }
  return out;
}

const tag = (xml: string, name: string) => new RegExp(`<(?:\\w+:)?${name}\\b[^>]*?(?:/>|>([\\s\\S]*?)</(?:\\w+:)?${name}>)`).exec(xml);
const tagAll = (xml: string, name: string) => [...xml.matchAll(new RegExp(`<(?:\\w+:)?${name}\\b[^>]*?(?:/>|>([\\s\\S]*?)</(?:\\w+:)?${name}>)`, 'g'))];
const attr = (el: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(el)?.[1];
const formulaIn = (xml?: string) => { const f = xml ? tag(xml, 'f')?.[1] : undefined; return f ? unesc(f.trim()) : undefined; };

/** Đọc một tệp chart*.xml thành mô hình (chỉ lấy loại, tiêu đề và tham chiếu dữ liệu). */
export function parseChartXml(xml: string): Omit<Chart, 'id' | 'anchor'> | null {
  const plot = tag(xml, 'plotArea')?.[1] ?? '';
  let kind: ChartKind | null = null;
  let body = '';
  for (const [t, k] of [['barChart', 'col'], ['bar3DChart', 'col'], ['lineChart', 'line'], ['line3DChart', 'line'], ['areaChart', 'area'],
    ['area3DChart', 'area'], ['pieChart', 'pie'], ['pie3DChart', 'pie'], ['doughnutChart', 'pie'], ['scatterChart', 'line'], ['radarChart', 'line']] as [string, ChartKind][]) {
    const m = tag(plot, t);
    if (m) { kind = k; body = m[1] ?? ''; break; }
  }
  if (!kind) return null;
  if (kind === 'col' && attr(tag(body, 'barDir')?.[0] ?? '', 'val') === 'bar') kind = 'bar';
  const series: ChartSeries[] = [];
  for (const s of tagAll(body, 'ser')) {
    const ser = s[1] ?? '';
    const val = formulaIn(tag(ser, 'val')?.[1] ?? tag(ser, 'yVal')?.[1]);
    if (!val) continue;
    const tx = tag(ser, 'tx')?.[1];
    const literal = tx ? tag(tx, 'v')?.[1] : undefined;
    series.push({
      name: formulaIn(tx) ?? (literal ? unesc(literal) : undefined),
      cat: formulaIn(tag(ser, 'cat')?.[1] ?? tag(ser, 'xVal')?.[1]),
      val,
    });
  }
  if (!series.length) return null;
  const t = tag(tag(xml, 'chart')?.[1] ?? xml, 'title');
  const title = t ? [...(t[1] ?? '').matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((x) => unesc(x[1])).join('') : '';
  return { kind, title: title || undefined, series };
}

const num = (xml: string, name: string) => Number(tag(xml, name)?.[1] ?? 0);

/** Đọc các biểu đồ của từng trang tính (theo tên trang). */
export async function readCharts(data: ArrayBuffer, newId: () => string): Promise<Map<string, Chart[]>> {
  const out = new Map<string, Chart[]>();
  const zip = await JSZip.loadAsync(data);
  for (const sh of await sheetParts(zip)) {
    const sheetXml = await zip.file(sh.path)?.async('string');
    const drawId = sheetXml ? attr(tag(sheetXml, 'drawing')?.[0] ?? '', 'r:id') : undefined;
    if (!drawId) continue;
    const rel = parseRels(await zip.file(relsPath(sh.path))?.async('string')).get(drawId);
    if (!rel) continue;
    const drawPath = resolvePath(sh.path, rel.target);
    const drawXml = await zip.file(drawPath)?.async('string');
    if (!drawXml) continue;
    const drawRels = parseRels(await zip.file(relsPath(drawPath))?.async('string'));
    const charts: Chart[] = [];
    for (const a of [...tagAll(drawXml, 'twoCellAnchor'), ...tagAll(drawXml, 'oneCellAnchor')]) {
      const body = a[1] ?? '';
      const cid = /<c:chart\b[^>]*\br:id="([^"]+)"/.exec(body)?.[1];
      const crel = cid ? drawRels.get(cid) : undefined;
      if (!crel) continue;
      const chartXml = await zip.file(resolvePath(drawPath, crel.target))?.async('string');
      const parsed = chartXml ? parseChartXml(chartXml) : null;
      if (!parsed) continue;
      const from = tag(body, 'from')?.[1] ?? '';
      const to = tag(body, 'to')?.[1];
      const r0 = num(from, 'row'), c0 = num(from, 'col');
      const anchor = to
        ? { r0, c0, r1: Math.max(r0 + 1, num(to, 'row') + (num(to, 'rowOff') > 0 ? 1 : 0)), c1: Math.max(c0 + 1, num(to, 'col') + (num(to, 'colOff') > 0 ? 1 : 0)) }
        : { r0, c0, r1: r0 + 15, c1: c0 + 6 };
      charts.push({ id: newId(), ...parsed, anchor });
    }
    if (charts.length) out.set(sh.name, charts);
  }
  return out;
}

function cacheXml(kind: 'str' | 'num', values: (string | number)[]) {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(String(v))}</c:v></c:pt>`).join('');
  return `<c:${kind}Cache>${kind === 'num' ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${values.length}"/>${pts}</c:${kind}Cache>`;
}

/** Tạo chart*.xml (kèm giá trị lưu sẵn để trình xem không tính lại vẫn hiện đúng). */
export function chartXml(chart: Chart, data: ChartData): string {
  const isPie = chart.kind === 'pie';
  const sers = chart.series.map((s, i) => {
    const d = data.series[i] ?? { name: '', values: [] };
    const color = PALETTE[i % PALETTE.length].slice(1);
    const tx = s.name && parseRef(s.name)
      ? `<c:tx><c:strRef><c:f>${esc(s.name)}</c:f>${cacheXml('str', [d.name])}</c:strRef></c:tx>`
      : s.name ? `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>` : '';
    const sp = isPie ? '' : chart.kind === 'line'
      ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>`
      : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
    const pts = isPie ? data.cats.map((_, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${PALETTE[j % PALETTE.length].slice(1)}"/></a:solidFill></c:spPr></c:dPt>`).join('') : '';
    const cat = s.cat ? `<c:cat><c:strRef><c:f>${esc(s.cat)}</c:f>${cacheXml('str', data.cats)}</c:strRef></c:cat>` : '';
    const val = `<c:val><c:numRef><c:f>${esc(s.val)}</c:f>${cacheXml('num', d.values)}</c:numRef></c:val>`;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${pts}${cat}${val}${chart.kind === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
  }).join('');
  const axIds = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  let plot: string;
  if (chart.kind === 'col' || chart.kind === 'bar') {
    plot = `<c:barChart><c:barDir val="${chart.kind}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="150"/>${axIds}</c:barChart>`;
  } else if (chart.kind === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>`;
  } else if (chart.kind === 'area') {
    plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${axIds}</c:areaChart>`;
  } else {
    plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  const horiz = chart.kind === 'bar';
  const axes = isPie ? '' :
    `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'l' : 'b'}"/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'b' : 'l'}"/><c:majorGridlines/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="vi-VN" sz="1400" b="0"/><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>` +
    `<c:legend><c:legendPos val="${isPie ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

function anchorXml(chart: Chart, rid: string, id: number) {
  const { r0, c0, r1, c1 } = chart.anchor;
  const pos = (t: string, r: number, c: number) => `<xdr:${t}><xdr:col>${c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${t}>`;
  return `<xdr:twoCellAnchor editAs="oneCell">${pos('from', r0, c0)}${pos('to', r1, c1)}` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Biểu đồ ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/>` +
    '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
}

const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const RELS_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

function nextRid(rels: Map<string, unknown>) {
  let n = rels.size + 1;
  while (rels.has(`rId${n}`)) n++;
  return `rId${n}`;
}

function addRel(xml: string | undefined, id: string, type: string, target: string) {
  const rel = `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/>`;
  return xml && xml.includes('</Relationships>') ? xml.replace('</Relationships>', `${rel}</Relationships>`) : `${RELS_HEAD}${rel}</Relationships>`;
}

// Các phần tử phải đứng sau <drawing> trong worksheet theo thứ tự của lược đồ.
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/**
 * Ghi biểu đồ vào tệp .xlsx do ExcelJS tạo. `charts`: tên trang tính → các biểu đồ kèm dữ liệu hiện tại.
 * Nếu trang đã có bản vẽ (ví dụ ảnh), biểu đồ được thêm vào bản vẽ đó.
 */
export async function writeCharts(buf: ArrayBuffer, charts: Map<string, { chart: Chart; data: ChartData }[]>): Promise<ArrayBuffer> {
  if (![...charts.values()].some((c) => c.length)) return buf;
  const zip = await JSZip.loadAsync(buf);
  const parts = await sheetParts(zip);
  let ct = (await zip.file('[Content_Types].xml')!.async('string'));
  let chartNo = 1;
  while (zip.file(`xl/charts/chart${chartNo}.xml`)) chartNo++;
  let drawNo = 1;
  const nextDrawing = () => { while (zip.file(`xl/drawings/drawing${drawNo}.xml`)) drawNo++; return drawNo; };
  for (let si = 0; si < parts.length; si++) {
    const list = charts.get(parts[si].name) ?? [];
    if (!list.length) continue;
    const sp = parts[si].path;
    let sheetXml = await zip.file(sp)!.async('string');
    const sRelsPath = relsPath(sp);
    let sRelsXml = await zip.file(sRelsPath)?.async('string');
    const sRels = parseRels(sRelsXml);
    const existingId = attr(tag(sheetXml, 'drawing')?.[0] ?? '', 'r:id');
    let drawPath: string;
    let drawXml: string;
    if (existingId && sRels.get(existingId)) {
      drawPath = resolvePath(sp, sRels.get(existingId)!.target);
      drawXml = (await zip.file(drawPath)?.async('string')) ?? '';
    } else {
      drawPath = `xl/drawings/drawing${nextDrawing()}.xml`;
      drawXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>';
      const rid = nextRid(sRels);
      sRelsXml = addRel(sRelsXml, rid, REL_DRAWING, `../drawings/${drawPath.split('/').pop()}`);
      zip.file(sRelsPath, sRelsXml);
      if (!/xmlns:r="/.test(sheetXml.slice(0, sheetXml.indexOf('>', sheetXml.indexOf('<worksheet')) + 1))) {
        sheetXml = sheetXml.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
      }
      const el = `<drawing r:id="${rid}"/>`;
      const after = AFTER_DRAWING.map((t) => sheetXml.search(new RegExp(`<${t}\\b`))).filter((i) => i >= 0);
      const at = after.length ? Math.min(...after) : sheetXml.lastIndexOf('</worksheet>');
      sheetXml = sheetXml.slice(0, at) + el + sheetXml.slice(at);
      zip.file(sp, sheetXml);
      ct = ct.replace('</Types>', `<Override PartName="/${drawPath}" ContentType="${CT_DRAWING}"/></Types>`);
    }
    const dRelsPath = relsPath(drawPath);
    let dRelsXml = await zip.file(dRelsPath)?.async('string');
    const dRels = parseRels(dRelsXml);
    let maxId = Math.max(1, ...[...drawXml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])));
    let anchors = '';
    for (const { chart, data } of list) {
      const path = `xl/charts/chart${chartNo++}.xml`;
      zip.file(path, chartXml(chart, data));
      ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="${CT_CHART}"/></Types>`);
      const rid = nextRid(dRels);
      dRels.set(rid, { target: '', type: REL_CHART });
      dRelsXml = addRel(dRelsXml, rid, REL_CHART, `../charts/${path.split('/').pop()}`);
      anchors += anchorXml(chart, rid, ++maxId);
    }
    zip.file(dRelsPath, dRelsXml!);
    drawXml = /<\/xdr:wsDr>/.test(drawXml) ? drawXml.replace('</xdr:wsDr>', `${anchors}</xdr:wsDr>`) : drawXml.replace(/<(\w+:)?wsDr\b([^>]*)\/>/, `<xdr:wsDr$2>${anchors}</xdr:wsDr>`);
    zip.file(drawPath, drawXml);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
