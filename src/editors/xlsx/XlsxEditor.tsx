import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { OpenedFile, saveFile } from '../../lib/fileio';
import { EditorShell, Sep, ToolButton } from '../../components/EditorShell';
import { uid } from '../../lib/uid';
import {
  Book, Cell, CellStyle, Engine, Merge, Sheet, addr, colName, emptySheet, findMatches, formatValue, key, mergeAt,
  moveFormula, parseInput, parseKey, rawText, replaceInCell, shiftFormula, shiftMerges, shiftSpan, usedRange,
} from './sheetModel';
import { CHART_KINDS, Chart, ChartKind, chartData, chartFromRange } from './charts';
import { CF_KINDS, CfKind, CfRule, Filter, cfEvaluator, columnValues, filterHidden } from './cf';
import ChartView from './ChartView';
import { LoadedBook, loadXlsx, saveXlsx } from './xlsxIO';
import { csvToBook } from './csv';

const ROW_H = 24;
const DEFAULT_W = 96;
const HEAD_W = 48;
const HEAD_H = 24;

interface Sel { r: number; c: number; r2: number; c2: number }
const norm = (s: Sel) => ({ r0: Math.min(s.r, s.r2), r1: Math.max(s.r, s.r2), c0: Math.min(s.c, s.c2), c1: Math.max(s.c, s.c2) });

/** Chỉ số i lớn nhất có xs[i] <= v (xs tăng dần, là toạ độ bắt đầu của từng hàng/cột). */
function indexAt(xs: number[], v: number) {
  let lo = 0, hi = xs.length - 2;
  if (hi < 0) return 0;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (xs[mid] <= v) lo = mid; else hi = mid - 1;
  }
  return lo;
}

/** Chỉ số có xs[i] gần v nhất (để bám biểu đồ vào lưới ô). */
function nearestIdx(xs: number[], v: number) {
  const i = indexAt(xs, v);
  return i + 1 < xs.length && v - xs[i] > xs[i + 1] - v ? i + 1 : i;
}

const NUM_FORMATS: [string, string][] = [
  ['General', 'Chung'], ['#,##0', 'Số 1.234'], ['#,##0.00', 'Số 1.234,56'], ['0%', 'Phần trăm'],
  ['0.00%', 'Phần trăm 0,00'], ['#,##0 "₫"', 'Tiền VNĐ'], ['"$"#,##0.00', 'Tiền USD'], ['dd/mm/yyyy', 'Ngày'],
];

export default function XlsxEditor({ file, onClose }: { file: OpenedFile; onClose: () => void }) {
  const isNew = file.data.byteLength === 0;
  const [book, setBook] = useState<Book | null>(isNew ? { sheets: [emptySheet('Sheet1', uid())] } : null);
  const loaded = useRef<LoadedBook | null>(null);
  const [active, setActive] = useState(0);
  const [sel, setSel] = useState<Sel>({ r: 0, c: 0, r2: 0, c2: 0 });
  const [editing, setEditing] = useState<{ text: string; inBar: boolean } | null>(null);
  const [history, setHistory] = useState<Book[]>([]);
  const [future, setFuture] = useState<Book[]>([]);
  const [dirty, setDirty] = useState(false);
  const [busy, setBusy] = useState<string | null>(isNew ? null : 'Đang mở…');
  const [error, setError] = useState<string | null>(null);
  const isCsv = /\.(csv|tsv)$/i.test(file.name);
  // CSV được lưu thành .xlsx (giữ được định dạng, công thức, nhiều trang): lưu lần đầu hỏi nơi lưu mới.
  const [path, setPath] = useState(isCsv ? undefined : file.path);
  const [name, setName] = useState(isCsv ? file.name.replace(/\.(csv|tsv)$/i, '.xlsx') : file.name);
  const clip = useRef<{ text: string; r: number; c: number; cells: (Cell | undefined)[][] } | null>(null);
  const [find, setFind] = useState<{ q: string; repl: string; idx: number; replace: boolean } | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [selChart, setSelChart] = useState<string | null>(null);
  const hiddenRef = useRef<Record<number, boolean>>({});
  /** Bảng chọn giá trị lọc của một cột. */
  const [fpop, setFpop] = useState<{ c: number; q: string; hide: Set<string> } | null>(null);
  /** Hộp thoại định dạng có điều kiện. */
  const [cfDlg, setCfDlg] = useState<{ kind: CfKind; a: string; b: string; fill: string; color: string; c1: string; c2: string; c3: string; three: boolean } | null>(null);
  const chartDrag = useRef<{ id: string; mode: 'move' | 'size'; x: number; y: number; moved: boolean } | null>(null);
  const [chartOff, setChartOff] = useState<{ id: string; dx: number; dy: number; mode: 'move' | 'size' } | null>(null);

  useEffect(() => {
    if (isNew) return;
    if (isCsv) {
      try { setBook(csvToBook(file.data, file.name, uid())); setNotice('Tệp CSV sẽ được lưu thành bảng tính .xlsx'); }
      catch (e) { setError(`Không đọc được tệp CSV: ${(e as Error).message}`); }
      setBusy(null);
      return;
    }
    loadXlsx(file.data)
      .then((l) => { loaded.current = l; setBook(l.book); })
      .catch((e) => setError(`Không mở được tệp Excel: ${e.message ?? e}`))
      .finally(() => setBusy(null));
  }, []);

  const engine = useMemo(() => (book ? new Engine(book) : null), [book]);
  const sheet = book?.sheets[active];

  const commit = useCallback((fn: (b: Book) => Book) => {
    setBook((b) => {
      if (!b) return b;
      setHistory((h) => [...h.slice(-49), b]);
      setFuture([]);
      return fn(b);
    });
    setDirty(true);
  }, []);

  const updateSheet = (fn: (s: Sheet) => Sheet) =>
    commit((b) => ({ sheets: b.sheets.map((s, i) => (i === active ? fn(s) : s)) }));

  const setCells = (entries: [number, number, Cell | undefined][]) =>
    updateSheet((s) => {
      const cells = { ...s.cells };
      let { rowCount, colCount } = s;
      for (const [r, c, cell] of entries) {
        const k = key(r, c);
        if (!cell || (cell.v == null && !cell.f && !cell.s)) delete cells[k];
        else cells[k] = cell;
        rowCount = Math.max(rowCount, r + 20);
        colCount = Math.max(colCount, c + 5);
      }
      return { ...s, cells, rowCount, colCount };
    });

  const undo = () => {
    const prev = history[history.length - 1];
    if (!prev || !book) return;
    setFuture((f) => [book, ...f]);
    setHistory((h) => h.slice(0, -1));
    setBook(prev);
    setActive((a) => Math.min(a, prev.sheets.length - 1));
  };
  const redo = () => {
    const next = future[0];
    if (!next || !book) return;
    setHistory((h) => [...h, book]);
    setFuture((f) => f.slice(1));
    setBook(next);
  };

  const commitEdit = (move?: [number, number]) => {
    if (!editing || !sheet) return;
    const old = sheet.cells[key(sel.r, sel.c)];
    if (editing.text !== rawText(old)) {
      setCells([[sel.r, sel.c, { ...parseInput(editing.text), s: old?.s }]]);
    }
    setEditing(null);
    if (move) moveSel(move[0], move[1]);
    gridRef.current?.focus();
  };

  /** Chọn một ô; nếu ô nằm trong vùng gộp thì chọn cả vùng. */
  const selectCell = (r: number, c: number): Sel => {
    const m = mergeAt(sheet?.merges, r, c);
    return m ? { r: m.r0, c: m.c0, r2: m.r1, c2: m.c1 } : { r, c, r2: r, c2: c };
  };

  const moveSel = (dr: number, dc: number, extend = false) => {
    setSel((s) => {
      if (extend) {
        let r2 = Math.max(0, s.r2 + dr);
        while (dr && r2 > 0 && hiddenRef.current[r2]) r2 += dr > 0 ? 1 : -1;
        return { ...s, r2, c2: Math.max(0, s.c2 + dc) };
      }
      const cur = mergeAt(sheet?.merges, s.r, s.c);
      let r = s.r + dr, c = s.c + dc;
      if (cur && dr > 0) r = cur.r1 + 1;
      if (cur && dc > 0) c = cur.c1 + 1;
      // Bỏ qua hàng đang ẩn (do lọc hoặc ẩn thủ công).
      while (dr && r > 0 && hiddenRef.current[r]) r += dr > 0 ? 1 : -1;
      return selectCell(Math.max(0, r), Math.max(0, c));
    });
  };

  const forRange = (fn: (r: number, c: number) => [number, number, Cell | undefined] | null) => {
    const { r0, r1, c0, c1 } = norm(sel);
    const out: [number, number, Cell | undefined][] = [];
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) { const e = fn(r, c); if (e) out.push(e); }
    setCells(out);
  };

  const applyStyle = (patch: Partial<CellStyle> | ((s: CellStyle) => Partial<CellStyle>)) =>
    forRange((r, c) => {
      const cell = sheet!.cells[key(r, c)] ?? {};
      const s = cell.s ?? {};
      const p = typeof patch === 'function' ? patch(s) : patch;
      return [r, c, { ...cell, s: { ...s, ...p } }];
    });

  const activeCell = sheet?.cells[key(sel.r, sel.c)];
  const toggle = (k: 'bold' | 'italic' | 'underline') => {
    const on = !activeCell?.s?.[k];
    applyStyle({ [k]: on });
  };

  const clearRange = () => forRange((r, c) => {
    const cell = sheet!.cells[key(r, c)];
    return cell ? [r, c, cell.s ? { s: cell.s } : undefined] : null;
  });

  const rangeTSV = () => {
    const { r0, r1, c0, c1 } = norm(sel);
    const rows: string[] = [];
    const cells: (Cell | undefined)[][] = [];
    for (let r = r0; r <= r1; r++) {
      const line: string[] = [];
      const row: (Cell | undefined)[] = [];
      for (let c = c0; c <= c1; c++) {
        const cell = sheet!.cells[key(r, c)];
        row.push(cell);
        line.push(cell?.f ? formatValue(engine!.evaluate(active, r, c), cell.s?.numFmt) : rawText(cell));
      }
      rows.push(line.join('\t'));
      cells.push(row);
    }
    const text = rows.join('\n');
    clip.current = { text, r: r0, c: c0, cells };
    return text;
  };

  const paste = (text: string) => {
    const out: [number, number, Cell | undefined][] = [];
    if (clip.current && clip.current.text === text) {
      const { cells, r, c } = clip.current;
      cells.forEach((row, i) => row.forEach((cell, j) => {
        out.push([sel.r + i, sel.c + j, cell?.f ? { ...cell, f: moveFormula(cell.f, sel.r - r, sel.c - c) } : cell]);
      }));
    } else {
      text.replace(/\r/g, '').replace(/\n$/, '').split('\n').forEach((line, i) =>
        line.split('\t').forEach((val, j) => {
          const old = sheet!.cells[key(sel.r + i, sel.c + j)];
          out.push([sel.r + i, sel.c + j, { ...parseInput(val), s: old?.s }]);
        }));
    }
    setCells(out);
  };

  const fill = (dir: 'down' | 'right') => {
    const { r0, r1, c0, c1 } = norm(sel);
    forRange((r, c) => {
      const [sr, sc] = dir === 'down' ? [r0, c] : [r, c0];
      if (r === sr && c === sc) return null;
      const src = sheet!.cells[key(sr, sc)];
      return [r, c, src?.f ? { ...src, f: moveFormula(src.f, r - sr, c - sc) } : src];
    });
    void r1; void c1;
  };

  const autoSum = () => {
    if (!sheet || !engine) return;
    const isNum = (r: number, c: number) => typeof engine.valueAt(active, r, c) === 'number';
    let r = sel.r - 1;
    while (r >= 0 && isNum(r, sel.c)) r--;
    if (r < sel.r - 1) return setCells([[sel.r, sel.c, { f: `SUM(${addr(r + 1, sel.c)}:${addr(sel.r - 1, sel.c)})`, s: activeCell?.s }]]);
    let c = sel.c - 1;
    while (c >= 0 && isNum(sel.r, c)) c--;
    if (c < sel.c - 1) return setCells([[sel.r, sel.c, { f: `SUM(${addr(sel.r, c + 1)}:${addr(sel.r, sel.c - 1)})`, s: activeCell?.s }]]);
    setEditing({ text: '=SUM()', inBar: false });
  };

  /** Chèn/xoá hàng/cột và dịch công thức ở mọi trang tính. */
  const shift = (axis: 'row' | 'col', delta: number) => {
    const { r0, r1, c0, c1 } = norm(sel);
    const at = axis === 'row' ? r0 : c0;
    const count = axis === 'row' ? r1 - r0 + 1 : c1 - c0 + 1;
    const d = delta > 0 ? count : -count;
    const target = sheet!.name;
    commit((b) => ({
      sheets: b.sheets.map((s, i) => {
        const cells: Record<string, Cell> = {};
        for (const k in s.cells) {
          let [r, c] = parseKey(k);
          let cell = s.cells[k];
          if (cell.f) cell = { ...cell, f: shiftFormula(cell.f, s.name, target, axis, at, d) };
          if (i === active) {
            const pos = axis === 'row' ? r : c;
            if (pos >= at) {
              if (d < 0 && pos < at - d) continue;
              if (axis === 'row') r += d; else c += d;
            }
          }
          cells[key(r, c)] = cell;
        }
        const refShift = (f?: string) => (f ? shiftFormula(f, s.name, target, axis, at, d) : f);
        s = { ...s, charts: s.charts?.map((ch) => ({ ...ch, series: ch.series.map((x) => ({ name: refShift(x.name), cat: refShift(x.cat), val: refShift(x.val)! })) })) };
        if (i !== active) return { ...s, cells };
        let colWidths = s.colWidths;
        let rowHeights = s.rowHeights;
        if (axis === 'row' && rowHeights) {
          rowHeights = {};
          for (const [ri, h] of Object.entries(s.rowHeights!)) {
            const n = Number(ri);
            if (n < at) rowHeights[n] = h;
            else if (d > 0) rowHeights[n + d] = h;
            else if (n >= at - d) rowHeights[n + d] = h;
          }
        }
        if (axis === 'col') {
          colWidths = {};
          for (const [ci, w] of Object.entries(s.colWidths)) {
            const n = Number(ci);
            if (n < at) colWidths[n] = w;
            else if (d > 0) colWidths[n + d] = w;
            else if (n >= at - d) colWidths[n + d] = w;
          }
        }
        const charts = s.charts?.flatMap((ch) => {
          const a = ch.anchor;
          // Neo biểu đồ: góc trên trái dịch theo hàng/cột; góc dưới phải là ranh giới (không bao gồm).
          const span = axis === 'row' ? shiftSpan(a.r0, a.r1 - 1, at, d) : shiftSpan(a.c0, a.c1 - 1, at, d);
          if (!span) return [];
          const anchor = axis === 'row' ? { ...a, r0: span[0], r1: span[1] + 1 } : { ...a, c0: span[0], c1: span[1] + 1 };
          return [{ ...ch, anchor }];
        });
        const sp = (a0: number, a1: number) => shiftSpan(a0, a1, at, d);
        const cf = s.cf?.flatMap((x) => {
          const span = axis === 'row' ? sp(x.range.r0, x.range.r1) : sp(x.range.c0, x.range.c1);
          if (!span) return [];
          return [{ ...x, range: axis === 'row' ? { ...x.range, r0: span[0], r1: span[1] } : { ...x.range, c0: span[0], c1: span[1] } }];
        });
        let filter = s.filter;
        if (filter) {
          const span = axis === 'row' ? sp(filter.r, filter.r1) : sp(filter.c0, filter.c1);
          const headDeleted = axis === 'row' && d < 0 && filter.r >= at && filter.r < at - d;
          if (!span || headDeleted) filter = undefined;
          else if (axis === 'row') filter = { ...filter, r: span[0], r1: span[1] };
          else {
            const hide: Record<number, string[]> = {};
            for (const [ci, v] of Object.entries(filter.hide)) { const n = sp(Number(ci), Number(ci)); if (n) hide[n[0]] = v; }
            filter = { ...filter, c0: span[0], c1: span[1], hide };
          }
        }
        let hiddenRows = s.hiddenRows;
        if (axis === 'row' && hiddenRows) {
          hiddenRows = {};
          for (const ri of Object.keys(s.hiddenRows!)) { const n = sp(Number(ri), Number(ri)); if (n) hiddenRows[n[0]] = true; }
        }
        return {
          ...s, cells, colWidths, rowHeights, merges: shiftMerges(s.merges, axis, at, d), charts, cf, filter, hiddenRows,
          rowCount: s.rowCount + (axis === 'row' && d > 0 ? d : 0), colCount: s.colCount + (axis === 'col' && d > 0 ? d : 0),
        };
      }),
    }));
  };

  const sort = (asc: boolean) => {
    if (!sheet || !engine) return;
    let { r0, r1, c0, c1 } = norm(sel);
    if (r0 === r1) {
      const f = sheet.filter;
      if (f && sel.c >= f.c0 && sel.c <= f.c1) ({ r0, r1, c0, c1 } = { r0: f.r + 1, r1: f.r1, c0: f.c0, c1: f.c1 });
      else {
        const u = usedRange(sheet);
        r0 = 1; r1 = u.maxR; c0 = 0; c1 = u.maxC;
      }
      if (r1 < r0) return;
    }
    sortRange(asc, { r0, r1, c0, c1 }, sel.c);
  };

  const sortRange = (asc: boolean, { r0, r1, c0, c1 }: { r0: number; r1: number; c0: number; c1: number }, col: number) => {
    if (!sheet || !engine) return;
    const rows = [];
    for (let r = r0; r <= r1; r++) rows.push(r);
    const val = (r: number) => engine.valueAt(active, r, col);
    rows.sort((a, b) => {
      const va = val(a), vb = val(b);
      if (va == null || va === '') return 1;
      if (vb == null || vb === '') return -1;
      const cmp = typeof va === 'number' && typeof vb === 'number' ? va - vb : String(va).localeCompare(String(vb), 'vi');
      return asc ? cmp : -cmp;
    });
    const out: [number, number, Cell | undefined][] = [];
    rows.forEach((src, i) => {
      const dst = r0 + i;
      for (let c = c0; c <= c1; c++) {
        const cell = sheet.cells[key(src, c)];
        out.push([dst, c, cell?.f ? { ...cell, f: moveFormula(cell.f, dst - src, 0) } : cell]);
      }
    });
    setCells(out);
  };

  /** Gộp vùng chọn (giữ dữ liệu ô trên cùng bên trái, căn giữa) hoặc bỏ gộp nếu vùng chọn đang chứa ô gộp. */
  const toggleMerge = () => {
    if (!sheet) return;
    const { r0, r1, c0, c1 } = norm(sel);
    const ms = sheet.merges ?? [];
    const hit = ms.filter((m) => !(m.r1 < r0 || m.r0 > r1 || m.c1 < c0 || m.c0 > c1));
    if (hit.length) return updateSheet((s) => ({ ...s, merges: (s.merges ?? []).filter((m) => !hit.some((h) => h.r0 === m.r0 && h.c0 === m.c0)) }));
    if (r0 === r1 && c0 === c1) return;
    let lose = false;
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      const cell = sheet.cells[key(r, c)];
      if ((r !== r0 || c !== c0) && cell && (cell.f || (cell.v != null && cell.v !== ''))) lose = true;
    }
    if (lose && !confirm('Gộp ô chỉ giữ dữ liệu của ô trên cùng bên trái. Tiếp tục?')) return;
    updateSheet((s) => {
      const cells = { ...s.cells };
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
        if (r === r0 && c === c0) continue;
        const k = key(r, c);
        if (cells[k]?.s) cells[k] = { s: cells[k].s }; else delete cells[k];
      }
      const master = cells[key(r0, c0)] ?? {};
      cells[key(r0, c0)] = { ...master, s: { ...master.s, align: master.s?.align ?? 'center' } };
      return { ...s, cells, merges: [...(s.merges ?? []), { r0, c0, r1, c1 }] };
    });
    setSel({ r: r0, c: c0, r2: r1, c2: c1 });
  };

  const fullBorder = (st?: CellStyle) => !!(st?.border?.t && st.border.r && st.border.b && st.border.l);
  const toggleBorder = () => {
    const on = !fullBorder(activeCell?.s);
    applyStyle({ border: on ? { t: true, r: true, b: true, l: true } : undefined });
  };

  const toggleFreeze = () => {
    if (!sheet) return;
    if (sheet.freeze) return updateSheet((s) => ({ ...s, freeze: undefined }));
    const rows = sel.r, cols = sel.c;
    updateSheet((s) => ({ ...s, freeze: rows || cols ? { rows, cols } : { rows: 1, cols: 0 } }));
  };

  // ---- Lọc dữ liệu ----
  const toggleFilter = () => {
    const { r0, r1, c0, c1 } = norm(sel);
    if (!sheet) return;
    if (sheet.filter) return updateSheet((s) => ({ ...s, filter: undefined }));
    const u = usedRange(sheet);
    let f: Filter;
    if (r0 !== r1 || c0 !== c1) f = { r: r0, r1: Math.min(r1, Math.max(u.maxR, r0 + 1)), c0, c1: Math.min(c1, Math.max(u.maxC, c0)), hide: {} };
    else {
      // Một ô: lấy vùng dữ liệu liền kề quanh ô, hàng đầu là tiêu đề.
      let a = c0, b = c0;
      const has = (r: number, c: number) => { const x = sheet.cells[key(r, c)]; return !!x && (x.f || (x.v !== null && x.v !== undefined && x.v !== '')); };
      while (a > 0 && has(r0, a - 1)) a--;
      while (has(r0, b + 1)) b++;
      let end = r0;
      const rowHas = (r: number) => { for (let c = a; c <= b; c++) if (has(r, c)) return true; return false; };
      while (end < u.maxR && rowHas(end + 1)) end++;
      f = { r: r0, c0: a, c1: b, r1: Math.max(end, r0 + 1), hide: {} };
    }
    updateSheet((s) => ({ ...s, filter: f }));
  };

  // ---- Định dạng có điều kiện ----
  const openCf = () => setCfDlg({ kind: 'gt', a: '', b: '', fill: '#FFC7CE', color: '#9C0006', c1: '#F8696B', c2: '#FFEB84', c3: '#63BE7B', three: true });
  const applyCf = () => {
    const { r0, r1, c0, c1 } = norm(sel);
    if (!cfDlg || !sheet) return;
    const u = usedRange(sheet);
    const range = { r0, c0, r1: Math.min(r1, Math.max(u.maxR, r0)), c1: Math.min(c1, Math.max(u.maxC, c0)) };
    const d = cfDlg;
    const rule: CfRule = d.kind === 'scale' ? { range, kind: 'scale', colors: d.three ? [d.c1, d.c2, d.c3] : [d.c1, d.c3] }
      : d.kind === 'bar' ? { range, kind: 'bar', colors: [d.fill === '#FFC7CE' ? '#638EC6' : d.fill] }
      : { range, kind: d.kind, a: d.a.trim(), b: d.kind === 'between' ? d.b.trim() : undefined, fill: d.fill || undefined, color: d.color || undefined };
    updateSheet((s) => ({ ...s, cf: [...(s.cf ?? []), rule] }));
    setCfDlg(null);
  };
  const selN = norm(sel);
  const cfInSel = (sheet?.cf ?? []).filter((x) => !(x.range.r1 < selN.r0 || x.range.r0 > selN.r1 || x.range.c1 < selN.c0 || x.range.c0 > selN.c1));
  const describeCf = (x: CfRule) => {
    const label = CF_KINDS.find(([k]) => k === x.kind)?.[1] ?? 'Quy tắc khác';
    const v = x.kind === 'between' ? ` ${x.a} – ${x.b}` : x.a ? ` ${x.a}` : '';
    return `${addr(x.range.r0, x.range.c0)}:${addr(x.range.r1, x.range.c1)} · ${label}${v}`;
  };

  // ---- Tìm & thay thế ----
  const matches = useMemo(() => (find?.q && sheet && engine ? findMatches(sheet, engine, active, find.q) : []), [find?.q, sheet, engine, active]);
  const gotoMatch = (i: number) => {
    if (!matches.length || !find) return;
    const n = ((i % matches.length) + matches.length) % matches.length;
    setFind({ ...find, idx: n });
    setSel(selectCell(matches[n][0], matches[n][1]));
  };
  const replaceOne = () => {
    if (!find || !sheet || !matches.length) return;
    const [r, c] = matches[Math.min(find.idx, matches.length - 1)];
    const next = replaceInCell(sheet.cells[key(r, c)], find.q, find.repl);
    if (next) setCells([[r, c, next]]);
    else gotoMatch(find.idx + 1);
  };
  const replaceAll = () => {
    if (!find || !sheet) return;
    const out: [number, number, Cell | undefined][] = [];
    for (const [r, c] of matches) {
      const next = replaceInCell(sheet.cells[key(r, c)], find.q, find.repl);
      if (next) out.push([r, c, next]);
    }
    if (out.length) setCells(out);
    setNotice(out.length ? `Đã thay ${out.length} ô.` : 'Không có ô nào để thay (ô công thức chỉ thay khi chữ nằm trong công thức).');
  };

  // ---- Biểu đồ ----
  const updateChart = (id: string, fn: (c: Chart) => Chart | null) =>
    updateSheet((s) => ({ ...s, charts: (s.charts ?? []).flatMap((c) => { if (c.id !== id) return [c]; const n = fn(c); return n ? [n] : []; }) }));
  const insertChart = (kind: ChartKind) => {
    if (!sheet || !engine) return;
    let m = norm(sel);
    if (m.r0 === m.r1 && m.c0 === m.c1) {
      // Một ô: lấy vùng dữ liệu liền kề quanh ô như Excel.
      const has = (r: number, c: number) => { const v = engine.valueAt(active, r, c); return v !== null && v !== ''; };
      if (!has(m.r0, m.c0)) { setNotice('Hãy chọn vùng dữ liệu (kèm hàng tiêu đề) trước khi chèn biểu đồ.'); return; }
      let { r0, r1, c0, c1 } = m;
      for (let grew = true; grew;) {
        grew = false;
        const rowHas = (r: number) => { for (let c = Math.max(0, c0 - 1); c <= c1 + 1; c++) if (has(r, c)) return true; return false; };
        const colHas = (c: number) => { for (let r = Math.max(0, r0 - 1); r <= r1 + 1; r++) if (has(r, c)) return true; return false; };
        if (r0 > 0 && rowHas(r0 - 1)) { r0--; grew = true; }
        if (rowHas(r1 + 1)) { r1++; grew = true; }
        if (c0 > 0 && colHas(c0 - 1)) { c0--; grew = true; }
        if (colHas(c1 + 1)) { c1++; grew = true; }
      }
      m = { r0, r1, c0, c1 };
    }
    const ch = chartFromRange(sheet, engine, active, m, kind, uid());
    if (!ch) { setNotice('Vùng chọn không có số liệu để vẽ biểu đồ.'); return; }
    // Màn hình rộng: đặt bên phải vùng dữ liệu (khoảng 480 × 288). Điện thoại: đặt ngay dưới, vừa bề ngang màn hình.
    const gw = gridRef.current?.clientWidth ?? 800;
    const narrow = gw < 600;
    const c0 = narrow ? m.c0 : m.c1 + 1, r0 = narrow ? m.r1 + 2 : m.r0;
    const c1 = Math.max(c0 + 1, nearestIdx(colX, colX[c0] + (narrow ? Math.max(200, gw - 24) : 480)));
    const r1 = Math.max(r0 + 1, nearestIdx(rowY, rowY[r0] + (narrow ? 240 : 288)));
    ch.anchor = { r0, c0, r1, c1 };
    requestAnimationFrame(() => {
      const el = gridRef.current;
      if (!el) return;
      const right = colX[c1] ?? 0, bottom = rowY[r1] ?? 0;
      if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = Math.min(colX[c0], right - el.clientWidth);
      if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = Math.min(rowY[r0], bottom - el.clientHeight);
    });
    updateSheet((s) => ({ ...s, charts: [...(s.charts ?? []), ch], rowCount: Math.max(s.rowCount, r1 + 20), colCount: Math.max(s.colCount, c1 + 5) }));
    setSelChart(ch.id);
    setNotice(null);
  };
  const activeChart = sheet?.charts?.find((c) => c.id === selChart);

  const save = async (saveAs = false) => {
    if (!book || !engine) return;
    setBusy('Đang lưu…');
    try {
      const out = await saveXlsx(book, loaded.current, (si, r, c) => engine.evaluate(si, r, c));
      const finalName = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
      const p = await saveFile({ name: finalName, kind: 'xlsx', path }, out, saveAs || ((isNew || isCsv) && !path));
      if (p) { setPath(p); setName(p.split(/[\\/]/).pop() ?? finalName); }
      setDirty(false);
    } catch (e) {
      setError(`Lưu thất bại: ${(e as Error).message}`);
    } finally {
      setBusy(null);
    }
  };

  // ---- Lưới ô ----
  const gridRef = useRef<HTMLDivElement>(null);
  const [scroll, setScroll] = useState({ x: 0, y: 0, w: 800, h: 600 });
  useLayoutEffect(() => {
    const el = gridRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setScroll((s) => ({ ...s, w: el.clientWidth, h: el.clientHeight })));
    ro.observe(el);
    return () => ro.disconnect();
  }, [book !== null]);

  const colW = (c: number) => sheet?.colWidths[c] ?? DEFAULT_W;
  const colX = useMemo(() => {
    const xs = [0];
    for (let c = 0; c < (sheet?.colCount ?? 0); c++) xs.push(xs[c] + (sheet?.colWidths[c] ?? DEFAULT_W));
    return xs;
  }, [sheet?.colWidths, sheet?.colCount]);
  /** Hàng ẩn: ẩn thủ công và ẩn do lọc. */
  const hidden = useMemo(() => {
    if (!sheet || !engine) return {} as Record<number, boolean>;
    return { ...(sheet.hiddenRows ?? {}), ...(sheet.filter ? filterHidden(sheet, engine, active, sheet.filter) : {}) };
  }, [sheet, engine, active]);
  hiddenRef.current = hidden;
  const cfLook = useMemo(() => (sheet?.cf?.length && engine ? cfEvaluator(sheet, engine, active) : null), [sheet, engine, active]);
  const rowY = useMemo(() => {
    const ys = [0];
    for (let r = 0; r < (sheet?.rowCount ?? 0); r++) ys.push(ys[r] + (hidden[r] ? 0 : sheet?.rowHeights?.[r] ?? ROW_H));
    return ys;
  }, [sheet?.rowHeights, sheet?.rowCount, hidden]);
  const rowH = (r: number) => (hidden[r] ? 0 : sheet?.rowHeights?.[r] ?? ROW_H);
  const FR = Math.min(sheet?.freeze?.rows ?? 0, (sheet?.rowCount ?? 1) - 1);
  const FC = Math.min(sheet?.freeze?.cols ?? 0, (sheet?.colCount ?? 1) - 1);
  const frozenH = rowY[FR] ?? 0, frozenW = colX[FC] ?? 0;
  /** Ô gộp: ô chính → vùng; ô bị che → true. */
  const mergeInfo = useMemo(() => {
    const masters = new Map<string, Merge>();
    const covered = new Set<string>();
    for (const m of sheet?.merges ?? []) {
      masters.set(key(m.r0, m.c0), m);
      for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) covered.add(key(r, c));
    }
    return { masters, covered };
  }, [sheet?.merges]);

  // Thông báo về tệp CSV chỉ hiện tới khi người dùng bắt đầu thao tác.
  useEffect(() => { setNotice((n) => (n?.startsWith('Tệp CSV') ? null : n)); }, [sel]);

  // Giữ ô đang chọn trong khung nhìn.
  useEffect(() => {
    const el = gridRef.current;
    if (!el || !sheet) return;
    const r = sel.r2, c = sel.c2;
    const top = rowY[r] ?? 0, bottom = rowY[r + 1] ?? top + ROW_H, left = colX[c] ?? 0, right = colX[c + 1] ?? left + DEFAULT_W;
    // Phần cố định che mất frozenH/frozenW phía trên/trái của khung cuộn.
    if (r >= FR) {
      if (top - frozenH < el.scrollTop) el.scrollTop = top - frozenH;
      else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
    }
    if (c >= FC) {
      if (left - frozenW < el.scrollLeft) el.scrollLeft = left - frozenW;
      else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth;
    }
    if (sheet && (r >= sheet.rowCount - 5 || c >= sheet.colCount - 2)) {
      setBook((b) => b && ({ sheets: b.sheets.map((s, i) => i === active ? { ...s, rowCount: Math.max(s.rowCount, r + 50), colCount: Math.max(s.colCount, c + 10) } : s) }));
    }
  }, [sel.r2, sel.c2]);

  const firstRow = indexAt(rowY, scroll.y);
  const lastRow = Math.min((sheet?.rowCount ?? 0) - 1, indexAt(rowY, scroll.y + scroll.h) + 1);
  const firstCol = indexAt(colX, scroll.x);
  const lastCol = Math.min((sheet?.colCount ?? 0) - 1, indexAt(colX, scroll.x + scroll.w) + 1);

  const dragging = useRef(false);
  const resizing = useRef<{ c: number; x: number; w: number } | null>(null);
  const resizingRow = useRef<{ r: number; y: number; h: number } | null>(null);

  const cellFromEvent = (e: React.PointerEvent) => {
    const el = gridRef.current!;
    const rect = el.getBoundingClientRect();
    const vx = e.clientX - rect.left, vy = e.clientY - rect.top;
    const x = FC && vx < frozenW ? vx : vx + el.scrollLeft;
    const y = FR && vy < frozenH ? vy : vy + el.scrollTop;
    return { r: indexAt(rowY, Math.max(0, y)), c: indexAt(colX, Math.max(0, x)) };
  };

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
      if (mod && (e.key.toLowerCase() === 'f' || e.key.toLowerCase() === 'h')) {
        e.preventDefault();
        setFind((f) => ({ q: f?.q ?? '', repl: f?.repl ?? '', idx: 0, replace: e.key.toLowerCase() === 'h' || !!f?.replace }));
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  const onGridKey = (e: React.KeyboardEvent) => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (selChart && (k === 'Delete' || k === 'Backspace')) { e.preventDefault(); updateChart(selChart, () => null); setSelChart(null); return; }
    if (selChart && k === 'Escape') { setSelChart(null); return; }
    if (mod && k.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (mod && k.toLowerCase() === 'y') { e.preventDefault(); redo(); return; }
    if (mod && k.toLowerCase() === 'b') { e.preventDefault(); toggle('bold'); return; }
    if (mod && k.toLowerCase() === 'i') { e.preventDefault(); toggle('italic'); return; }
    if (mod && k.toLowerCase() === 'u') { e.preventDefault(); toggle('underline'); return; }
    if (mod && k.toLowerCase() === 'd') { e.preventDefault(); fill('down'); return; }
    if (mod && k.toLowerCase() === 'r') { e.preventDefault(); fill('right'); return; }
    if (mod && k.toLowerCase() === 'a') { e.preventDefault(); setSel({ r: 0, c: 0, r2: sheet!.rowCount - 1, c2: sheet!.colCount - 1 }); return; }
    if (mod) return;
    const moves: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[k]) { e.preventDefault(); moveSel(...moves[k], e.shiftKey); return; }
    if (k === 'Enter') { e.preventDefault(); moveSel(e.shiftKey ? -1 : 1, 0); return; }
    if (k === 'Tab') { e.preventDefault(); moveSel(0, e.shiftKey ? -1 : 1); return; }
    if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); clearRange(); return; }
    if (k === 'F2') { e.preventDefault(); setEditing({ text: rawText(activeCell), inBar: false }); return; }
    if (k.length === 1) { e.preventDefault(); setEditing({ text: k, inBar: false }); }
  };

  const onEditKey = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); commitEdit([e.shiftKey ? -1 : 1, 0]); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit([0, e.shiftKey ? -1 : 1]); }
    else if (e.key === 'Escape') { e.preventDefault(); setEditing(null); gridRef.current?.focus(); }
  };

  // Thống kê vùng chọn.
  const stats = useMemo(() => {
    if (!engine || !sheet) return '';
    const { r0, r1, c0, c1 } = norm(sel);
    if (r0 === r1 && c0 === c1) return '';
    let sum = 0, count = 0, nums = 0;
    const u = usedRange(sheet);
    for (let r = r0; r <= Math.min(r1, u.maxR); r++)
      for (let c = c0; c <= Math.min(c1, u.maxC); c++) {
        const v = engine.valueAt(active, r, c);
        if (v === null || v === '') continue;
        count++;
        if (typeof v === 'number') { sum += v; nums++; }
      }
    if (!count) return '';
    return `Tổng: ${formatValue(sum, '#,##0.##')} · Trung bình: ${nums ? formatValue(sum / nums, '#,##0.##') : '-'} · Đếm: ${count}`;
  }, [engine, sel, active]);

  if (!book || !sheet || !engine) {
    return (
      <EditorShell kind="xlsx" title={name} dirty={false} onClose={onClose} onSave={() => {}} onSaveAs={() => {}}
        busy={busy} error={error} onDismissError={() => setError(null)} toolbar={null}>
        <div className="loading">{error ? '' : 'Đang mở bảng tính…'}</div>
      </EditorShell>
    );
  }

  const { r0, r1, c0, c1 } = norm(sel);
  const s = activeCell?.s ?? {};
  const barText = editing ? editing.text : rawText(activeCell);
  const activeMerge = mergeInfo.masters.get(key(sel.r, sel.c));
  const activeW = activeMerge ? colX[activeMerge.c1 + 1] - colX[activeMerge.c0] : colW(sel.c);
  const activeH = activeMerge ? rowY[activeMerge.r1 + 1] - rowY[activeMerge.r0] : rowH(sel.r);
  const range = (a: number, b: number) => (b < a ? [] : Array.from({ length: b - a + 1 }, (_, i) => a + i));

  /** Vẽ một ô (ô chính của vùng gộp thì phủ cả vùng). `opaque`: nền trắng để che phần cuộn bên dưới vùng cố định. */
  const renderCell = (r: number, c: number, opaque = false) => {
    const k = key(r, c);
    const m = mergeInfo.masters.get(k);
    const cell = sheet.cells[k];
    const value = cell ? (cell.f ? engine.evaluate(active, r, c) : cell.v) : null;
    const st = cell?.s;
    const isNum = typeof value === 'number';
    const isErr = !!value && typeof value === 'object';
    const b = st?.border;
    const shadows = [b?.t && 'inset 0 1px 0 #000', b?.l && 'inset 1px 0 0 #000'].filter(Boolean).join(',');
    const inSel = opaque && r >= r0 && r <= r1 && c >= c0 && c <= c1;
    const look = cfLook?.(r, c);
    const flt = sheet.filter;
    const fbtn = flt && r === flt.r && c >= flt.c0 && c <= flt.c1;
    const bar = look?.bar;
    return (
      <div key={k} className={`xl-cell${isErr ? ' err' : ''}${st?.wrap ? ' wrap' : ''}${m ? ' merged' : ''}${inSel ? ' insel' : ''}`} style={{
        left: colX[c], top: rowY[r],
        width: m ? colX[m.c1 + 1] - colX[m.c0] : colW(c), height: m ? rowY[m.r1 + 1] - rowY[m.r0] : rowH(r),
        lineHeight: m || st?.wrap ? undefined : `${rowH(r) - 1}px`,
        fontWeight: st?.bold ? 700 : undefined, fontStyle: st?.italic ? 'italic' : undefined,
        textDecoration: st?.underline ? 'underline' : undefined, color: look?.color ?? st?.color,
        backgroundColor: look?.fill ?? st?.fill ?? (opaque || m ? '#fff' : undefined),
        ...(bar ? {
          backgroundImage: `linear-gradient(90deg, ${bar.color}, ${bar.color}55)`,
          backgroundSize: `${Math.round(bar.pct * 100)}% 70%`, backgroundRepeat: 'no-repeat', backgroundPosition: 'left center',
        } : {}),
        paddingRight: fbtn ? 20 : undefined,
        borderRightColor: b?.r ? '#000' : undefined, borderBottomColor: b?.b ? '#000' : undefined,
        boxShadow: shadows || undefined,
        textAlign: st?.align ?? (isNum ? 'right' : typeof value === 'boolean' || isErr ? 'center' : 'left'),
        justifyContent: st?.align === 'center' ? 'center' : st?.align === 'right' || (!st?.align && isNum) ? 'flex-end' : undefined,
      }}>
        {m || st?.wrap ? <span>{formatValue(value, st?.numFmt)}</span> : formatValue(value, st?.numFmt)}
        {fbtn && (
          <button type="button" className={`xl-fbtn${flt.hide[c]?.length ? ' on' : ''}`} title="Lọc"
            onPointerDown={(e) => e.stopPropagation()}
            onClick={(e) => { e.stopPropagation(); setFpop({ c, q: '', hide: new Set(flt.hide[c] ?? []) }); }}>
            {flt.hide[c]?.length ? '⧩' : '▾'}
          </button>
        )}
      </div>
    );
  };

  /** Các ô của một khối hàng × cột, bỏ ô bị gộp che; ô gộp vẽ một lần qua ô chính. */
  const renderBlock = (rows: number[], cols: number[], opaque = false) => {
    const out: JSX.Element[] = [];
    for (const r of rows) for (const c of cols) if (!mergeInfo.covered.has(key(r, c))) out.push(renderCell(r, c, opaque));
    const ra = rows[0], rb = rows[rows.length - 1], ca = cols[0], cb = cols[cols.length - 1];
    if (ra === undefined || ca === undefined) return out;
    // Vùng gộp được vẽ ở khối (cố định hàng?, cố định cột?) chứa ô chính, kể cả khi ô chính đã cuộn khỏi khung nhìn.
    const frozenRows = ra < FR, frozenCols = ca < FC;
    for (const m of sheet.merges ?? []) {
      if (m.r1 < ra || m.r0 > rb || m.c1 < ca || m.c0 > cb) continue;
      if ((m.r0 < FR) === frozenRows && (m.c0 < FC) === frozenCols) out.push(renderCell(m.r0, m.c0, opaque));
    }
    return out;
  };

  /** Khung vùng chọn và ô hiện hành, cắt theo khối (để vẽ lại trong vùng cố định). */
  const renderSel = (ra: number, rb: number, ca: number, cb: number) => {
    const sr0 = Math.max(r0, ra), sr1 = Math.min(r1, rb), sc0 = Math.max(c0, ca), sc1 = Math.min(c1, cb);
    const showActive = sel.r >= ra && sel.r <= rb && sel.c >= ca && sel.c <= cb;
    return (
      <>
        {sr0 <= sr1 && sc0 <= sc1 && (
          <div className="xl-sel" style={{ left: colX[sc0], top: rowY[sr0], width: colX[sc1 + 1] - colX[sc0], height: rowY[sr1 + 1] - rowY[sr0] }} />
        )}
        {showActive && <div className="xl-active" style={{ left: colX[sel.c], top: rowY[sel.r], width: activeW, height: activeH }} />}
      </>
    );
  };

  const visRows = range(Math.max(firstRow, FR), lastRow).filter((r) => !hidden[r]), visCols = range(Math.max(firstCol, FC), lastCol);
  const frRows = range(0, FR - 1).filter((r) => !hidden[r]), frCols = range(0, FC - 1);
  const inFrozen = (r: number, c: number) => r < FR || c < FC;

  return (
    <EditorShell
      kind="xlsx"
      title={name}
      dirty={dirty}
      onClose={onClose}
      onSave={() => save()}
      onSaveAs={() => save(true)}
      busy={busy}
      error={error}
      onDismissError={() => setError(null)}
      toolbar={
        <>
          <ToolButton icon="↶" label="Hoàn tác" onClick={undo} disabled={!history.length} />
          <ToolButton icon="↷" label="Làm lại" onClick={redo} disabled={!future.length} />
          <Sep />
          <ToolButton icon={<b>B</b>} label="Đậm" active={!!s.bold} onClick={() => toggle('bold')} />
          <ToolButton icon={<i>I</i>} label="Nghiêng" active={!!s.italic} onClick={() => toggle('italic')} />
          <ToolButton icon={<u>U</u>} label="Gạch chân" active={!!s.underline} onClick={() => toggle('underline')} />
          <label className="tb-field" title="Màu chữ"><span className="tb-mini">A</span>
            <input type="color" value={s.color ?? '#000000'} onChange={(e) => applyStyle({ color: e.target.value })} /></label>
          <label className="tb-field" title="Màu nền ô"><span className="tb-mini">▮</span>
            <input type="color" value={s.fill ?? '#ffffff'} onChange={(e) => applyStyle({ fill: e.target.value })} /></label>
          <ToolButton icon="∅" label="Bỏ màu nền" onClick={() => applyStyle({ fill: undefined })} />
          <ToolButton icon="▦" label={fullBorder(s) ? 'Bỏ kẻ khung' : 'Kẻ khung'} active={fullBorder(s)} onClick={toggleBorder} />
          <Sep />
          <ToolButton icon="⇤" label="Căn trái" active={s.align === 'left'} onClick={() => applyStyle({ align: 'left' })} />
          <ToolButton icon="↔" label="Căn giữa" active={s.align === 'center'} onClick={() => applyStyle({ align: 'center' })} />
          <ToolButton icon="⇥" label="Căn phải" active={s.align === 'right'} onClick={() => applyStyle({ align: 'right' })} />
          <ToolButton icon="↵" label="Xuống dòng" active={!!s.wrap} onClick={() => applyStyle({ wrap: !s.wrap || undefined })} />
          <ToolButton icon="⊞" label={activeMerge ? 'Bỏ gộp ô' : 'Gộp ô'} active={!!activeMerge} onClick={toggleMerge} />
          <label className="tb-field" title="Định dạng số">
            <select value={s.numFmt ?? 'General'} onChange={(e) => applyStyle({ numFmt: e.target.value === 'General' ? undefined : e.target.value })}>
              {NUM_FORMATS.map(([f, l]) => <option key={f} value={f}>{l}</option>)}
              {s.numFmt && !NUM_FORMATS.some(([f]) => f === s.numFmt) && <option value={s.numFmt}>{s.numFmt}</option>}
            </select>
          </label>
          <Sep />
          <ToolButton icon="Σ" label="Tự động tính tổng" onClick={autoSum} />
          <ToolButton icon="A↓" label="Sắp xếp tăng" onClick={() => sort(true)} />
          <ToolButton icon="Z↓" label="Sắp xếp giảm" onClick={() => sort(false)} />
          <ToolButton icon="🔍" label="Tìm & thay" active={!!find} onClick={() => setFind(find ? null : { q: '', repl: '', idx: 0, replace: true })} />
          <ToolButton icon="❄" label={sheet.freeze ? 'Bỏ cố định' : 'Cố định'} active={!!sheet.freeze} onClick={toggleFreeze} />
          <ToolButton icon="⧩" label={sheet.filter ? 'Bỏ lọc' : 'Lọc dữ liệu'} active={!!sheet.filter} onClick={toggleFilter} />
          <ToolButton icon="◧" label="Định dạng có điều kiện" active={cfInSel.length > 0} onClick={openCf} />
          <ToolButton icon="📊" label="Chèn biểu đồ" onClick={() => insertChart('col')} />
          {activeChart && (
            <>
              <label className="tb-field" title="Loại biểu đồ">
                <select value={activeChart.kind} onChange={(e) => {
                  const kind = e.target.value as ChartKind;
                  updateChart(activeChart.id, (c) => ({ ...c, kind }));
                }}>
                  {CHART_KINDS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                </select>
              </label>
              <ToolButton icon="T" label="Tiêu đề biểu đồ" onClick={() => {
                const t = prompt('Tiêu đề biểu đồ (để trống để bỏ)', activeChart.title ?? '');
                if (t !== null) updateChart(activeChart.id, (c) => ({ ...c, title: t.trim() || undefined }));
              }} />
              <ToolButton icon="🗑" label="Xoá biểu đồ" onClick={() => { updateChart(activeChart.id, () => null); setSelChart(null); }} />
            </>
          )}
          <Sep />
          <ToolButton icon="+≡" label="Chèn hàng" onClick={() => shift('row', 1)} />
          <ToolButton icon="−≡" label="Xoá hàng" onClick={() => shift('row', -1)} />
          <ToolButton icon="+‖" label="Chèn cột" onClick={() => shift('col', 1)} />
          <ToolButton icon="−‖" label="Xoá cột" onClick={() => shift('col', -1)} />
        </>
      }
      status={notice || stats || `${sheet.name} · ${addr(sel.r, sel.c)}${sheet.freeze ? ` · Cố định ${FR} hàng, ${FC} cột` : ''}`}
    >
      <div className="xl">
        {find && (
          <form className="xl-find" onSubmit={(e) => { e.preventDefault(); gotoMatch(find.idx + (matches.length && sel.r === matches[find.idx]?.[0] && sel.c === matches[find.idx]?.[1] ? 1 : 0)); }}>
            <input autoFocus placeholder="Tìm trong trang tính…" value={find.q}
              onChange={(e) => { setFind({ ...find, q: e.target.value, idx: 0 }); setNotice(null); }}
              onKeyDown={(e) => {
                if (e.key === 'Escape') { setFind(null); gridRef.current?.focus(); }
                if (e.key === 'Enter') { e.preventDefault(); e.currentTarget.form?.requestSubmit(); }
              }} />
            <span className="xl-find-count">{find.q ? (matches.length ? `${Math.min(find.idx, matches.length - 1) + 1}/${matches.length}` : 'Không thấy') : ''}</span>
            <button type="button" className="btn" onClick={() => gotoMatch(find.idx - 1)} disabled={!matches.length} title="Kết quả trước">‹</button>
            <button type="button" className="btn" onClick={() => gotoMatch(find.idx + 1)} disabled={!matches.length} title="Kết quả sau">›</button>
            {find.replace && (
              <>
                <input placeholder="Thay bằng…" value={find.repl} onChange={(e) => setFind({ ...find, repl: e.target.value })} />
                <button type="button" className="btn" onClick={replaceOne} disabled={!matches.length}>Thay</button>
                <button type="button" className="btn" onClick={replaceAll} disabled={!matches.length}>Thay tất cả</button>
              </>
            )}
            <button type="button" className="icon-btn" onClick={() => { setFind(null); setNotice(null); gridRef.current?.focus(); }} title="Đóng">✕</button>
          </form>
        )}
        <div className="xl-bar">
          <span className="xl-addr">{r0 === r1 && c0 === c1 || activeMerge ? addr(sel.r, sel.c) : `${addr(r0, c0)}:${addr(r1, c1)}`}</span>
          <span className="xl-fx">fx</span>
          <input
            className="xl-formula"
            value={barText}
            onFocus={() => !editing && setEditing({ text: rawText(activeCell), inBar: true })}
            onChange={(e) => setEditing({ text: e.target.value, inBar: true })}
            onKeyDown={onEditKey}
            onBlur={() => editing?.inBar && commitEdit()}
          />
        </div>
        <div className="xl-grid-wrap">
          <div className="xl-corner" onClick={() => setSel({ r: 0, c: 0, r2: sheet.rowCount - 1, c2: sheet.colCount - 1 })} />
          <div className="xl-colhead" style={{ left: HEAD_W, height: HEAD_H }}>
            <div style={{ transform: `translateX(${-scroll.x}px)`, position: 'relative', height: HEAD_H }}>
              {[...frCols, ...visCols].map((c) => (
                <div key={c} className={`xl-hcell${c >= c0 && c <= c1 ? ' on' : ''}${c < FC ? ' frz' : ''}`}
                  style={{ left: colX[c] + (c < FC ? scroll.x : 0), width: colW(c) }}
                  onPointerDown={(e) => { if (!(e.target as HTMLElement).dataset.resize) setSel(e.shiftKey ? { ...sel, c2: c, r: 0, r2: sheet.rowCount - 1 } : { r: 0, c, r2: sheet.rowCount - 1, c2: c }); }}>
                  {colName(c)}
                  <span className="xl-resize" data-resize="1"
                    onPointerDown={(e) => { e.stopPropagation(); (e.target as HTMLElement).setPointerCapture(e.pointerId); resizing.current = { c, x: e.clientX, w: colW(c) }; }}
                    onPointerMove={(e) => {
                      const rz = resizing.current;
                      if (!rz) return;
                      const w = Math.max(24, rz.w + e.clientX - rz.x);
                      setBook((b) => b && ({ sheets: b.sheets.map((sh, i) => i === active ? { ...sh, colWidths: { ...sh.colWidths, [rz.c]: w } } : sh) }));
                    }}
                    onPointerUp={() => { if (resizing.current) setDirty(true); resizing.current = null; }} />
                </div>
              ))}
            </div>
          </div>
          <div className="xl-rowhead" style={{ top: HEAD_H, width: HEAD_W }}>
            <div style={{ transform: `translateY(${-scroll.y}px)`, position: 'relative' }}>
              {[...frRows, ...visRows].map((r) => (
                <div key={r} className={`xl-rcell${r >= r0 && r <= r1 ? ' on' : ''}${r < FR ? ' frz' : ''}`}
                  style={{ top: rowY[r] + (r < FR ? scroll.y : 0), height: rowH(r) }}
                  onPointerDown={(e) => { if (!(e.target as HTMLElement).dataset.resize) setSel(e.shiftKey ? { ...sel, r2: r, c: 0, c2: sheet.colCount - 1 } : { r, c: 0, r2: r, c2: sheet.colCount - 1 }); }}>
                  {r + 1}
                  <span className="xl-rresize" data-resize="1"
                    onPointerDown={(e) => { e.stopPropagation(); (e.target as HTMLElement).setPointerCapture(e.pointerId); resizingRow.current = { r, y: e.clientY, h: rowH(r) }; }}
                    onPointerMove={(e) => {
                      const rz = resizingRow.current;
                      if (!rz) return;
                      const h = Math.max(12, rz.h + e.clientY - rz.y);
                      setBook((b) => b && ({ sheets: b.sheets.map((sh, i) => i === active ? { ...sh, rowHeights: { ...sh.rowHeights, [rz.r]: h } } : sh) }));
                    }}
                    onPointerUp={() => { if (resizingRow.current) setDirty(true); resizingRow.current = null; }} />
                </div>
              ))}
            </div>
          </div>
          <div
            ref={gridRef}
            className="xl-grid"
            tabIndex={0}
            style={{ left: HEAD_W, top: HEAD_H }}
            onScroll={(e) => { const el = e.currentTarget; setScroll({ x: el.scrollLeft, y: el.scrollTop, w: el.clientWidth, h: el.clientHeight }); }}
            onKeyDown={onGridKey}
            onCopy={(e) => { if (editing) return; e.preventDefault(); e.clipboardData.setData('text/plain', rangeTSV()); }}
            onCut={(e) => { if (editing) return; e.preventDefault(); e.clipboardData.setData('text/plain', rangeTSV()); clearRange(); }}
            onPaste={(e) => { if (editing) return; e.preventDefault(); paste(e.clipboardData.getData('text/plain')); }}
            onPointerDown={(e) => {
              if ((e.target as HTMLElement).closest('.xl-editor, .xl-chart')) return;
              if (selChart) setSelChart(null);
              const { r, c } = cellFromEvent(e);
              if (editing) commitEdit();
              if (e.shiftKey) setSel({ ...sel, r2: r, c2: c });
              else if (e.pointerType !== 'mouse' && (r === sel.r && c === sel.c || mergeAt(sheet.merges, r, c) === activeMerge && activeMerge) && (r0 === r1 && c0 === c1 || activeMerge)) {
                setEditing({ text: rawText(activeCell), inBar: false });
                return;
              } else setSel(selectCell(r, c));
              if (e.pointerType === 'mouse') {
                dragging.current = true;
                gridRef.current!.setPointerCapture(e.pointerId);
              }
            }}
            onPointerMove={(e) => {
              if (!dragging.current) return;
              const { r, c } = cellFromEvent(e);
              setSel((s) => (s.r2 === r && s.c2 === c ? s : { ...s, r2: r, c2: c }));
            }}
            onPointerUp={() => (dragging.current = false)}
            onDoubleClick={() => setEditing({ text: rawText(activeCell), inBar: false })}
          >
            <div style={{ width: colX[colX.length - 1], height: rowY[rowY.length - 1], position: 'relative' }}>
              {renderBlock(visRows, visCols)}
              {renderSel(0, sheet.rowCount, 0, sheet.colCount)}
              {FR > 0 && (
                <div className="xl-frz-top" style={{ zIndex: 2 }}>
                  {renderBlock(frRows, visCols, true)}
                  {renderSel(0, FR - 1, FC, sheet.colCount)}
                  <div className="xl-frz-line h" style={{ top: frozenH - 1, left: 0, width: colX[colX.length - 1] }} />
                </div>
              )}
              {FC > 0 && (
                <div className="xl-frz-left" style={{ zIndex: 2 }}>
                  {renderBlock(visRows, frCols, true)}
                  {renderSel(FR, sheet.rowCount, 0, FC - 1)}
                  <div className="xl-frz-line v" style={{ left: frozenW - 1, top: 0, height: rowY[rowY.length - 1] }} />
                </div>
              )}
              {FR > 0 && FC > 0 && (
                <div className="xl-frz-corner" style={{ zIndex: 3 }}>
                  {renderBlock(frRows, frCols, true)}
                  {renderSel(0, FR - 1, 0, FC - 1)}
                </div>
              )}
              {sheet.charts?.map((ch) => {
                const a = ch.anchor;
                if (a.c1 >= colX.length || a.r1 >= rowY.length) return null;
                const off = chartOff?.id === ch.id ? chartOff : null;
                const left = colX[a.c0] + (off?.mode === 'move' ? off.dx : 0), top = rowY[a.r0] + (off?.mode === 'move' ? off.dy : 0);
                const w = Math.max(80, colX[a.c1] - colX[a.c0] + (off?.mode === 'size' ? off.dx : 0));
                const h = Math.max(60, rowY[a.r1] - rowY[a.r0] + (off?.mode === 'size' ? off.dy : 0));
                const on = selChart === ch.id;
                const start = (e: React.PointerEvent, mode: 'move' | 'size') => {
                  e.stopPropagation();
                  if (editing) commitEdit();
                  setSelChart(ch.id);
                  gridRef.current?.focus({ preventScroll: true });
                  if (mode === 'move' && !on && e.pointerType !== 'mouse') return; // chạm lần đầu chỉ chọn, để vẫn cuộn được
                  (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
                  chartDrag.current = { id: ch.id, mode, x: e.clientX, y: e.clientY, moved: false };
                };
                const move = (e: React.PointerEvent) => {
                  const d = chartDrag.current;
                  if (!d || d.id !== ch.id) return;
                  e.stopPropagation();
                  const dx = e.clientX - d.x, dy = e.clientY - d.y;
                  if (Math.abs(dx) + Math.abs(dy) > 3) d.moved = true;
                  setChartOff({ id: ch.id, dx, dy, mode: d.mode });
                };
                const end = (e: React.PointerEvent) => {
                  const d = chartDrag.current;
                  chartDrag.current = null;
                  setChartOff(null);
                  if (!d || !d.moved) return;
                  e.stopPropagation();
                  const dx = e.clientX - d.x, dy = e.clientY - d.y;
                  updateChart(ch.id, (c) => {
                    const an = c.anchor;
                    if (d.mode === 'size') {
                      return { ...c, anchor: { ...an, c1: Math.max(an.c0 + 1, nearestIdx(colX, colX[an.c1] + dx)), r1: Math.max(an.r0 + 1, nearestIdx(rowY, rowY[an.r1] + dy)) } };
                    }
                    const c0 = nearestIdx(colX, Math.max(0, colX[an.c0] + dx)), r0 = nearestIdx(rowY, Math.max(0, rowY[an.r0] + dy));
                    return { ...c, anchor: { r0, c0, r1: r0 + an.r1 - an.r0, c1: c0 + an.c1 - an.c0 } };
                  });
                };
                return (
                  <div key={ch.id} className={`xl-chart${on ? ' on' : ''}`} style={{ left, top, width: w, height: h, touchAction: on ? 'none' : undefined }}
                    onPointerDown={(e) => start(e, 'move')} onPointerMove={move} onPointerUp={end} onPointerCancel={end}
                    onDoubleClick={(e) => e.stopPropagation()}>
                    <ChartView kind={ch.kind} title={ch.title} data={chartData(ch, book, engine, active)} w={w} h={h} />
                    {on && <span className="xl-chart-size" onPointerDown={(e) => start(e, 'size')} onPointerMove={move} onPointerUp={end} onPointerCancel={end} />}
                  </div>
                );
              })}
              {editing && !editing.inBar && (
                <textarea
                  className="xl-editor"
                  autoFocus
                  value={editing.text}
                  rows={1}
                  style={{
                    left: colX[sel.c] + (inFrozen(sel.r, sel.c) && sel.c < FC ? scroll.x : 0),
                    top: rowY[sel.r] + (inFrozen(sel.r, sel.c) && sel.r < FR ? scroll.y : 0),
                    minWidth: activeW, height: Math.max(activeH, ROW_H),
                  }}
                  onFocus={(e) => { const t = e.currentTarget; t.setSelectionRange(t.value.length, t.value.length); }}
                  onChange={(e) => setEditing({ text: e.target.value, inBar: false })}
                  onKeyDown={onEditKey}
                  onBlur={() => commitEdit()}
                />
              )}
            </div>
          </div>
        </div>
        <div className="xl-tabs">
          <button className="xl-tab add" title="Thêm trang tính" onClick={() => {
            const n = book.sheets.length + 1;
            let nm = `Sheet${n}`;
            let k = n;
            while (book.sheets.some((x) => x.name === nm)) nm = `Sheet${++k}`;
            commit((b) => ({ sheets: [...b.sheets, emptySheet(nm, uid())] }));
            setActive(book.sheets.length);
            setSel({ r: 0, c: 0, r2: 0, c2: 0 });
          }}>＋</button>
          {book.sheets.map((sh, i) => (
            <button key={sh.id} className={`xl-tab${i === active ? ' on' : ''}`}
              onClick={() => { setActive(i); setSel({ r: 0, c: 0, r2: 0, c2: 0 }); setEditing(null); }}
              onDoubleClick={() => {
                const nm = prompt('Đổi tên trang tính', sh.name)?.trim();
                if (!nm || nm === sh.name) return;
                if (book.sheets.some((x) => x.name === nm)) return setError('Tên trang tính đã tồn tại.');
                commit((b) => ({ sheets: b.sheets.map((x, j) => j === i ? { ...x, name: nm } : x) }));
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                if (book.sheets.length < 2) return;
                if (!confirm(`Xoá trang tính "${sh.name}"?`)) return;
                commit((b) => ({ sheets: b.sheets.filter((_, j) => j !== i) }));
                setActive(0);
              }}
              title="Bấm đúp để đổi tên, chuột phải để xoá">
              {sh.name}
            </button>
          ))}
        </div>
      </div>
      {fpop && sheet.filter && (() => {
        const f = sheet.filter;
        const all = columnValues(sheet, engine, active, f, fpop.c);
        const q = fpop.q.toLocaleLowerCase('vi');
        const shown = q ? all.filter((v) => v.toLocaleLowerCase('vi').includes(q)) : all;
        const setHide = (hide: Set<string>) => setFpop({ ...fpop, hide });
        const head = formatValue(engine.valueAt(active, f.r, fpop.c), sheet.cells[key(f.r, fpop.c)]?.s?.numFmt) || colName(fpop.c);
        return (
          <div className="modal-backdrop" onClick={() => setFpop(null)}>
            <form className="modal xl-fpop" onClick={(e) => e.stopPropagation()} onSubmit={(e) => {
              e.preventDefault();
              const hide = [...fpop.hide].filter((v) => all.includes(v));
              updateSheet((s) => s.filter ? ({ ...s, filter: { ...s.filter, hide: { ...s.filter.hide, [fpop.c]: hide } } }) : s);
              setFpop(null);
            }}>
              <h3>Lọc cột “{head}”</h3>
              <div className="xl-fpop-sort">
                <button type="button" className="btn" onClick={() => { setFpop(null); sortRange(true, { r0: f.r + 1, r1: f.r1, c0: f.c0, c1: f.c1 }, fpop.c); }}>A→Z</button>
                <button type="button" className="btn" onClick={() => { setFpop(null); sortRange(false, { r0: f.r + 1, r1: f.r1, c0: f.c0, c1: f.c1 }, fpop.c); }}>Z→A</button>
              </div>
              <input placeholder="Tìm giá trị…" value={fpop.q} onChange={(e) => setFpop({ ...fpop, q: e.target.value })} />
              <label className="row"><input type="checkbox" checked={shown.every((v) => !fpop.hide.has(v))}
                onChange={(e) => { const h = new Set(fpop.hide); shown.forEach((v) => (e.target.checked ? h.delete(v) : h.add(v))); setHide(h); }} /> (Chọn tất cả)</label>
              <div className="xl-fpop-list">
                {shown.map((v) => (
                  <label key={v} className="row"><input type="checkbox" checked={!fpop.hide.has(v)}
                    onChange={(e) => { const h = new Set(fpop.hide); if (e.target.checked) h.delete(v); else h.add(v); setHide(h); }} /> {v === '' ? '(Trống)' : v}</label>
                ))}
              </div>
              <div className="modal-actions">
                <button type="button" className="btn" onClick={() => setHide(new Set())}>Bỏ lọc cột</button>
                <span style={{ flex: 1 }} />
                <button type="button" className="btn" onClick={() => setFpop(null)}>Huỷ</button>
                <button type="submit" className="btn primary">Lọc</button>
              </div>
            </form>
          </div>
        );
      })()}
      {cfDlg && (
        <div className="modal-backdrop" onClick={() => setCfDlg(null)}>
          <form className="modal xl-cfdlg" onClick={(e) => e.stopPropagation()} onSubmit={(e) => { e.preventDefault(); applyCf(); }}>
            <h3>Định dạng có điều kiện</h3>
            <p className="muted">Áp dụng cho vùng {addr(r0, c0)}{r0 !== r1 || c0 !== c1 ? `:${addr(r1, c1)}` : ''}</p>
            <label>Quy tắc
              <select value={cfDlg.kind} onChange={(e) => setCfDlg({ ...cfDlg, kind: e.target.value as CfKind })}>
                {CF_KINDS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
              </select>
            </label>
            {cfDlg.kind !== 'scale' && cfDlg.kind !== 'bar' && (
              <div className="xl-cf-row">
                <label>{cfDlg.kind === 'between' ? 'Từ' : cfDlg.kind === 'text' ? 'Chữ cần chứa' : 'Giá trị'}
                  <input name="a" value={cfDlg.a} required onChange={(e) => setCfDlg({ ...cfDlg, a: e.target.value })} /></label>
                {cfDlg.kind === 'between' && <label>Đến<input name="b" value={cfDlg.b} required onChange={(e) => setCfDlg({ ...cfDlg, b: e.target.value })} /></label>}
              </div>
            )}
            {cfDlg.kind === 'scale' ? (
              <div className="xl-cf-row">
                <label>Nhỏ nhất<input type="color" value={cfDlg.c1} onChange={(e) => setCfDlg({ ...cfDlg, c1: e.target.value })} /></label>
                {cfDlg.three && <label>Giữa<input type="color" value={cfDlg.c2} onChange={(e) => setCfDlg({ ...cfDlg, c2: e.target.value })} /></label>}
                <label>Lớn nhất<input type="color" value={cfDlg.c3} onChange={(e) => setCfDlg({ ...cfDlg, c3: e.target.value })} /></label>
                <label className="row"><input type="checkbox" checked={cfDlg.three} onChange={(e) => setCfDlg({ ...cfDlg, three: e.target.checked })} /> 3 màu</label>
              </div>
            ) : cfDlg.kind === 'bar' ? (
              <div className="xl-cf-row">
                <label>Màu thanh<input type="color" value={cfDlg.fill === '#FFC7CE' ? '#638EC6' : cfDlg.fill} onChange={(e) => setCfDlg({ ...cfDlg, fill: e.target.value })} /></label>
              </div>
            ) : (
              <div className="xl-cf-row">
                <label>Màu nền<input type="color" value={cfDlg.fill} onChange={(e) => setCfDlg({ ...cfDlg, fill: e.target.value })} /></label>
                <label>Màu chữ<input type="color" value={cfDlg.color} onChange={(e) => setCfDlg({ ...cfDlg, color: e.target.value })} /></label>
                <span className="xl-cf-sample" style={{ background: cfDlg.fill, color: cfDlg.color }}>Mẫu 123</span>
              </div>
            )}
            {cfInSel.length > 0 && (
              <div className="xl-cf-list">
                <b>Quy tắc trong vùng chọn</b>
                {cfInSel.map((x, i) => (
                  <div key={i} className="xl-cf-item">
                    <span className="xl-cf-sample" style={{ background: x.kind === 'scale' ? `linear-gradient(90deg, ${(x.colors ?? []).join(',')})` : x.kind === 'bar' ? x.colors?.[0] : x.fill, color: x.color }}>Aa</span>
                    <span>{describeCf(x)}</span>
                    <button type="button" className="icon-btn" title="Xoá quy tắc" onClick={() => updateSheet((s) => ({ ...s, cf: (s.cf ?? []).filter((y) => y !== x) }))}>✕</button>
                  </div>
                ))}
              </div>
            )}
            <div className="modal-actions">
              <span style={{ flex: 1 }} />
              <button type="button" className="btn" onClick={() => setCfDlg(null)}>Đóng</button>
              <button type="submit" className="btn primary">Áp dụng</button>
            </div>
          </form>
        </div>
      )}
    </EditorShell>
  );
}
