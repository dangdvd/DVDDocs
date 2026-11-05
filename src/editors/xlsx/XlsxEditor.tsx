import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { OpenedFile, saveFile } from '../../lib/fileio';
import { EditorShell, Sep, ToolButton } from '../../components/EditorShell';
import { uid } from '../../lib/uid';
import {
  Book, Cell, CellStyle, Engine, Sheet, addr, colName, emptySheet, formatValue, key, moveFormula,
  parseInput, parseKey, rawText, shiftFormula, usedRange,
} from './sheetModel';
import { LoadedBook, loadXlsx, saveXlsx } from './xlsxIO';

const ROW_H = 24;
const DEFAULT_W = 96;
const HEAD_W = 48;
const HEAD_H = 24;

interface Sel { r: number; c: number; r2: number; c2: number }
const norm = (s: Sel) => ({ r0: Math.min(s.r, s.r2), r1: Math.max(s.r, s.r2), c0: Math.min(s.c, s.c2), c1: Math.max(s.c, s.c2) });

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
  const [path, setPath] = useState(file.path);
  const [name, setName] = useState(file.name);
  const clip = useRef<{ text: string; r: number; c: number; cells: (Cell | undefined)[][] } | null>(null);

  useEffect(() => {
    if (isNew) return;
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

  const moveSel = (dr: number, dc: number, extend = false) => {
    setSel((s) => {
      if (extend) return { ...s, r2: Math.max(0, s.r2 + dr), c2: Math.max(0, s.c2 + dc) };
      const r = Math.max(0, s.r + dr), c = Math.max(0, s.c + dc);
      return { r, c, r2: r, c2: c };
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
        let colWidths = s.colWidths;
        if (i === active && axis === 'col') {
          colWidths = {};
          for (const [ci, w] of Object.entries(s.colWidths)) {
            const n = Number(ci);
            if (n < at) colWidths[n] = w;
            else if (d > 0) colWidths[n + d] = w;
            else if (n >= at - d) colWidths[n + d] = w;
          }
        }
        return { ...s, cells, colWidths, rowCount: s.rowCount + (axis === 'row' && i === active && d > 0 ? d : 0), colCount: s.colCount + (axis === 'col' && i === active && d > 0 ? d : 0) };
      }),
    }));
  };

  const sort = (asc: boolean) => {
    if (!sheet || !engine) return;
    let { r0, r1, c0, c1 } = norm(sel);
    if (r0 === r1) {
      const u = usedRange(sheet);
      r0 = 1; r1 = u.maxR; c0 = 0; c1 = u.maxC;
      if (r1 < 1) return;
    }
    const col = sel.c;
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

  const save = async (saveAs = false) => {
    if (!book || !engine) return;
    setBusy('Đang lưu…');
    try {
      const out = await saveXlsx(book, loaded.current, (si, r, c) => engine.evaluate(si, r, c));
      const finalName = name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
      const p = await saveFile({ name: finalName, kind: 'xlsx', path }, out, saveAs || (isNew && !path));
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

  // Giữ ô đang chọn trong khung nhìn.
  useEffect(() => {
    const el = gridRef.current;
    if (!el || !sheet) return;
    const r = sel.r2, c = sel.c2;
    const top = r * ROW_H, bottom = top + ROW_H, left = colX[c] ?? 0, right = colX[c + 1] ?? left + DEFAULT_W;
    if (top < el.scrollTop) el.scrollTop = top;
    else if (bottom > el.scrollTop + el.clientHeight) el.scrollTop = bottom - el.clientHeight;
    if (left < el.scrollLeft) el.scrollLeft = left;
    else if (right > el.scrollLeft + el.clientWidth) el.scrollLeft = right - el.clientWidth;
    if (sheet && (r >= sheet.rowCount - 5 || c >= sheet.colCount - 2)) {
      setBook((b) => b && ({ sheets: b.sheets.map((s, i) => i === active ? { ...s, rowCount: Math.max(s.rowCount, r + 50), colCount: Math.max(s.colCount, c + 10) } : s) }));
    }
  }, [sel.r2, sel.c2]);

  const firstRow = Math.floor(scroll.y / ROW_H);
  const lastRow = Math.min((sheet?.rowCount ?? 0) - 1, Math.ceil((scroll.y + scroll.h) / ROW_H));
  let firstCol = 0;
  while (firstCol < colX.length - 2 && colX[firstCol + 1] < scroll.x) firstCol++;
  let lastCol = firstCol;
  while (lastCol < colX.length - 2 && colX[lastCol] < scroll.x + scroll.w) lastCol++;

  const dragging = useRef(false);
  const resizing = useRef<{ c: number; x: number; w: number } | null>(null);

  const cellFromEvent = (e: React.PointerEvent) => {
    const rect = gridRef.current!.getBoundingClientRect();
    const x = e.clientX - rect.left + gridRef.current!.scrollLeft;
    const y = e.clientY - rect.top + gridRef.current!.scrollTop;
    let c = 0;
    while (c < colX.length - 2 && colX[c + 1] <= x) c++;
    return { r: Math.max(0, Math.floor(y / ROW_H)), c };
  };

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  const onGridKey = (e: React.KeyboardEvent) => {
    if (editing) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
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
          <Sep />
          <ToolButton icon="⇤" label="Căn trái" active={s.align === 'left'} onClick={() => applyStyle({ align: 'left' })} />
          <ToolButton icon="↔" label="Căn giữa" active={s.align === 'center'} onClick={() => applyStyle({ align: 'center' })} />
          <ToolButton icon="⇥" label="Căn phải" active={s.align === 'right'} onClick={() => applyStyle({ align: 'right' })} />
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
          <Sep />
          <ToolButton icon="+≡" label="Chèn hàng" onClick={() => shift('row', 1)} />
          <ToolButton icon="−≡" label="Xoá hàng" onClick={() => shift('row', -1)} />
          <ToolButton icon="+‖" label="Chèn cột" onClick={() => shift('col', 1)} />
          <ToolButton icon="−‖" label="Xoá cột" onClick={() => shift('col', -1)} />
        </>
      }
      status={stats || `${sheet.name} · ${addr(sel.r, sel.c)}`}
    >
      <div className="xl">
        <div className="xl-bar">
          <span className="xl-addr">{r0 === r1 && c0 === c1 ? addr(sel.r, sel.c) : `${addr(r0, c0)}:${addr(r1, c1)}`}</span>
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
              {Array.from({ length: lastCol - firstCol + 1 }, (_, i) => firstCol + i).map((c) => (
                <div key={c} className={`xl-hcell${c >= c0 && c <= c1 ? ' on' : ''}`} style={{ left: colX[c], width: colW(c) }}
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
              {Array.from({ length: Math.max(0, lastRow - firstRow + 1) }, (_, i) => firstRow + i).map((r) => (
                <div key={r} className={`xl-rcell${r >= r0 && r <= r1 ? ' on' : ''}`} style={{ top: r * ROW_H, height: ROW_H }}
                  onPointerDown={(e) => setSel(e.shiftKey ? { ...sel, r2: r, c: 0, c2: sheet.colCount - 1 } : { r, c: 0, r2: r, c2: sheet.colCount - 1 })}>
                  {r + 1}
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
              if ((e.target as HTMLElement).closest('.xl-editor')) return;
              const { r, c } = cellFromEvent(e);
              if (editing) commitEdit();
              if (e.shiftKey) setSel({ ...sel, r2: r, c2: c });
              else if (e.pointerType !== 'mouse' && r === sel.r && c === sel.c && r0 === r1 && c0 === c1) {
                setEditing({ text: rawText(activeCell), inBar: false });
                return;
              } else setSel({ r, c, r2: r, c2: c });
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
            <div style={{ width: colX[colX.length - 1], height: sheet.rowCount * ROW_H, position: 'relative' }}>
              {Array.from({ length: Math.max(0, lastRow - firstRow + 1) }, (_, i) => firstRow + i).flatMap((r) =>
                Array.from({ length: lastCol - firstCol + 1 }, (_, j) => firstCol + j).map((c) => {
                  const cell = sheet.cells[key(r, c)];
                  const value = cell ? (cell.f ? engine.evaluate(active, r, c) : cell.v) : null;
                  const st = cell?.s;
                  const isNum = typeof value === 'number';
                  const isErr = !!value && typeof value === 'object';
                  return (
                    <div key={`${r},${c}`} className={`xl-cell${isErr ? ' err' : ''}`} style={{
                      left: colX[c], top: r * ROW_H, width: colW(c), height: ROW_H,
                      fontWeight: st?.bold ? 700 : undefined, fontStyle: st?.italic ? 'italic' : undefined,
                      textDecoration: st?.underline ? 'underline' : undefined, color: st?.color, background: st?.fill,
                      textAlign: st?.align ?? (isNum ? 'right' : typeof value === 'boolean' || isErr ? 'center' : 'left'),
                    }}>
                      {formatValue(value, st?.numFmt)}
                    </div>
                  );
                }),
              )}
              <div className="xl-sel" style={{ left: colX[c0], top: r0 * ROW_H, width: colX[c1 + 1] - colX[c0], height: (r1 - r0 + 1) * ROW_H }} />
              <div className="xl-active" style={{ left: colX[sel.c], top: sel.r * ROW_H, width: colW(sel.c), height: ROW_H }} />
              {editing && !editing.inBar && (
                <input
                  className="xl-editor"
                  autoFocus
                  value={editing.text}
                  style={{ left: colX[sel.c], top: sel.r * ROW_H, minWidth: colW(sel.c), height: ROW_H }}
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
    </EditorShell>
  );
}
