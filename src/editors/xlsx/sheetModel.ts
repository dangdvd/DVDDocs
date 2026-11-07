// Mô hình bảng tính thuần (không phụ thuộc giao diện): ô, định dạng, công thức, dịch tham chiếu.
import FormulaParser from 'fast-formula-parser';

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // #RRGGBB
  fill?: string; // #RRGGBB
  align?: 'left' | 'center' | 'right';
  numFmt?: string;
  /** Viền mảnh theo từng cạnh: trên, phải, dưới, trái. */
  border?: { t?: boolean; r?: boolean; b?: boolean; l?: boolean };
  /** Xuống dòng trong ô. */
  wrap?: boolean;
}

/** Vùng ô gộp (chỉ số từ 0, bao gồm cả hai đầu). */
export interface Merge { r0: number; c0: number; r1: number; c1: number }

export interface Cell {
  /** Giá trị nhập (khi không phải công thức). */
  v?: string | number | boolean | null;
  /** Công thức, không có dấu "=" ở đầu. */
  f?: string;
  s?: CellStyle;
}

export interface Sheet {
  id: string;
  name: string;
  cells: Record<string, Cell>;
  colWidths: Record<number, number>; // px
  /** Chiều cao hàng tuỳ chỉnh (px). */
  rowHeights?: Record<number, number>;
  rowCount: number;
  colCount: number;
  merges?: Merge[];
  /** Số hàng phía trên và số cột bên trái được cố định khi cuộn. */
  freeze?: { rows: number; cols: number };
}

export interface Book { sheets: Sheet[] }

export const key = (r: number, c: number) => `${r},${c}`;
export const parseKey = (k: string) => k.split(',').map(Number) as [number, number];

export function colName(c: number): string {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r: number, c: number) => `${colName(c)}${r + 1}`;

/** Chuyển chữ người dùng gõ thành nội dung ô. */
export function parseInput(text: string): Pick<Cell, 'v' | 'f'> {
  if (text.startsWith('=') && text.length > 1) return { f: text.slice(1) };
  const t = text.trim();
  if (t === '') return { v: null };
  if (/^(true|false)$/i.test(t)) return { v: t.toLowerCase() === 'true' };
  const num = t.replace(/,(?=\d{3}(\D|$))/g, '');
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?%?$/.test(num)) {
    return num.endsWith('%') ? { v: parseFloat(num) / 100 } : { v: parseFloat(num) };
  }
  return { v: text };
}

/** Nội dung hiển thị trên thanh công thức. */
export function rawText(cell?: Cell): string {
  if (!cell) return '';
  if (cell.f) return `=${cell.f}`;
  if (cell.v === null || cell.v === undefined) return '';
  return String(cell.v);
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
export const dateToSerial = (d: Date) => (d.getTime() - EXCEL_EPOCH) / 86400000;
export const serialToDate = (n: number) => new Date(EXCEL_EPOCH + Math.round(n * 86400000));

export function formatValue(v: unknown, numFmt?: string): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'object' && v && 'error' in v) return String((v as { error: string }).error);
  if (typeof v !== 'number') return String(v);
  if (!Number.isFinite(v)) return '#NUM!';
  const fmt = numFmt && numFmt !== 'General' ? numFmt : '';
  if (fmt) {
    const isDate = /[dy]/i.test(fmt.replace(/"[^"]*"/g, '')) || /h+:m+/i.test(fmt);
    if (isDate) {
      const d = serialToDate(v);
      const p = (n: number) => String(n).padStart(2, '0');
      const date = `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
      return /h/i.test(fmt) ? `${date} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}` : date;
    }
    const dec = /0\.(0+)/.exec(fmt)?.[1].length ?? 0;
    const grouping = fmt.includes('#,##') || fmt.includes(',');
    const pct = fmt.includes('%');
    const n = pct ? v * 100 : v;
    let s = n.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: grouping });
    if (pct) s += '%';
    const cur = /"([^"]+)"|\[\$([^\]-]+)/.exec(fmt);
    if (cur) {
      const sym = cur[1] ?? cur[2];
      s = fmt.trim().startsWith('"') || fmt.trim().startsWith('[$') ? `${sym}${s}` : `${s} ${sym.trim()}`;
    }
    return s;
  }
  if (Number.isInteger(v)) return String(v);
  return String(parseFloat(v.toPrecision(12)));
}

const REF_RE = /((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(])/g;

/**
 * Dịch các tham chiếu ô trong công thức khi chèn/xoá hàng hoặc cột.
 * axis 'row' | 'col'; at = vị trí (0-based); delta > 0 chèn, < 0 xoá.
 */
export function shiftFormula(f: string, sheetName: string, targetSheet: string, axis: 'row' | 'col', at: number, delta: number): string {
  const parts = f.split(/("(?:[^"]|"")*")/); // bỏ qua chuỗi trong ngoặc kép
  return parts.map((part, i) => i % 2 ? part : part.replace(REF_RE, (m, sheet: string | undefined, dc, col, dr, row) => {
    const refSheet = sheet ? sheet.slice(0, -1).replace(/^'|'$/g, '') : sheetName;
    if (refSheet !== targetSheet) return m;
    let r = Number(row) - 1;
    let c = colIndex(col);
    const pos = axis === 'row' ? r : c;
    if (pos >= at) {
      if (delta < 0 && pos < at - delta) return `${sheet ?? ''}#REF!`;
      if (axis === 'row') r += delta; else c += delta;
    }
    return `${sheet ?? ''}${dc}${colName(c)}${dr}${r + 1}`;
  })).join('');
}

/** Dịch tham chiếu tương đối khi sao chép công thức sang vị trí khác (như Excel). */
export function moveFormula(f: string, dRow: number, dCol: number): string {
  const parts = f.split(/("(?:[^"]|"")*")/);
  return parts.map((part, i) => i % 2 ? part : part.replace(REF_RE, (_m, sheet = '', dc, col, dr, row) => {
    const c = dc ? colIndex(col) : colIndex(col) + dCol;
    const r = dr ? Number(row) - 1 : Number(row) - 1 + dRow;
    if (c < 0 || r < 0) return `${sheet}#REF!`;
    return `${sheet}${dc}${colName(c)}${dr}${r + 1}`;
  })).join('');
}

export function usedRange(sheet: Sheet) {
  let maxR = -1, maxC = -1;
  for (const k in sheet.cells) {
    const [r, c] = parseKey(k);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  return { maxR, maxC };
}

/** Bộ tính công thức: tính lười, có ghi nhớ và phát hiện tham chiếu vòng. */
export class Engine {
  private cache = new Map<string, unknown>();
  private visiting = new Set<string>();
  private extents = new Map<number, { maxR: number; maxC: number }>();
  private parser: any;

  constructor(private book: Book) {
    this.parser = new FormulaParser({
      onCell: ({ sheet, row, col }: { sheet: string; row: number; col: number }) =>
        this.valueAt(this.sheetIdx(sheet), row - 1, col - 1),
      onRange: (ref: { sheet: string; from: { row: number; col: number }; to: { row: number; col: number } }) => {
        const si = this.sheetIdx(ref.sheet);
        const ext = this.extent(si);
        const toR = Math.min(ref.to.row - 1, ext.maxR), toC = Math.min(ref.to.col - 1, ext.maxC);
        const out: unknown[][] = [];
        for (let r = ref.from.row - 1; r <= toR; r++) {
          const row: unknown[] = [];
          for (let c = ref.from.col - 1; c <= toC; c++) row.push(this.valueAt(si, r, c));
          out.push(row);
        }
        return out.length ? out : [[null]];
      },
    });
  }

  private sheetIdx(name: string) {
    const i = this.book.sheets.findIndex((s) => s.name.toLowerCase() === String(name).toLowerCase());
    return i < 0 ? 0 : i;
  }

  private extent(si: number) {
    let e = this.extents.get(si);
    if (!e) { e = usedRange(this.book.sheets[si]); this.extents.set(si, e); }
    return e;
  }

  valueAt(si: number, r: number, c: number): unknown {
    const cell = this.book.sheets[si]?.cells[key(r, c)];
    if (!cell) return null;
    if (cell.f) return this.evaluate(si, r, c);
    return cell.v ?? null;
  }

  evaluate(si: number, r: number, c: number): unknown {
    const id = `${si}:${r},${c}`;
    if (this.cache.has(id)) return this.cache.get(id);
    const sheet = this.book.sheets[si];
    const cell = sheet?.cells[key(r, c)];
    if (!cell?.f) return cell?.v ?? null;
    if (this.visiting.has(id)) return { error: '#CIRC!' };
    this.visiting.add(id);
    let res: unknown;
    try {
      res = this.parser.parse(cell.f, { sheet: sheet.name, row: r + 1, col: c + 1 });
      if (Array.isArray(res)) res = Array.isArray(res[0]) ? res[0][0] : res[0];
      if (res && typeof res === 'object' && '_error' in res) res = { error: (res as { _error: string })._error };
      if (res === undefined) res = 0;
    } catch (e) {
      const err = e as { _error?: string; error?: string };
      res = { error: err._error ?? err.error ?? '#ERROR!' };
    }
    this.visiting.delete(id);
    this.cache.set(id, res);
    return res;
  }
}

/** Vùng gộp chứa ô (r, c), nếu có. */
export function mergeAt(merges: Merge[] | undefined, r: number, c: number): Merge | undefined {
  return merges?.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
}

/** "A1:C2" → vùng gộp. */
export function parseRange(ref: string): Merge | null {
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(ref.trim());
  if (!m) return null;
  const a = { r: Number(m[2]) - 1, c: colIndex(m[1]) };
  const b = m[3] ? { r: Number(m[4]) - 1, c: colIndex(m[3]) } : a;
  return { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
}

export const rangeRef = (m: Merge) => `${addr(m.r0, m.c0)}:${addr(m.r1, m.c1)}`;

/**
 * Dịch một đoạn [a0, a1] khi chèn (delta > 0) hoặc xoá (delta < 0) hàng/cột tại vị trí `at`.
 * Trả về null nếu đoạn bị xoá hết.
 */
export function shiftSpan(a0: number, a1: number, at: number, delta: number): [number, number] | null {
  if (delta > 0) return [a0 >= at ? a0 + delta : a0, a1 >= at ? a1 + delta : a1];
  const end = at - delta; // các vị trí [at, end) bị xoá
  const map = (p: number, start: boolean) => (p < at ? p : p >= end ? p + delta : start ? at : at - 1);
  const n0 = map(a0, true), n1 = map(a1, false);
  return n1 < n0 ? null : [n0, n1];
}

/** Dịch các vùng gộp khi chèn/xoá hàng hoặc cột; bỏ vùng chỉ còn một ô. */
export function shiftMerges(merges: Merge[] | undefined, axis: 'row' | 'col', at: number, delta: number): Merge[] | undefined {
  if (!merges?.length) return merges;
  const out: Merge[] = [];
  for (const m of merges) {
    const span = axis === 'row' ? shiftSpan(m.r0, m.r1, at, delta) : shiftSpan(m.c0, m.c1, at, delta);
    if (!span) continue;
    const n = axis === 'row' ? { ...m, r0: span[0], r1: span[1] } : { ...m, c0: span[0], c1: span[1] };
    if (n.r1 > n.r0 || n.c1 > n.c0) out.push(n);
  }
  return out;
}

/** Văn bản hiển thị để tìm kiếm: công thức thì lấy kết quả, kèm cả chữ gốc. */
export function findMatches(sheet: Sheet, engine: Engine, si: number, query: string): [number, number][] {
  const q = query.normalize('NFC').toLocaleLowerCase('vi');
  if (!q) return [];
  const out: [number, number][] = [];
  for (const k in sheet.cells) {
    const cell = sheet.cells[k];
    const [r, c] = parseKey(k);
    const shown = formatValue(cell.f ? engine.evaluate(si, r, c) : cell.v, cell.s?.numFmt);
    const hay = `${shown}\u0000${rawText(cell)}`.normalize('NFC').toLocaleLowerCase('vi');
    if (hay.includes(q)) out.push([r, c]);
  }
  return out.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}

/** Thay chữ trong nội dung gốc của ô (không phân biệt hoa thường). Trả về null nếu không đổi. */
export function replaceInCell(cell: Cell, query: string, repl: string): Cell | null {
  const raw = rawText(cell);
  const esc = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const next = raw.normalize('NFC').replace(new RegExp(esc.normalize('NFC'), 'gi'), () => repl);
  if (next === raw.normalize('NFC')) return null;
  return { ...parseInput(next), s: cell.s };
}

export function emptySheet(name: string, id: string): Sheet {
  return { id, name, cells: {}, colWidths: {}, rowCount: 100, colCount: 26 };
}
