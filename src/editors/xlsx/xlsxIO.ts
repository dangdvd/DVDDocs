// Đọc/ghi .xlsx bằng ExcelJS. Khi lưu, ghi đè lên workbook gốc để giữ nguyên những gì ta không sửa.
import ExcelJS from 'exceljs';
import { Book, Cell, CellStyle, Sheet, addr, dateToSerial, key, parseKey, parseRange, rangeRef } from './sheetModel';
import { uid } from '../../lib/uid';

const argbToHex = (argb?: string) => (argb && argb.length >= 6 ? `#${argb.slice(-6)}` : undefined);
const hexToArgb = (hex?: string) => (hex ? `FF${hex.replace('#', '').toUpperCase()}` : undefined);

function readStyle(c: ExcelJS.Cell): CellStyle | undefined {
  const s: CellStyle = {};
  if (c.font?.bold) s.bold = true;
  if (c.font?.italic) s.italic = true;
  if (c.font?.underline) s.underline = true;
  const color = argbToHex(c.font?.color?.argb);
  if (color && color !== '#000000') s.color = color;
  const fill = c.fill as ExcelJS.FillPattern | undefined;
  if (fill?.type === 'pattern' && fill.pattern === 'solid') s.fill = argbToHex(fill.fgColor?.argb);
  const h = c.alignment?.horizontal;
  if (h === 'left' || h === 'center' || h === 'right') s.align = h;
  if (c.numFmt && c.numFmt !== 'General') s.numFmt = c.numFmt;
  if (c.alignment?.wrapText) s.wrap = true;
  const b = c.border;
  if (b) {
    const has = (e?: Partial<ExcelJS.Border>) => !!e?.style;
    const border = { t: has(b.top), r: has(b.right), b: has(b.bottom), l: has(b.left) };
    if (border.t || border.r || border.b || border.l) s.border = border;
  }
  return Object.keys(s).length ? s : undefined;
}

function readValue(v: ExcelJS.CellValue): Pick<Cell, 'v' | 'f'> {
  if (v === null || v === undefined) return { v: null };
  if (v instanceof Date) return { v: dateToSerial(v) };
  if (typeof v === 'object') {
    if ('formula' in v || 'sharedFormula' in v) {
      const f = (v as ExcelJS.CellFormulaValue).formula;
      if (f) return { f };
      const r = (v as ExcelJS.CellFormulaValue).result;
      return { v: r instanceof Date ? dateToSerial(r) : typeof r === 'object' ? null : (r as never) };
    }
    if ('richText' in v) return { v: v.richText.map((t) => t.text).join('') };
    if ('text' in v) return { v: String((v as { text: unknown }).text) };
    if ('error' in v) return { v: String(v.error) };
    return { v: null };
  }
  return { v: v as string | number | boolean };
}

export interface LoadedBook { book: Book; wb: ExcelJS.Workbook; idMap: Map<string, number> }

export async function loadXlsx(data: ArrayBuffer): Promise<LoadedBook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const idMap = new Map<string, number>();
  const sheets: Sheet[] = [];
  wb.eachSheet((ws) => {
    const sheet: Sheet = { id: uid(), name: ws.name, cells: {}, colWidths: {}, rowCount: 100, colCount: 26 };
    idMap.set(sheet.id, ws.id);
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      if (row.height) (sheet.rowHeights ??= {})[rn - 1] = Math.max(8, Math.round((row.height * 4) / 3));
    });
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      row.eachCell({ includeEmpty: false }, (c, cn) => {
        // Ô gộp phụ: bỏ qua nội dung (Excel chỉ giữ ô chính).
        if (c.isMerged && c.master !== c) return;
        const cell: Cell = { ...readValue(c.value) };
        const s = readStyle(c);
        if (s) cell.s = s;
        if (cell.f || (cell.v !== null && cell.v !== undefined) || cell.s) sheet.cells[key(rn - 1, cn - 1)] = cell;
      });
    });
    ws.columns?.forEach((col, i) => {
      if (col.width) sheet.colWidths[i] = Math.round(col.width * 7.5);
    });
    const merges = (ws.model as { merges?: string[] }).merges ?? [];
    sheet.merges = merges.map(parseRange).filter((m): m is NonNullable<typeof m> => !!m);
    const view = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number } | undefined;
    if (view?.state === 'frozen' && ((view.xSplit ?? 0) > 0 || (view.ySplit ?? 0) > 0)) {
      sheet.freeze = { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 };
    }
    sheet.rowCount = Math.max(100, ws.rowCount + 20);
    sheet.colCount = Math.max(26, ws.columnCount + 5);
    sheets.push(sheet);
  });
  if (!sheets.length) sheets.push({ id: uid(), name: 'Sheet1', cells: {}, colWidths: {}, rowCount: 100, colCount: 26 });
  return { book: { sheets }, wb, idMap };
}

function writeCell(c: ExcelJS.Cell, cell: Cell | undefined, result: unknown) {
  if (!cell) {
    c.value = null;
    c.style = {};
    return;
  }
  if (cell.f) {
    const r = result && typeof result === 'object' ? undefined : (result as string | number | boolean | undefined);
    c.value = { formula: cell.f, result: r } as ExcelJS.CellFormulaValue;
  } else c.value = cell.v ?? null;
  const s = cell.s ?? {};
  c.font = { ...(c.font ?? {}), bold: !!s.bold, italic: !!s.italic, underline: !!s.underline, color: s.color ? { argb: hexToArgb(s.color) } : undefined };
  if (s.fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(s.fill) } };
  else if ((c.fill as ExcelJS.FillPattern | undefined)?.pattern === 'solid') c.fill = { type: 'pattern', pattern: 'none' };
  c.alignment = { ...(c.alignment ?? {}), horizontal: s.align, wrapText: s.wrap || undefined };
  c.numFmt = s.numFmt ?? 'General';
  // Viền: giữ kiểu viền gốc của cạnh còn bật, cạnh mới bật dùng nét mảnh.
  const old = c.border ?? {};
  const thin: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FF000000' } };
  const side = (on: boolean | undefined, prev?: Partial<ExcelJS.Border>) => (on ? (prev?.style ? prev : thin) : undefined);
  const nb: Partial<ExcelJS.Borders> = {
    top: side(s.border?.t, old.top), right: side(s.border?.r, old.right),
    bottom: side(s.border?.b, old.bottom), left: side(s.border?.l, old.left),
  };
  if (old.diagonal) nb.diagonal = old.diagonal;
  c.border = nb as ExcelJS.Borders;
}

/**
 * Ghi mô hình hiện tại ra .xlsx. `evaluate` trả về kết quả công thức để lưu kèm (giúp ứng dụng khác hiển thị ngay).
 * Nếu `loaded` không có (tài liệu mới) thì tạo workbook mới.
 */
export async function saveXlsx(
  book: Book,
  loaded: LoadedBook | null,
  evaluate: (si: number, r: number, c: number) => unknown,
): Promise<ArrayBuffer> {
  const wb = loaded?.wb ?? new ExcelJS.Workbook();
  wb.creator = wb.creator || 'DVDDocs';
  wb.modified = new Date();
  const keep = new Set<number>();
  const order: ExcelJS.Worksheet[] = [];
  book.sheets.forEach((sheet, si) => {
    const wsId = loaded?.idMap.get(sheet.id);
    let ws = wsId !== undefined ? wb.getWorksheet(wsId) : undefined;
    if (!ws) {
      ws = wb.addWorksheet(sheet.name);
      loaded?.idMap.set(sheet.id, ws.id);
    }
    ws.name = sheet.name;
    keep.add(ws.id);
    order.push(ws);
    // Bỏ gộp cũ trước khi ghi ô, gộp lại theo mô hình sau khi ghi.
    for (const ref of [...((ws.model as { merges?: string[] }).merges ?? [])]) ws.unMergeCells(ref);
    // Xoá các ô cũ không còn trong mô hình.
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      if (row.height) (sheet.rowHeights ??= {})[rn - 1] = Math.max(8, Math.round((row.height * 4) / 3));
    });
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      row.eachCell({ includeEmpty: false }, (c, cn) => {
        if (!sheet.cells[key(rn - 1, cn - 1)] && !(c.isMerged && c.master !== c)) writeCell(c, undefined, null);
      });
    });
    for (const k in sheet.cells) {
      const [r, c] = parseKey(k);
      writeCell(ws.getCell(r + 1, c + 1), sheet.cells[k], sheet.cells[k].f ? evaluate(si, r, c) : null);
    }
    for (const [ci, px] of Object.entries(sheet.colWidths)) ws.getColumn(Number(ci) + 1).width = px / 7.5;
    for (const [ri, px] of Object.entries(sheet.rowHeights ?? {})) ws.getRow(Number(ri) + 1).height = Math.round(px * 0.75 * 100) / 100;
    for (const m of sheet.merges ?? []) {
      try { ws.mergeCells(rangeRef(m)); } catch { /* vùng chồng lấn: bỏ qua */ }
    }
    const fr = sheet.freeze;
    ws.views = fr && (fr.rows || fr.cols)
      ? [{ state: 'frozen', xSplit: fr.cols, ySplit: fr.rows, topLeftCell: addr(fr.rows, fr.cols) }]
      : [{ state: 'normal' }];
  });
  wb.worksheets.filter((ws) => !keep.has(ws.id)).forEach((ws) => wb.removeWorksheet(ws.id));
  order.forEach((ws, i) => ((ws as unknown as { orderNo: number }).orderNo = i));
  const buf = await wb.xlsx.writeBuffer();
  return buf instanceof ArrayBuffer ? buf : (new Uint8Array(buf as ArrayBufferLike).buffer as ArrayBuffer);
}
