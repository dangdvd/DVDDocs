// Đọc/ghi .xlsx bằng ExcelJS. Khi lưu, ghi đè lên workbook gốc để giữ nguyên những gì ta không sửa.
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Book, Cell, CellStyle, Engine, Sheet, addr, dateToSerial, key, parseKey, parseRange, rangeRef } from './sheetModel';
import { ChartData, Chart, chartData, readCharts, writeCharts } from './charts';
import { autoFilterXml, columnValues, filterHidden, readAutoFilters, readCf, writeAutoFilters, writeCf } from './cf';
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

/**
 * Một số chương trình (ví dụ openpyxl, LibreOffice) ghi drawing.xml không có tiền tố "xdr:" — ExcelJS không đọc được.
 * Thêm tiền tố cho các thẻ không có tiền tố để ExcelJS đọc như tệp của Excel.
 */
export async function normalizeDrawings(data: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data);
  let changed = false;
  for (const name of Object.keys(zip.files).filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))) {
    const xml = await zip.file(name)!.async('string');
    if (!/<wsDr[\s>]/.test(xml)) continue;
    const fixed = xml
      .replace(/<(\/?)([A-Za-z][\w.-]*)(?=[\s/>])/g, (m, slash: string, tag: string) => `<${slash}xdr:${tag}`)
      .replace(/(<xdr:wsDr\b[^>]*?)\sxmlns="([^"]*)"/, '$1 xmlns:xdr="$2"')
      // Không gian tên mặc định khai báo lại bên trong (hiếm): đổi thành xdr.
      .replace(/\sxmlns="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/spreadsheetDrawing"/g, '');
    zip.file(name, fixed);
    changed = true;
  }
  return changed ? zip.generateAsync({ type: 'arraybuffer' }) : data;
}

export async function loadXlsx(data: ArrayBuffer): Promise<LoadedBook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(data);
  } catch (e) {
    const fixed = await normalizeDrawings(data);
    if (fixed === data) throw e;
    await wb.xlsx.load(fixed);
  }
  const idMap = new Map<string, number>();
  const sheets: Sheet[] = [];
  wb.eachSheet((ws) => {
    const sheet: Sheet = { id: uid(), name: ws.name, cells: {}, colWidths: {}, rowCount: 100, colCount: 26 };
    idMap.set(sheet.id, ws.id);
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      if (row.height) (sheet.rowHeights ??= {})[rn - 1] = Math.max(8, Math.round((row.height * 4) / 3));
      if (row.hidden) (sheet.hiddenRows ??= {})[rn - 1] = true;
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
    try { const cf = readCf(ws); if (cf.length) sheet.cf = cf; } catch { /* bỏ qua */ }
    const af = ws.autoFilter as unknown as string | { from: unknown; to: unknown } | undefined;
    const end = (e: unknown) => (typeof e === 'string' ? e : e && typeof e === 'object' ? addr(Number((e as { row: number }).row) - 1, Number((e as { column: number }).column) - 1) : '');
    const afRef = typeof af === 'string' ? af : af && typeof af === 'object' ? `${end(af.from)}:${end(af.to)}` : undefined;
    const fr = afRef ? parseRange(afRef) : null;
    if (fr) sheet.filter = { r: fr.r0, c0: fr.c0, c1: fr.c1, r1: fr.r1, hide: {} };
    sheet.rowCount = Math.max(100, ws.rowCount + 20);
    sheet.colCount = Math.max(26, ws.columnCount + 5);
    sheets.push(sheet);
  });
  try {
    const charts = await readCharts(data, uid);
    for (const sh of sheets) { const c = charts.get(sh.name); if (c) sh.charts = c; }
  } catch { /* biểu đồ hỏng: vẫn mở được dữ liệu */ }
  // Tiêu chí lọc: đổi danh sách giá trị còn hiện thành giá trị bị ẩn; hàng ẩn trong vùng lọc là do lọc.
  if (sheets.some((s) => s.filter)) {
    try {
      const crit = await readAutoFilters(data);
      const engine = new Engine({ sheets });
      sheets.forEach((sh, si) => {
        const f = sh.filter, fc = crit.get(sh.name);
        if (!f) return;
        if (fc) for (const [col, { vals, blank }] of Object.entries(fc.shown)) {
          const c = f.c0 + Number(col), keep = new Set(vals);
          const hide = columnValues(sh, engine, si, f, c).filter((v) => (v === '' ? !blank : !keep.has(v)));
          if (hide.length) f.hide[c] = hide;
        }
        if (!fc?.unsupported && sh.hiddenRows) for (let r = f.r + 1; r <= f.r1; r++) delete sh.hiddenRows[r];
      });
    } catch { /* bỏ qua */ }
  }
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
  const engine = new Engine(book);
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
    // Hàng ẩn: ẩn thủ công và ẩn do lọc.
    const hidden = { ...(sheet.hiddenRows ?? {}), ...(sheet.filter ? filterHidden(sheet, engine, si, sheet.filter) : {}) };
    ws.eachRow({ includeEmpty: true }, (row, rn) => { if (row.hidden && !hidden[rn - 1]) row.hidden = false; });
    for (const r of Object.keys(hidden)) ws.getRow(Number(r) + 1).hidden = true;
    const f = sheet.filter;
    ws.autoFilter = f ? `${addr(f.r, f.c0)}:${addr(f.r1, f.c1)}` : (undefined as never);
    writeCf(ws, sheet.cf);
    const fr = sheet.freeze;
    ws.views = fr && (fr.rows || fr.cols)
      ? [{ state: 'frozen', xSplit: fr.cols, ySplit: fr.rows, topLeftCell: addr(fr.rows, fr.cols) }]
      : [{ state: 'normal' }];
  });
  wb.worksheets.filter((ws) => !keep.has(ws.id)).forEach((ws) => wb.removeWorksheet(ws.id));
  order.forEach((ws, i) => ((ws as unknown as { orderNo: number }).orderNo = i));
  const raw = await wb.xlsx.writeBuffer();
  const buf = raw instanceof ArrayBuffer ? raw : (new Uint8Array(raw as ArrayBufferLike).buffer as ArrayBuffer);
  // ExcelJS không ghi biểu đồ: thêm vào gói sau khi ghi.
  const charts = new Map<string, { chart: Chart; data: ChartData }[]>();
  book.sheets.forEach((s, si) => {
    if (s.charts?.length) charts.set(s.name, s.charts.map((chart) => ({ chart, data: chartData(chart, book, engine, si) })));
  });
  const filters = new Map<string, string>();
  book.sheets.forEach((s, si) => { if (s.filter) filters.set(s.name, autoFilterXml(s.filter, s, engine, si)); });
  return writeAutoFilters(await writeCharts(buf, charts), filters);
}
