// Dựng bảng HTML để in một trang tính (hoặc vùng chọn): giữ giá trị đã tính, định dạng số, kiểu chữ,
// màu, căn lề, kẻ khung, ô gộp, độ rộng cột và hàng ẩn.
import { Engine, Sheet, formatValue, key, usedRange } from './sheetModel';
import { escapeHtml } from '../../lib/print';

export interface PrintArea { r0: number; c0: number; r1: number; c1: number }

/** Vùng cần in: vùng chọn nếu chọn nhiều ô, nếu không thì toàn bộ vùng có dữ liệu. Trả null nếu trang trống. */
export function printArea(sheet: Sheet, sel?: PrintArea): PrintArea | null {
  const u = usedRange(sheet);
  if (u.maxR < 0) return null;
  let maxR = u.maxR, maxC = u.maxC;
  for (const m of sheet.merges ?? []) { maxR = Math.max(maxR, m.r1); maxC = Math.max(maxC, m.c1); }
  // Vùng chọn nhiều ô (cắt bớt phần trống phía sau, ví dụ khi chọn cả trang tính).
  if (sel && (sel.r1 > sel.r0 || sel.c1 > sel.c0)) {
    const a = { r0: sel.r0, c0: sel.c0, r1: Math.min(sel.r1, maxR), c1: Math.min(sel.c1, maxC) };
    return a.r1 >= a.r0 && a.c1 >= a.c0 ? a : null;
  }
  return { r0: 0, c0: 0, r1: maxR, c1: maxC };
}

type Look = { fill?: string; color?: string } | null | undefined;

export function sheetToHtml(sheet: Sheet, engine: Engine, si: number, area: PrintArea, opts: {
  hidden?: Record<number, boolean>; look?: (r: number, c: number) => Look; defaultWidth: number; title?: string;
}): string {
  const covered = new Set<string>();
  const spans = new Map<string, { rs: number; cs: number }>();
  for (const m of sheet.merges ?? []) {
    if (m.r1 < area.r0 || m.r0 > area.r1 || m.c1 < area.c0 || m.c0 > area.c1) continue;
    const r0 = Math.max(m.r0, area.r0), c0 = Math.max(m.c0, area.c0), r1 = Math.min(m.r1, area.r1), c1 = Math.min(m.c1, area.c1);
    let rs = 0;
    for (let r = r0; r <= r1; r++) if (!opts.hidden?.[r]) rs++;
    spans.set(key(r0, c0), { rs: Math.max(1, rs), cs: c1 - c0 + 1 });
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) if (r !== r0 || c !== c0) covered.add(key(r, c));
  }
  const cols: string[] = [];
  for (let c = area.c0; c <= area.c1; c++) cols.push(`<col style="width:${sheet.colWidths[c] ?? opts.defaultWidth}px">`);
  const rows: string[] = [];
  for (let r = area.r0; r <= area.r1; r++) {
    if (opts.hidden?.[r]) continue;
    const tds: string[] = [];
    for (let c = area.c0; c <= area.c1; c++) {
      const k = key(r, c);
      if (covered.has(k)) continue;
      const cell = sheet.cells[k];
      const value = cell ? (cell.f ? engine.evaluate(si, r, c) : cell.v) : null;
      const st = cell?.s;
      const look = opts.look?.(r, c);
      const isNum = typeof value === 'number';
      const css: string[] = [];
      if (st?.bold) css.push('font-weight:700');
      if (st?.italic) css.push('font-style:italic');
      if (st?.underline) css.push('text-decoration:underline');
      const color = look?.color ?? st?.color, fill = look?.fill ?? st?.fill;
      if (color) css.push(`color:${color}`);
      if (fill) css.push(`background:${fill}`);
      css.push(`text-align:${st?.align ?? (isNum ? 'right' : typeof value === 'boolean' ? 'center' : 'left')}`);
      if (st?.wrap) css.push('white-space:pre-wrap');
      const b = st?.border;
      if (b?.t) css.push('border-top:1px solid #000');
      if (b?.r) css.push('border-right:1px solid #000');
      if (b?.b) css.push('border-bottom:1px solid #000');
      if (b?.l) css.push('border-left:1px solid #000');
      if (sheet.rowHeights?.[r]) css.push(`height:${sheet.rowHeights[r]}px`);
      const sp = spans.get(k);
      const attrs = sp ? `${sp.rs > 1 ? ` rowspan="${sp.rs}"` : ''}${sp.cs > 1 ? ` colspan="${sp.cs}"` : ''}` : '';
      tds.push(`<td${attrs} style="${css.join(';')}">${escapeHtml(formatValue(value, st?.numFmt))}</td>`);
    }
    rows.push(`<tr>${tds.join('')}</tr>`);
  }
  const title = opts.title ? `<div class="print-sheet-title">${escapeHtml(opts.title)}</div>` : '';
  return `${title}<table class="print-sheet"><colgroup>${cols.join('')}</colgroup><tbody>${rows.join('')}</tbody></table>`;
}
