// Mở tệp CSV/TSV thành một trang tính (khi lưu sẽ thành .xlsx).
import type { Book, Cell } from './sheetModel';
import { emptySheet, key, parseInput } from './sheetModel';

/** Giải mã chữ: UTF-8 (bỏ BOM), nếu lỗi thì thử bảng mã Windows tiếng Việt. */
export function decodeText(data: ArrayBuffer): string {
  const bytes = new Uint8Array(data);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
  } catch {
    return new TextDecoder('windows-1258').decode(bytes);
  }
}

/** Đoán dấu phân cách từ vài dòng đầu (bỏ qua phần trong ngoặc kép). */
export function detectDelimiter(text: string): string {
  const head = text.slice(0, 4096).replace(/"[^"]*"/g, '');
  const lines = head.split(/\r?\n/).filter(Boolean).slice(0, 10);
  let best = ',', score = -1;
  for (const d of [',', ';', '\t', '|']) {
    const counts = lines.map((l) => l.split(d).length - 1);
    const min = Math.min(...counts);
    const s = min > 0 ? min * 10 + counts.reduce((a, b) => a + b, 0) / counts.length : 0;
    if (s > score) { score = s; best = d; }
  }
  return best;
}

/** Tách CSV thành các hàng (hỗ trợ ngoặc kép, "" bên trong, xuống dòng trong ô). */
export function parseCsv(text: string, delim = detectDelimiter(text)): string[][] {
  const rows: string[][] = [];
  let row: string[] = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else q = false;
      } else field += ch;
    } else if (ch === '"' && field === '') q = true;
    else if (ch === delim) { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/** Giá trị của một ô CSV: số, đúng/sai, công thức; giữ nguyên số có số 0 ở đầu (mã, số điện thoại). */
function csvCell(text: string, delim: string): Pick<Cell, 'v' | 'f'> {
  const t = text.trim();
  if (/^0\d+$/.test(t) || /^\+\d{6,}$/.test(t)) return { v: text };
  // Với dấu ";" (kiểu châu Âu/Việt Nam), dấu phẩy là phần thập phân: 1.234,5 → 1234.5
  if (delim === ';' && /^[-+]?\d{1,3}(\.\d{3})*(,\d+)?$|^[-+]?\d+,\d+$/.test(t)) return { v: Number(t.replace(/\./g, '').replace(',', '.')) };
  return parseInput(text);
}

export function csvToBook(data: ArrayBuffer, name: string, id: string): Book {
  const text = decodeText(data);
  const delim = /\.tsv$/i.test(name) ? '\t' : detectDelimiter(text);
  const rows = parseCsv(text, delim);
  const sheet = emptySheet(name.replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1', id);
  let maxC = 0;
  rows.forEach((row, r) => row.forEach((v, c) => {
    if (v === '') return;
    sheet.cells[key(r, c)] = csvCell(v, delim) as Cell;
    maxC = Math.max(maxC, c);
  }));
  // Độ rộng cột theo chữ dài nhất (giới hạn).
  for (let c = 0; c <= maxC; c++) {
    let w = 0;
    for (let r = 0; r < Math.min(rows.length, 200); r++) w = Math.max(w, (rows[r][c] ?? '').length);
    if (w > 10) sheet.colWidths[c] = Math.min(320, w * 8 + 16);
  }
  sheet.rowCount = Math.max(100, rows.length + 20);
  sheet.colCount = Math.max(26, maxC + 5);
  return { sheets: [sheet] };
}
