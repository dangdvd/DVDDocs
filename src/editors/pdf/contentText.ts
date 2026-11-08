// Xoá chữ gốc trong luồng nội dung của trang PDF tại các vùng được sửa,
// để chữ cũ không còn trong tệp (không tìm, không sao chép được), thay vì chỉ bị che.
import { PDFArray, PDFDocument, PDFName, PDFPage, PDFRawStream, PDFRef, PDFStream, decodePDFRawStream } from 'pdf-lib';

type M = [number, number, number, number, number, number];
const I: M = [1, 0, 0, 1, 0, 0];
const mul = (a: M, b: M): M => [
  a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5],
];

interface Tok { s: number; e: number; kind: 'num' | 'str' | 'arr' | 'name' | 'dict' | 'op' | 'other'; v: string }

/** Tách luồng nội dung thành các phần tử (giữ vị trí để sửa tại chỗ). */
function* tokens(src: string): Generator<Tok> {
  const n = src.length;
  let i = 0;
  const ws = (c: string) => c === ' ' || c === '\n' || c === '\r' || c === '\t' || c === '\f' || c === '\0';
  const delim = (c: string) => '()<>[]{}/%'.includes(c);
  while (i < n) {
    const c = src[i];
    if (ws(c)) { i++; continue; }
    if (c === '%') { while (i < n && src[i] !== '\n' && src[i] !== '\r') i++; continue; }
    const s = i;
    if (c === '(') {
      let depth = 0;
      for (; i < n; i++) {
        const ch = src[i];
        if (ch === '\\') { i++; continue; }
        if (ch === '(') depth++;
        else if (ch === ')' && --depth === 0) { i++; break; }
      }
      yield { s, e: i, kind: 'str', v: '' };
    } else if (c === '<' && src[i + 1] === '<') {
      // Từ điển (ví dụ thuộc tính của BDC): bỏ qua cả khối, có lồng nhau.
      let depth = 0;
      for (; i < n; i++) {
        if (src[i] === '<' && src[i + 1] === '<') { depth++; i++; } else if (src[i] === '>' && src[i + 1] === '>') { depth--; i++; if (!depth) { i++; break; } }
        else if (src[i] === '(') { for (let d = 0; i < n; i++) { if (src[i] === '\\') { i++; continue; } if (src[i] === '(') d++; else if (src[i] === ')' && --d === 0) break; } }
      }
      yield { s, e: i, kind: 'dict', v: '' };
    } else if (c === '<') {
      i = src.indexOf('>', i) + 1 || n;
      yield { s, e: i, kind: 'str', v: '' };
    } else if (c === '[') {
      // Mảng (TJ): đọc tới dấu ] tương ứng, bỏ qua chuỗi bên trong.
      let depth = 0;
      for (; i < n; i++) {
        const ch = src[i];
        if (ch === '(') { for (let d = 0; i < n; i++) { if (src[i] === '\\') { i++; continue; } if (src[i] === '(') d++; else if (src[i] === ')' && --d === 0) break; } continue; }
        if (ch === '<') { i = src.indexOf('>', i); if (i < 0) i = n; continue; }
        if (ch === '[') depth++;
        else if (ch === ']' && --depth === 0) { i++; break; }
      }
      yield { s, e: i, kind: 'arr', v: '' };
    } else if (c === '/') {
      i++;
      while (i < n && !ws(src[i]) && !delim(src[i])) i++;
      yield { s, e: i, kind: 'name', v: src.slice(s, i) };
    } else if (c === ']' || c === ')' || c === '>' || c === '{' || c === '}') {
      i++;
      yield { s, e: i, kind: 'other', v: c };
    } else {
      while (i < n && !ws(src[i]) && !delim(src[i])) i++;
      const v = src.slice(s, i);
      if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(v)) yield { s, e: i, kind: 'num', v };
      else {
        yield { s, e: i, kind: 'op', v };
        // Ảnh nội tuyến: dữ liệu nhị phân giữa ID và EI.
        if (v === 'ID') {
          const end = src.slice(i).search(/\sEI(\s|$)/);
          i = end < 0 ? n : i + end + 3;
        }
      }
    }
  }
}

export interface Box { x0: number; y0: number; x1: number; y1: number }

/**
 * Xoá các lệnh vẽ chữ (Tj, TJ, ', ") bắt đầu bên trong các vùng (toạ độ người dùng PDF).
 * Trả về số lệnh đã xoá; luồng nội dung chỉ được ghi lại khi có xoá.
 */
export function stripTextInBoxes(doc: PDFDocument, page: PDFPage, boxes: Box[]): number {
  if (!boxes.length) return 0;
  const contents = page.node.get(PDFName.of('Contents'));
  const refs: (PDFRef | PDFStream)[] = contents instanceof PDFArray ? contents.asArray() as (PDFRef | PDFStream)[] : contents ? [contents as PDFRef] : [];
  const parts: string[] = [];
  for (const r of refs) {
    const st = r instanceof PDFRef ? doc.context.lookup(r) : r;
    if (!(st instanceof PDFRawStream)) return 0; // luồng lạ: không động vào
    const bytes = decodePDFRawStream(st).decode();
    let s = '';
    for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode(...bytes.subarray(i, i + 8192));
    parts.push(s);
  }
  const src = parts.join('\n');
  const edits: [number, number, string][] = [];
  let ctm: M = I, tm: M = I, tlm: M = I, leading = 0;
  const stack: M[] = [];
  let ops: Tok[] = [];
  const nums = () => ops.filter((t) => t.kind === 'num').map((t) => Number(t.v));
  const inside = () => {
    const m = mul(tm, ctm);
    const x = m[4], y = m[5];
    return boxes.some((b) => x >= b.x0 - 2 && x <= b.x1 + 2 && y >= b.y0 - 2 && y <= b.y1 + 2);
  };
  const td = (tx: number, ty: number) => { tlm = mul([1, 0, 0, 1, tx, ty], tlm); tm = tlm; };
  for (const t of tokens(src)) {
    if (t.kind !== 'op') { ops.push(t); continue; }
    const v = nums();
    switch (t.v) {
      case 'q': stack.push(ctm); break;
      case 'Q': ctm = stack.pop() ?? I; break;
      case 'cm': if (v.length >= 6) ctm = mul(v.slice(-6) as M, ctm); break;
      case 'BT': tm = tlm = I; break;
      case 'TL': if (v.length) leading = v[v.length - 1]; break;
      case 'Td': if (v.length >= 2) td(v[v.length - 2], v[v.length - 1]); break;
      case 'TD': if (v.length >= 2) { leading = -v[v.length - 1]; td(v[v.length - 2], v[v.length - 1]); } break;
      case 'Tm': if (v.length >= 6) { tlm = tm = v.slice(-6) as M; } break;
      case 'T*': td(0, -leading); break;
      case "'": case '"': td(0, -leading); /* fallthrough */
      case 'Tj': case 'TJ': {
        const arg = [...ops].reverse().find((o) => o.kind === (t.v === 'TJ' ? 'arr' : 'str'));
        if (arg && inside()) edits.push([arg.s, arg.e, t.v === 'TJ' ? '[]' : '()']);
        break;
      }
    }
    ops = [];
  }
  if (!edits.length) return 0;
  let out = '';
  let pos = 0;
  for (const [s, e, rep] of edits.sort((a, b) => a[0] - b[0])) { out += src.slice(pos, s) + rep; pos = e; }
  out += src.slice(pos);
  const bytes = new Uint8Array(out.length);
  for (let i = 0; i < out.length; i++) bytes[i] = out.charCodeAt(i) & 255;
  page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.flateStream(bytes)));
  return edits.length;
}
