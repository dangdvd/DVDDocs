// Các thao tác chỉnh sửa PDF thuần (không phụ thuộc giao diện), dùng pdf-lib.
import { PDFDocument, PDFFont, PDFPage, StandardFonts, degrees, rgb } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';

/** Toạ độ "hiển thị" tính theo điểm (pt) ở tỉ lệ 1, gốc góc trên-trái của trang đã xoay. */
export type Annot =
  | { id: string; page: number; type: 'text'; x: number; y: number; text: string; size: number; color: string }
  | { id: string; page: number; type: 'highlight' | 'rect' | 'whiteout'; x: number; y: number; w: number; h: number; color: string }
  | { id: string; page: number; type: 'ink'; points: [number, number][]; color: string; width: number }
  | { id: string; page: number; type: 'image'; x: number; y: number; w: number; h: number; dataUrl: string };

export function hexToRgb(hex: string) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  const n = m ? parseInt(m[1], 16) : 0;
  return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
}

/** Kích thước hiển thị (pt) của trang sau khi xoay. */
export function displaySize(page: PDFPage) {
  const box = page.getCropBox();
  const r = normRotation(page.getRotation().angle);
  return r % 180 === 0 ? { w: box.width, h: box.height } : { w: box.height, h: box.width };
}

export function normRotation(angle: number) {
  return (((Math.round(angle / 90) * 90) % 360) + 360) % 360;
}

/** Chuyển toạ độ hiển thị (gốc trên-trái) sang toạ độ không gian người dùng của PDF. */
export function toPdfPoint(page: PDFPage, dx: number, dy: number): [number, number] {
  const b = page.getCropBox();
  const x0 = b.x, y0 = b.y, x1 = b.x + b.width, y1 = b.y + b.height;
  switch (normRotation(page.getRotation().angle)) {
    case 90: return [x0 + dy, y0 + dx];
    case 180: return [x1 - dx, y0 + dy];
    case 270: return [x1 - dy, y1 - dx];
    default: return [x0 + dx, y1 - dy];
  }
}

let fontBytesCache: ArrayBuffer | null = null;
async function loadUnicodeFont(): Promise<ArrayBuffer | null> {
  if (fontBytesCache) return fontBytesCache;
  try {
    const res = await fetch(new URL('fonts/DejaVuSans.ttf', document.baseURI).href);
    if (!res.ok) return null;
    fontBytesCache = await res.arrayBuffer();
    return fontBytesCache;
  } catch {
    return null;
  }
}

function dataUrlBytes(dataUrl: string): { bytes: Uint8Array; png: boolean } {
  const [head, b64] = dataUrl.split(',');
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { bytes, png: head.includes('png') };
}

/** Ghi (làm phẳng) các chú thích vào nội dung PDF. */
export async function applyAnnotations(
  bytes: Uint8Array,
  annots: Annot[],
  fontBytes?: ArrayBuffer | null,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  if (annots.length === 0) return doc.save();
  let font: PDFFont | null = null;
  if (annots.some((a) => a.type === 'text')) {
    doc.registerFontkit(fontkit);
    const fb = fontBytes === undefined ? await loadUnicodeFont() : fontBytes;
    font = fb ? await doc.embedFont(fb, { subset: true }) : await doc.embedFont(StandardFonts.Helvetica);
  }
  const pages = doc.getPages();
  for (const a of annots) {
    const page = pages[a.page];
    if (!page) continue;
    const rot = degrees(normRotation(page.getRotation().angle));
    if (a.type === 'text' && font) {
      const lineH = a.size * 1.25;
      a.text.split('\n').forEach((line, i) => {
        const [x, y] = toPdfPoint(page, a.x, a.y + a.size * 0.95 + i * lineH);
        page.drawText(line, { x, y, size: a.size, font: font!, color: hexToRgb(a.color), rotate: rot });
      });
    } else if (a.type === 'highlight' || a.type === 'rect' || a.type === 'whiteout') {
      const [x, y] = toPdfPoint(page, a.x, a.y + a.h);
      const color = a.type === 'whiteout' ? rgb(1, 1, 1) : hexToRgb(a.color);
      page.drawRectangle({
        x, y, width: a.w, height: a.h, rotate: rot,
        ...(a.type === 'rect'
          ? { borderColor: color, borderWidth: 1.5 }
          : { color, opacity: a.type === 'highlight' ? 0.35 : 1 }),
      });
    } else if (a.type === 'ink') {
      for (let i = 1; i < a.points.length; i++) {
        const [sx, sy] = toPdfPoint(page, ...a.points[i - 1]);
        const [ex, ey] = toPdfPoint(page, ...a.points[i]);
        page.drawLine({ start: { x: sx, y: sy }, end: { x: ex, y: ey }, thickness: a.width, color: hexToRgb(a.color), lineCap: 1 });
      }
    } else if (a.type === 'image') {
      const { bytes: img, png } = dataUrlBytes(a.dataUrl);
      const embedded = png ? await doc.embedPng(img) : await doc.embedJpg(img);
      const [x, y] = toPdfPoint(page, a.x, a.y + a.h);
      page.drawImage(embedded, { x, y, width: a.w, height: a.h, rotate: rot });
    }
  }
  return doc.save();
}

export async function rotatePage(bytes: Uint8Array, index: number, delta: number) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const p = doc.getPage(index);
  p.setRotation(degrees(normRotation(p.getRotation().angle + delta)));
  return doc.save();
}

export async function deletePage(bytes: Uint8Array, index: number) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  if (doc.getPageCount() <= 1) throw new Error('Tài liệu phải còn ít nhất một trang.');
  doc.removePage(index);
  return doc.save();
}

export async function movePage(bytes: Uint8Array, from: number, to: number) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  if (to < 0 || to >= doc.getPageCount()) return doc.save();
  const [copied] = await doc.copyPages(doc, [from]);
  doc.removePage(from);
  doc.insertPage(to, copied);
  return doc.save();
}

export async function insertBlankPage(bytes: Uint8Array, at: number) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const ref = doc.getPage(Math.min(at, doc.getPageCount() - 1));
  const { width, height } = ref.getSize();
  doc.insertPage(at, [width, height]);
  return doc.save();
}

export async function mergePdf(bytes: Uint8Array, other: Uint8Array) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const src = await PDFDocument.load(other, { ignoreEncryption: true });
  const pages = await doc.copyPages(src, src.getPageIndices());
  pages.forEach((p) => doc.addPage(p));
  return doc.save();
}

/** Tạo PDF trắng (khổ A4) để bắt đầu từ đầu. */
export async function createBlankPdf() {
  const doc = await PDFDocument.create();
  doc.addPage([595.28, 841.89]);
  return doc.save();
}
