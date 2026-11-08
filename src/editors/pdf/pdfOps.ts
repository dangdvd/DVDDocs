// Các thao tác chỉnh sửa PDF thuần (không phụ thuộc giao diện), dùng pdf-lib.
import {
  PDFCheckBox, PDFDocument, PDFDropdown, PDFFont, PDFName, PDFOptionList, PDFPage, PDFRadioGroup, PDFTextField,
  StandardFonts, degrees, rgb,
} from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { Box, stripTextInBoxes } from './contentText';
import { FontStyle, fontKey, loadFontBytes } from './fonts';

/** Toạ độ "hiển thị" tính theo điểm (pt) ở tỉ lệ 1, gốc góc trên-trái của trang đã xoay. */
export type Annot =
  | { id: string; page: number; type: 'text'; x: number; y: number; text: string; size: number; color: string }
  | { id: string; page: number; type: 'highlight' | 'rect' | 'whiteout'; x: number; y: number; w: number; h: number; color: string }
  | { id: string; page: number; type: 'ink'; points: [number, number][]; color: string; width: number }
  | { id: string; page: number; type: 'image'; x: number; y: number; w: number; h: number; dataUrl: string }
  /** Sửa chữ có sẵn: che vùng chữ cũ (x, y, w, h) bằng màu nền `bg` rồi viết chữ mới cùng vị trí và cỡ. */
  | { id: string; page: number; type: 'edit'; x: number; y: number; w: number; h: number; text: string; size: number; color: string; bg: string; font?: FontStyle };

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

const loadUnicodeFont = () => loadFontBytes();

function dataUrlBytes(dataUrl: string): { bytes: Uint8Array; png: boolean } {
  const [head, b64] = dataUrl.split(',');
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return { bytes, png: head.includes('png') };
}

/** Giá trị trường biểu mẫu: chữ (ô nhập, danh sách, nút chọn một) hoặc đúng/sai (ô đánh dấu). */
export type FieldValues = Record<string, string | boolean>;

/** Điền giá trị vào các trường biểu mẫu (AcroForm). Trường không tồn tại hoặc giá trị không hợp lệ được bỏ qua. */
async function fillFields(doc: PDFDocument, fields: FieldValues, getFont: () => Promise<PDFFont>) {
  const form = doc.getForm();
  let text = false;
  for (const [name, val] of Object.entries(fields)) {
    try {
      const f = form.getField(name);
      if (f instanceof PDFTextField) { f.setText(String(val)); text = true; }
      else if (f instanceof PDFCheckBox) val ? f.check() : f.uncheck();
      else if (f instanceof PDFRadioGroup) {
        const v = String(val);
        if (!v) f.clear();
        else if (f.getOptions().includes(v)) f.select(v);
        // pdf.js trả về tên trạng thái của nút (ví dụ "0", "1") thay vì nhãn lựa chọn.
        else f.acroField.setValue(PDFName.of(v));
      }
      else if (f instanceof PDFDropdown || f instanceof PDFOptionList) { if (val) f.select(String(val)); else f.clear(); }
    } catch {
      /* trường lạ hoặc giá trị không nằm trong danh sách: bỏ qua */
    }
  }
  // Vẽ lại hình hiển thị của trường bằng phông có dấu tiếng Việt.
  if (text) form.updateFieldAppearances(await getFont());
}

/** Ghi (làm phẳng) các chú thích và điền biểu mẫu vào nội dung PDF. */
export async function applyAnnotations(
  bytes: Uint8Array,
  annots: Annot[],
  fontBytes?: ArrayBuffer | null,
  fields?: FieldValues,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const hasFields = !!fields && Object.keys(fields).length > 0;
  if (annots.length === 0 && !hasFields) return doc.save();
  let font: PDFFont | null = null;
  const getFont = async () => {
    if (font) return font;
    doc.registerFontkit(fontkit);
    const fb = fontBytes === undefined ? await loadUnicodeFont() : fontBytes;
    font = fb ? await doc.embedFont(fb, { subset: true }) : await doc.embedFont(StandardFonts.Helvetica);
    return font;
  };
  // Phông gần giống phông gốc cho chữ sửa (có chân, đậm, nghiêng, đơn cách); thiếu tệp thì dùng phông mặc định.
  const styled = new Map<string, PDFFont>();
  const fontFor = async (s: FontStyle | undefined, text: string): Promise<PDFFont> => {
    const base = await getFont();
    if (!s || fontKey(s) === 'sans' || fontBytes !== undefined) return base;
    const k = fontKey(s);
    if (!styled.has(k)) {
      const fb = await loadFontBytes(s);
      styled.set(k, fb ? await doc.embedFont(fb, { subset: true }) : base);
    }
    const f = styled.get(k)!;
    // Phông thiếu chữ (ví dụ phông đơn cách thiếu một số dấu tiếng Việt): dùng phông không chân cùng kiểu đậm/nghiêng.
    const has = new Set(f.getCharacterSet());
    if ([...text].every((ch) => ch === '\n' || has.has(ch.codePointAt(0)!))) return f;
    return s.family === 'sans' ? base : fontFor({ ...s, family: 'sans' }, '');
  };
  if (hasFields) await fillFields(doc, fields!, getFont);
  if (annots.some((a) => a.type === 'text' || a.type === 'edit')) await getFont();
  const pages = doc.getPages();
  // Sửa chữ có sẵn: xoá hẳn chữ cũ khỏi nội dung trang (nếu đọc được), sau đó vẫn che nền để chắc chắn.
  const boxes = new Map<number, Box[]>();
  for (const a of annots) {
    if (a.type !== 'edit' || !pages[a.page]) continue;
    const p0 = toPdfPoint(pages[a.page], a.x, a.y), p1 = toPdfPoint(pages[a.page], a.x + a.w, a.y + a.h);
    const list = boxes.get(a.page) ?? [];
    list.push({ x0: Math.min(p0[0], p1[0]), y0: Math.min(p0[1], p1[1]), x1: Math.max(p0[0], p1[0]), y1: Math.max(p0[1], p1[1]) });
    boxes.set(a.page, list);
  }
  for (const [pi, list] of boxes) {
    try { stripTextInBoxes(doc, pages[pi], list); } catch { /* luồng nội dung lạ: chỉ che */ }
  }
  for (const a of annots) {
    const page = pages[a.page];
    if (!page) continue;
    const rot = degrees(normRotation(page.getRotation().angle));
    if (a.type === 'edit' && font) {
      // Che chữ gốc (rộng hơn một chút để không lộ viền khử răng cưa) rồi viết chữ mới.
      const [x, y] = toPdfPoint(page, a.x - 1, a.y + a.h + 0.5);
      page.drawRectangle({ x, y, width: a.w + 2, height: a.h + 1, color: hexToRgb(a.bg), rotate: rot });
    }
    if ((a.type === 'text' || a.type === 'edit') && font) {
      const f = a.type === 'edit' ? await fontFor(a.font, a.text) : font;
      const lineH = a.size * 1.25;
      a.text.split('\n').forEach((line, i) => {
        const [x, y] = toPdfPoint(page, a.x, a.y + a.size * 0.95 + i * lineH);
        page.drawText(line, { x, y, size: a.size, font: f, color: hexToRgb(a.color), rotate: rot });
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

/** Đọc chuỗi trang kiểu "1-3, 5, 8-" (đánh số từ 1) thành danh sách chỉ số trang (từ 0), giữ thứ tự, bỏ trùng. */
export function parsePageRanges(spec: string, count: number): number[] {
  const out: number[] = [];
  for (const part of spec.split(/[,;\s]+/).filter(Boolean)) {
    const m = /^(\d*)\s*-\s*(\d*)$/.exec(part);
    let a: number, b: number;
    if (m) { a = m[1] ? Number(m[1]) : 1; b = m[2] ? Number(m[2]) : count; }
    else if (/^\d+$/.test(part)) a = b = Number(part);
    else throw new Error(`Không hiểu "${part}". Ví dụ hợp lệ: 1-3, 5`);
    if (a < 1 || b > count || a > b) throw new Error(`Trang "${part}" nằm ngoài tài liệu (1–${count}).`);
    for (let i = a; i <= b; i++) if (!out.includes(i - 1)) out.push(i - 1);
  }
  if (!out.length) throw new Error('Chưa nhập trang nào.');
  return out;
}

/** Tạo PDF mới chỉ gồm các trang đã chọn. */
export async function extractPages(bytes: Uint8Array, indices: number[]) {
  const src = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const doc = await PDFDocument.create();
  (await doc.copyPages(src, indices)).forEach((p) => doc.addPage(p));
  return doc.save();
}

/** Tạo PDF trắng (khổ A4) để bắt đầu từ đầu. */
export async function createBlankPdf() {
  const doc = await PDFDocument.create();
  doc.addPage([595.28, 841.89]);
  return doc.save();
}
