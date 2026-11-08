// Phông chữ kèm ứng dụng (DejaVu, có đủ dấu tiếng Việt) và cách chọn phông gần giống phông gốc của PDF.

export type FontFamily = 'sans' | 'serif' | 'mono';
export interface FontStyle { family: FontFamily; bold: boolean; italic: boolean }

const FILES: Record<FontFamily, [string, string, string, string]> = {
  // thường, đậm, nghiêng, đậm nghiêng
  sans: ['DejaVuSans', 'DejaVuSans-Bold', 'DejaVuSans-Oblique', 'DejaVuSans-BoldOblique'],
  serif: ['DejaVuSerif', 'DejaVuSerif-Bold', 'DejaVuSerif-Italic', 'DejaVuSerif-BoldItalic'],
  mono: ['DejaVuSansMono', 'DejaVuSansMono-Bold', 'DejaVuSansMono-Oblique', 'DejaVuSansMono-BoldOblique'],
};

/** Tên CSS dùng để xem trước, trùng với phông nhúng khi lưu. */
export const CSS_FAMILY: Record<FontFamily, string> = { sans: 'DejaVu Sans', serif: 'DejaVu Serif', mono: 'DejaVu Sans Mono' };

export const fontFile = (s: Partial<FontStyle> = {}) =>
  `fonts/${FILES[s.family ?? 'sans'][(s.bold ? 1 : 0) + (s.italic ? 2 : 0)]}.ttf`;

export const fontKey = (s: Partial<FontStyle> = {}) => `${s.family ?? 'sans'}${s.bold ? '-b' : ''}${s.italic ? '-i' : ''}`;

/** Đoán kiểu phông từ tên phông trong PDF (ví dụ "ABCDEF+TimesNewRomanPS-BoldItalicMT") và gợi ý của pdf.js. */
export function guessFont(name: string | undefined, hint?: { serif?: boolean; mono?: boolean; bold?: boolean; italic?: boolean; family?: string }): FontStyle {
  const n = (name ?? '').replace(/^[A-Z]{6}\+/, '');
  const mono = hint?.mono || /courier|mono|consol|menlo|typewriter|fixed/i.test(n) || hint?.family === 'monospace';
  const sans = /sans|arial|helvetica|verdana|tahoma|calibri|segoe|roboto|gothic|futura|lato|inter\b|open ?sans|noto ?sans|myriad|frutiger|univers/i.test(n);
  const serif = !mono && !sans && (hint?.serif || /times|serif|roman|georgia|garamond|cambria|book|minion|palatino|baskerville|century|caslon|bodoni|didot|vni-times/i.test(n) || hint?.family === 'serif');
  return {
    family: mono ? 'mono' : serif ? 'serif' : 'sans',
    bold: !!hint?.bold || /bold|black|heavy|semibold|demibold|[-,]bd\b|\bbd$/i.test(n),
    italic: !!hint?.italic || /italic|oblique|[-,]it\b|[-,]i$/i.test(n),
  };
}

/** Chữ tiếng Việt mà phông đơn cách DejaVu không có: khi gặp thì dùng phông không chân (khớp với lúc lưu). */
const MONO_MISSING = /[ảẳẵắầẩẫấẻềểễếỉỏồổỗốởủửỷẢẲẴẮẦẨẪẤẺỀỂỄẾỈỎỒỔỖỐỞỦỬỶ]/;
export const previewFamily = (s: Partial<FontStyle> | undefined, text: string): FontFamily =>
  s?.family === 'mono' && MONO_MISSING.test(text) ? 'sans' : s?.family ?? 'sans';

let registered = false;
/** Khai báo các phông cho trình duyệt (chỉ tải khi được dùng). */
export function registerPreviewFonts() {
  if (registered || typeof FontFace === 'undefined' || typeof document === 'undefined') return;
  registered = true;
  for (const family of Object.keys(FILES) as FontFamily[]) {
    for (const bold of [false, true]) for (const italic of [false, true]) {
      const url = new URL(fontFile({ family, bold, italic }), document.baseURI).href;
      const face = new FontFace(CSS_FAMILY[family], `url(${url})`, { weight: bold ? '700' : '400', style: italic ? 'italic' : 'normal' });
      document.fonts.add(face);
      // Phông thường tải sẵn (dùng nhiều nhất).
      if (!bold && !italic && family === 'sans') face.load().catch(() => {});
    }
  }
}

const cache = new Map<string, Promise<ArrayBuffer | null>>();
export function loadFontBytes(s: Partial<FontStyle> = {}): Promise<ArrayBuffer | null> {
  const k = fontKey(s);
  let p = cache.get(k);
  if (!p) {
    p = fetch(new URL(fontFile(s), document.baseURI).href)
      .then((r) => (r.ok ? r.arrayBuffer() : null))
      .catch(() => null);
    cache.set(k, p);
  }
  return p;
}
