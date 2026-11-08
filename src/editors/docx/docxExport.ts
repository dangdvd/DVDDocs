// Chuyển nội dung trình soạn thảo (JSON của TipTap/ProseMirror) sang tệp .docx thật bằng thư viện `docx`.
import {
  AlignmentType, BorderStyle, DeletedTextRun, Document, ExternalHyperlink, Header, InsertedTextRun, LineRuleType, TableOfContents, Footer, HeadingLevel, ImageRun, LevelFormat, Packer,
  PageBreak, PageNumber, PageOrientation, Paragraph, ParagraphChild, ShadingType, Table, TableCell, TableRow, TextRun, WidthType,
} from 'docx';

/** Thiết lập trang khi xuất: khổ ngang và đánh số trang ở chân trang. */
export interface PageSetup {
  landscape: boolean;
  pageNumbers: boolean;
  /** Bật theo dõi thay đổi khi mở trong Word. */
  track?: boolean;
  /** Chữ ở đầu trang / chân trang (mỗi dòng một đoạn, căn giữa). */
  header?: string;
  footer?: string;
}

/** "12pt" | "16px" | "1.5em" → nửa điểm (đơn vị cỡ chữ của Word), hoặc undefined. */
export function halfPoints(size?: string | null): number | undefined {
  const m = /^([\d.]+)\s*(pt|px)?$/i.exec(String(size ?? '').trim());
  if (!m) return undefined;
  const pt = m[2]?.toLowerCase() === 'px' ? Number(m[1]) * 0.75 : Number(m[1]);
  return pt > 0 ? Math.round(pt * 2) : undefined;
}

export interface PMNode {
  type: string;
  attrs?: Record<string, any>;
  content?: PMNode[];
  text?: string;
  marks?: { type: string; attrs?: Record<string, any> }[];
}

const HEADINGS = [
  HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6,
];

const ALIGN: Record<string, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED,
};

/** "#ff0000" | "rgb(255, 0, 0)" | "red" → "FF0000" (hoặc undefined). */
export function toHex(color?: string | null): string | undefined {
  if (!color) return undefined;
  const c = color.trim();
  let m = /^#([0-9a-f]{3})$/i.exec(c);
  if (m) return m[1].split('').map((ch) => ch + ch).join('').toUpperCase();
  m = /^#([0-9a-f]{6})$/i.exec(c);
  if (m) return m[1].toUpperCase();
  m = /^rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(c);
  if (m) return [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('').toUpperCase();
  const named: Record<string, string> = { red: 'FF0000', blue: '0000FF', green: '008000', black: '000000', yellow: 'FFFF00', white: 'FFFFFF' };
  return named[c.toLowerCase()];
}

export function decodeDataUrl(src: string): { bytes: Uint8Array; type: 'png' | 'jpg' | 'gif' | 'bmp' } | null {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/i.exec(src);
  if (!m) return null;
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const t = m[1].toLowerCase();
  return { bytes, type: t === 'jpeg' || t === 'jpg' ? 'jpg' : (t as 'png' | 'gif' | 'bmp') };
}

/** Đọc kích thước ảnh từ phần đầu tệp PNG/JPEG/GIF (không cần trình duyệt). */
export function imageSize(b: Uint8Array): { w: number; h: number } | null {
  if (b[0] === 0x89 && b[1] === 0x50) return { w: (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19], h: (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23] };
  if (b[0] === 0x47 && b[1] === 0x49) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const marker = b[i + 1];
      const len = (b[i + 2] << 8) | b[i + 3];
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc)
        return { h: (b[i + 5] << 8) | b[i + 6], w: (b[i + 7] << 8) | b[i + 8] };
      i += 2 + len;
    }
  }
  return null;
}

const MAX_IMG_W = 600;

class Converter {
  listInstance = 0;
  revision = 1;
  hasToc = false;

  runs(nodes: PMNode[] = []): ParagraphChild[] {
    const out: ParagraphChild[] = [];
    for (const n of nodes) {
      if (n.type === 'hardBreak') out.push(new TextRun({ text: '', break: 1 }));
      else if (n.type === 'image') {
        const img = this.image(n);
        if (img) out.push(img);
      } else if (n.type === 'text' && n.text) {
        const marks = n.marks ?? [];
        const has = (t: string) => marks.find((m) => m.type === t);
        const ts = has('textStyle')?.attrs;
        const color = toHex(ts?.color);
        const family = (ts?.fontFamily as string | undefined)?.replace(/["']/g, '').split(',')[0].trim();
        const hl = has('highlight');
        const link = has('link');
        const opts = {
          text: n.text,
          bold: !!has('bold'),
          italics: !!has('italic'),
          underline: has('underline') || link ? {} : undefined,
          strike: !!has('strike'),
          color: color ?? (link ? '0563C1' : undefined),
          font: has('code') ? 'Courier New' : family || undefined,
          size: halfPoints(ts?.fontSize),
          superScript: !!has('superscript'),
          subScript: !!has('subscript'),
          shading: hl ? { type: ShadingType.CLEAR, color: 'auto', fill: toHex(hl.attrs?.color) ?? 'FFFF00' } : undefined,
        };
        // Thay đổi được theo dõi → <w:ins>/<w:del> để Word hiện đúng và cho chấp nhận/từ chối.
        const change = has('insertion') ?? has('deletion');
        if (change) {
          const rev = { id: this.revision++, author: String(change.attrs?.author || 'DVDDocs'), date: String(change.attrs?.date || new Date().toISOString()) };
          out.push(change.type === 'insertion' ? new InsertedTextRun({ ...opts, ...rev }) : new DeletedTextRun({ ...opts, ...rev }));
          continue;
        }
        const run = new TextRun(opts);
        out.push(link?.attrs?.href ? new ExternalHyperlink({ link: link.attrs.href, children: [run] }) : run);
      }
    }
    return out;
  }

  image(n: PMNode): ImageRun | null {
    const decoded = decodeDataUrl(n.attrs?.src ?? '');
    if (!decoded) return null;
    const natural = imageSize(decoded.bytes) ?? { w: 300, h: 200 };
    let w = Number(n.attrs?.width) || natural.w;
    let h = Number(n.attrs?.height) || (natural.h / natural.w) * w;
    if (w > MAX_IMG_W) { h = (h / w) * MAX_IMG_W; w = MAX_IMG_W; }
    return new ImageRun({ type: decoded.type, data: decoded.bytes, transformation: { width: Math.round(w), height: Math.round(h) } });
  }

  blocks(nodes: PMNode[] = [], ctx: { list?: { kind: 'bullet' | 'ordered'; level: number; instance: number }; quote?: boolean } = {}): (Paragraph | Table)[] {
    const out: (Paragraph | Table)[] = [];
    for (const n of nodes) {
      switch (n.type) {
        case 'paragraph':
        case 'heading': {
          const align = ALIGN[n.attrs?.textAlign as string];
          out.push(new Paragraph({
            children: this.runs(n.content),
            heading: n.type === 'heading' ? HEADINGS[(n.attrs?.level ?? 1) - 1] : undefined,
            alignment: align,
            ...(Number(n.attrs?.lineHeight) > 0 ? { spacing: { line: Math.round(Number(n.attrs!.lineHeight) * 240), lineRule: LineRuleType.AUTO } } : {}),
            ...(ctx.list
              ? ctx.list.kind === 'bullet'
                ? { bullet: { level: ctx.list.level } }
                : { numbering: { reference: 'dvd-ordered', level: ctx.list.level, instance: ctx.list.instance } }
              : {}),
            ...(ctx.quote ? { indent: { left: 720 }, border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'BBBBBB', space: 8 } } } : {}),
          }));
          break;
        }
        case 'bulletList':
        case 'orderedList': {
          const level = ctx.list ? ctx.list.level + 1 : 0;
          const kind = n.type === 'bulletList' ? 'bullet' : 'ordered';
          const instance = kind === 'ordered' ? ++this.listInstance : 0;
          for (const item of n.content ?? []) {
            out.push(...this.blocks(item.content, { ...ctx, list: { kind, level: Math.min(level, 8), instance } }));
          }
          break;
        }
        case 'blockquote':
          out.push(...this.blocks(n.content, { ...ctx, quote: true }));
          break;
        case 'codeBlock':
          for (const line of (n.content?.map((c) => c.text ?? '').join('') ?? '').split('\n')) {
            out.push(new Paragraph({
              children: [new TextRun({ text: line, font: 'Courier New', size: 20 })],
              shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'F3F3F3' },
            }));
          }
          break;
        case 'tableOfContents':
          // Tiêu đề "Mục lục" không dùng kiểu Heading để không tự xuất hiện trong chính mục lục.
          out.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Mục lục', bold: true, size: 32, color: '2F5496' })] }));
          out.push(new TableOfContents('Mục lục', { hyperlink: true, headingStyleRange: '1-3' }) as unknown as Paragraph);
          this.hasToc = true;
          break;
        case 'pageBreak':
          out.push(new Paragraph({ children: [new PageBreak()] }));
          break;
        case 'horizontalRule':
          out.push(new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } } }));
          break;
        case 'image': {
          const img = this.image(n);
          if (img) out.push(new Paragraph({ children: [img] }));
          break;
        }
        case 'table':
          out.push(this.table(n));
          out.push(new Paragraph({ children: [] }));
          break;
        default:
          if (n.content) out.push(...this.blocks(n.content, ctx));
      }
    }
    return out;
  }

  table(n: PMNode): Table {
    const rows = (n.content ?? []).map((row) => new TableRow({
      children: (row.content ?? []).map((cell) => {
        const children = this.blocks(cell.content).filter(Boolean);
        return new TableCell({
          children: children.length ? children : [new Paragraph({ children: [] })],
          columnSpan: (cell.attrs?.colspan ?? 1) > 1 ? cell.attrs!.colspan : undefined,
          rowSpan: (cell.attrs?.rowspan ?? 1) > 1 ? cell.attrs!.rowspan : undefined,
          shading: toHex(cell.attrs?.backgroundColor)
            ? { type: ShadingType.CLEAR, color: 'auto', fill: toHex(cell.attrs?.backgroundColor) }
            : cell.type === 'tableHeader' ? { type: ShadingType.CLEAR, color: 'auto', fill: 'EDEFF2' } : undefined,
        });
      }),
    }));
    return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
  }
}

/** Mỗi dòng chữ đầu/chân trang thành một đoạn căn giữa, cỡ 9pt. */
const lines = (text: string) => text.split('\n').map((t) => new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, size: 18 })] }));

export function buildDocument(json: PMNode, setup: PageSetup = { landscape: false, pageNumbers: false }): Document {
  const conv = new Converter();
  const children = conv.blocks(json.content);
  const fmts = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN];
  return new Document({
    creator: 'DVDDocs',
    features: setup.track || conv.hasToc ? { trackRevisions: !!setup.track, updateFields: conv.hasToc } : undefined,
    styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
    numbering: {
      config: [{
        reference: 'dvd-ordered',
        levels: Array.from({ length: 9 }, (_, level) => ({
          level,
          format: fmts[level % 3],
          text: `%${level + 1}.`,
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
        })),
      }],
    },
    sections: [{
      properties: setup.landscape ? { page: { size: { orientation: PageOrientation.LANDSCAPE } } } : undefined,
      headers: setup.header?.trim() ? { default: new Header({ children: lines(setup.header) }) } : undefined,
      footers: setup.pageNumbers || setup.footer?.trim()
        ? {
          default: new Footer({
            children: [
              ...(setup.footer?.trim() ? lines(setup.footer) : []),
              ...(setup.pageNumbers ? [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ children: ['Trang ', PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], size: 18 })],
              })] : []),
            ],
          }),
        }
        : undefined,
      children: children.length ? children : [new Paragraph({ children: [] })],
    }],
  });
}

export async function exportDocx(json: PMNode, setup?: PageSetup): Promise<ArrayBuffer> {
  const buf = await Packer.toArrayBuffer(buildDocument(json, setup));
  return json.content?.some((n) => n.type === 'tableOfContents') ? fillToc(buf, json) : buf;
}

const xmlEsc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/**
 * Trường mục lục do thư viện tạo ra để trống cho tới khi Word cập nhật. Điền sẵn danh sách tiêu đề làm kết quả của trường
 * để các ứng dụng không tự cập nhật (xem trước trên điện thoại, LibreOffice) vẫn thấy mục lục; Word sẽ cập nhật lại kèm số trang.
 */
async function fillToc(buf: ArrayBuffer, json: PMNode): Promise<ArrayBuffer> {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('word/document.xml')!.async('string');
  const heads: { level: number; text: string }[] = [];
  const walk = (nodes: PMNode[] = []) => nodes.forEach((n) => {
    if (n.type === 'heading' && (n.attrs?.level ?? 1) <= 3) {
      const text = (n.content ?? []).map((c) => c.text ?? '').join('');
      if (text.trim()) heads.push({ level: n.attrs?.level ?? 1, text });
    } else walk(n.content);
  });
  walk(json.content);
  const entries = heads.map((h) => `<w:p><w:pPr><w:pStyle w:val="TOC${h.level}"/><w:ind w:left="${(h.level - 1) * 440}"/></w:pPr><w:r><w:t xml:space="preserve">${xmlEsc(h.text)}</w:t></w:r></w:p>`).join('');
  const marker = '<w:fldChar w:fldCharType="separate"/></w:r></w:p>';
  if (!xml.includes(marker)) return buf;
  zip.file('word/document.xml', xml.replace(marker, marker + entries));
  return zip.generateAsync({ type: 'arraybuffer' });
}
