// Chuyển nội dung trình soạn thảo (JSON của TipTap/ProseMirror) sang tệp .docx thật bằng thư viện `docx`.
import {
  AlignmentType, BorderStyle, Document, ExternalHyperlink, HeadingLevel, ImageRun, LevelFormat, Packer,
  Paragraph, ParagraphChild, ShadingType, Table, TableCell, TableRow, TextRun, WidthType,
} from 'docx';

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
        const color = toHex(has('textStyle')?.attrs?.color);
        const hl = has('highlight');
        const link = has('link');
        const run = new TextRun({
          text: n.text,
          bold: !!has('bold'),
          italics: !!has('italic'),
          underline: has('underline') || link ? {} : undefined,
          strike: !!has('strike'),
          color: color ?? (link ? '0563C1' : undefined),
          font: has('code') ? 'Courier New' : undefined,
          shading: hl ? { type: ShadingType.CLEAR, color: 'auto', fill: toHex(hl.attrs?.color) ?? 'FFFF00' } : undefined,
        });
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
          shading: cell.type === 'tableHeader' ? { type: ShadingType.CLEAR, color: 'auto', fill: 'EDEFF2' } : undefined,
        });
      }),
    }));
    return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
  }
}

export function buildDocument(json: PMNode): Document {
  const conv = new Converter();
  const children = conv.blocks(json.content);
  const fmts = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN];
  return new Document({
    creator: 'DVDDocs',
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
    sections: [{ children: children.length ? children : [new Paragraph({ children: [] })] }],
  });
}

export async function exportDocx(json: PMNode): Promise<ArrayBuffer> {
  return Packer.toArrayBuffer(buildDocument(json));
}
