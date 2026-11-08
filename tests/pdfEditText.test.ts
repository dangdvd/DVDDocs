import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { PDFArray, PDFDocument, PDFName, PDFRawStream, StandardFonts, decodePDFRawStream } from 'pdf-lib';
import { applyAnnotations } from '../src/editors/pdf/pdfOps';

const font = readFileSync(new URL('../public/fonts/DejaVuSans.ttf', import.meta.url)).buffer as ArrayBuffer;

async function contentOf(bytes: Uint8Array) {
  const d = await PDFDocument.load(bytes);
  const c = d.getPage(0).node.get(PDFName.of('Contents'));
  const list = c instanceof PDFArray ? c.asArray() : [c];
  return list.map((r) => Buffer.from(decodePDFRawStream(d.context.lookup(r as never) as PDFRawStream).decode()).toString('latin1')).join('\n');
}

describe('Sửa chữ có sẵn trong PDF', () => {
  it('xoá chữ cũ khỏi nội dung trang và viết chữ mới', async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([600, 800]);
    const f = await doc.embedFont(StandardFonts.Helvetica);
    page.drawText('Gia cu 100.000', { x: 50, y: 700, size: 20, font: f });
    page.drawText('Dong giu nguyen', { x: 50, y: 650, size: 20, font: f });
    const src = await doc.save();
    const hex = (s: string) => Buffer.from(s, 'latin1').toString('hex').toUpperCase();
    expect(await contentOf(src)).toContain(hex('Gia cu 100.000'));

    // Dòng đầu: đường chân chữ ở y = 800 - 700 = 100 (toạ độ hiển thị), cỡ 20.
    const out = await applyAnnotations(src, [
      { id: 'e', page: 0, type: 'edit', x: 50, y: 100 - 20 * 0.95, w: 140, h: 24, text: 'Giá mới 200.000', size: 20, color: '#000000', bg: '#ffffff' },
    ], font);
    const re = await PDFDocument.load(out);
    expect(re.getPageCount()).toBe(1);
    const all = await contentOf(out);
    expect(all).not.toContain(hex('Gia cu 100.000'));
    expect(all).toContain(hex('Dong giu nguyen'));
  });
});
