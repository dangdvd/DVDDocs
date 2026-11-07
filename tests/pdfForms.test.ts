import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { PDFDocument } from 'pdf-lib';
import { applyAnnotations, extractPages, parsePageRanges, createBlankPdf, mergePdf } from '../src/editors/pdf/pdfOps';

const font = readFileSync(new URL('../public/fonts/DejaVuSans.ttf', import.meta.url)).buffer as ArrayBuffer;

/** PDF mẫu có đủ các loại trường biểu mẫu. */
export async function makeFormPdf() {
  const doc = await PDFDocument.create();
  const page = doc.addPage([595, 842]);
  const form = doc.getForm();
  form.createTextField('ho_ten').addToPage(page, { x: 100, y: 700, width: 250, height: 22 });
  form.createCheckBox('dong_y').addToPage(page, { x: 100, y: 650, width: 16, height: 16 });
  const g = form.createRadioGroup('gioi_tinh');
  g.addOptionToPage('Nam', page, { x: 100, y: 600, width: 14, height: 14 });
  g.addOptionToPage('Nu', page, { x: 160, y: 600, width: 14, height: 14 });
  const dd = form.createDropdown('tinh');
  dd.addOptions(['Hà Nội', 'Đà Nẵng', 'TP. Hồ Chí Minh']);
  dd.addToPage(page, { x: 100, y: 550, width: 150, height: 22 });
  return doc.save();
}

describe('PDF: biểu mẫu và trích trang', () => {
  it('điền biểu mẫu với chữ tiếng Việt', async () => {
    const out = await applyAnnotations(await makeFormPdf(), [], font, {
      ho_ten: 'Nguyễn Văn Dũng', dong_y: true, gioi_tinh: 'Nu', tinh: 'Đà Nẵng', khong_ton_tai: 'x',
    });
    const form = (await PDFDocument.load(out)).getForm();
    expect(form.getTextField('ho_ten').getText()).toBe('Nguyễn Văn Dũng');
    expect(form.getCheckBox('dong_y').isChecked()).toBe(true);
    expect(form.getRadioGroup('gioi_tinh').getSelected()).toBe('Nu');
    expect(form.getDropdown('tinh').getSelected()).toEqual(['Đà Nẵng']);

    // Giá trị nút chọn theo tên trạng thái như pdf.js báo ("0", "1").
    const out2 = await applyAnnotations(out, [], font, { gioi_tinh: '0' });
    expect((await PDFDocument.load(out2)).getForm().getRadioGroup('gioi_tinh').getSelected()).toBe('Nam');
  });

  it('đọc dải trang', () => {
    expect(parsePageRanges('1-3, 5', 6)).toEqual([0, 1, 2, 4]);
    expect(parsePageRanges('4-', 6)).toEqual([3, 4, 5]);
    expect(parsePageRanges('2 2 1', 3)).toEqual([1, 0]);
    expect(() => parsePageRanges('7', 6)).toThrow();
    expect(() => parsePageRanges('abc', 6)).toThrow();
  });

  it('trích trang ra tệp mới', async () => {
    let b = await createBlankPdf();
    b = await mergePdf(b, await createBlankPdf());
    b = await mergePdf(b, await createBlankPdf());
    const out = await extractPages(b, [2, 0]);
    expect((await PDFDocument.load(out)).getPageCount()).toBe(2);
  });
});
