import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { PDFDocument, degrees } from 'pdf-lib';
import { applyAnnotations, deletePage, mergePdf, movePage, rotatePage, toPdfPoint, createBlankPdf } from '../src/editors/pdf/pdfOps';
import { buildDocument, imageSize } from '../src/editors/docx/docxExport';
import { Packer } from 'docx';
import { loadXlsx, saveXlsx } from '../src/editors/xlsx/xlsxIO';
import { Engine, emptySheet, key } from '../src/editors/xlsx/sheetModel';

const font = readFileSync(new URL('../public/fonts/DejaVuSans.ttf', import.meta.url)).buffer as ArrayBuffer;

describe('PDF', () => {
  it('quy đổi toạ độ hiển thị theo góc xoay', async () => {
    const doc = await PDFDocument.create();
    const p = doc.addPage([600, 800]);
    expect(toPdfPoint(p, 10, 20)).toEqual([10, 780]);
    p.setRotation(degrees(90));
    expect(toPdfPoint(p, 10, 20)).toEqual([20, 10]);
    p.setRotation(degrees(180));
    expect(toPdfPoint(p, 10, 20)).toEqual([590, 20]);
    p.setRotation(degrees(270));
    expect(toPdfPoint(p, 10, 20)).toEqual([580, 790]);
  });

  it('ghi chữ tiếng Việt, hình, nét vẽ vào PDF', async () => {
    const blank = await createBlankPdf();
    const out = await applyAnnotations(blank, [
      { id: '1', page: 0, type: 'text', x: 50, y: 50, text: 'Xin chào Việt Nam\nDòng hai', size: 14, color: '#d1242f' },
      { id: '2', page: 0, type: 'highlight', x: 40, y: 40, w: 200, h: 30, color: '#ffd400' },
      { id: '3', page: 0, type: 'whiteout', x: 40, y: 100, w: 100, h: 20, color: '#000000' },
      { id: '4', page: 0, type: 'ink', points: [[10, 10], [20, 30], [40, 35]], color: '#0000ff', width: 2 },
    ], font);
    const re = await PDFDocument.load(out);
    expect(re.getPageCount()).toBe(1);
    expect(out.byteLength).toBeGreaterThan(blank.byteLength + 1000);
  });

  it('thao tác trang: xoay, gộp, di chuyển, xoá', async () => {
    let b = await createBlankPdf();
    b = await mergePdf(b, await createBlankPdf());
    b = await rotatePage(b, 1, 90);
    b = await movePage(b, 1, 0);
    let d = await PDFDocument.load(b);
    expect(d.getPageCount()).toBe(2);
    expect(d.getPage(0).getRotation().angle).toBe(90);
    b = await deletePage(b, 0);
    d = await PDFDocument.load(b);
    expect(d.getPageCount()).toBe(1);
    await expect(deletePage(b, 0)).rejects.toThrow();
  });
});

describe('Word', () => {
  it('xuất docx với tiêu đề, định dạng, danh sách và bảng', async () => {
    const doc = buildDocument({
      type: 'doc',
      content: [
        { type: 'heading', attrs: { level: 1, textAlign: 'center' }, content: [{ type: 'text', text: 'Báo cáo' }] },
        { type: 'paragraph', content: [
          { type: 'text', text: 'Đậm', marks: [{ type: 'bold' }] },
          { type: 'text', text: ' đỏ', marks: [{ type: 'textStyle', attrs: { color: '#ff0000' } }] },
        ] },
        { type: 'orderedList', content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Mục 1' }] }] }] },
        { type: 'table', content: [{ type: 'tableRow', content: [
          { type: 'tableHeader', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Cột A' }] }] },
          { type: 'tableCell', content: [{ type: 'paragraph' }] },
        ] }] },
      ],
    });
    const buf = await Packer.toBuffer(doc);
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('Báo cáo');
    expect(xml).toContain('w:val="center"');
    expect(xml).toContain('FF0000');
    expect(xml).toContain('<w:tbl>');
    expect(xml).toContain('Mục 1');
  });

  it('đọc kích thước ảnh PNG', () => {
    const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAYAAAC56t6BAAAAEklEQVR4nGP4z8DwHwYYYBwAeHgH+drPpx8AAAAASUVORK5CYII='), (c) => c.charCodeAt(0));
    expect(imageSize(png)).toEqual({ w: 2, h: 3 });
  });
});

describe('Excel', () => {
  it('lưu rồi mở lại giữ giá trị, công thức, định dạng và nhiều trang', async () => {
    const s1 = emptySheet('Doanh thu', 'a');
    s1.cells[key(0, 0)] = { v: 'Tháng', s: { bold: true, fill: '#ffeeaa' } };
    s1.cells[key(1, 0)] = { v: 1500000, s: { numFmt: '#,##0 "₫"', align: 'right' } };
    s1.cells[key(2, 0)] = { v: 2500000 };
    s1.cells[key(3, 0)] = { f: 'SUM(A2:A3)' };
    s1.colWidths[0] = 150;
    const s2 = emptySheet('Ghi chú', 'b');
    s2.cells[key(0, 0)] = { v: 'Xin chào' };
    const book = { sheets: [s1, s2] };
    const e = new Engine(book);
    const buf = await saveXlsx(book, null, (si, r, c) => e.evaluate(si, r, c));
    const re = await loadXlsx(buf);
    expect(re.book.sheets.map((s) => s.name)).toEqual(['Doanh thu', 'Ghi chú']);
    const c = re.book.sheets[0].cells;
    expect(c[key(0, 0)]).toMatchObject({ v: 'Tháng', s: { bold: true, fill: '#FFEEAA' } });
    expect(c[key(1, 0)]).toMatchObject({ v: 1500000, s: { numFmt: '#,##0 "₫"', align: 'right' } });
    expect(c[key(3, 0)].f).toBe('SUM(A2:A3)');
    expect(re.book.sheets[0].colWidths[0]).toBe(150);

    // Sửa tiếp trên workbook đã mở: xoá một ô, xoá một trang, ghi lại.
    delete re.book.sheets[0].cells[key(2, 0)];
    re.book.sheets.pop();
    const e2 = new Engine(re.book);
    const buf2 = await saveXlsx(re.book, re, (si, r, cc) => e2.evaluate(si, r, cc));
    const re2 = await loadXlsx(buf2);
    expect(re2.book.sheets.length).toBe(1);
    expect(re2.book.sheets[0].cells[key(2, 0)]).toBeUndefined();
  });
});
