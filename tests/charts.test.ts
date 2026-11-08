import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { Engine, emptySheet, key } from '../src/editors/xlsx/sheetModel';
import { absRef, chartData, chartFromRange, parseRef } from '../src/editors/xlsx/charts';
import { loadXlsx, saveXlsx } from '../src/editors/xlsx/xlsxIO';

function sample() {
  const s = emptySheet('Doanh thu', 'a');
  const rows: (string | number)[][] = [['Tháng', 'Hà Nội', 'Đà Nẵng'], ['T1', 10, 5], ['T2', 20, 7], ['T3', 15, 9]];
  rows.forEach((row, r) => row.forEach((v, c) => (s.cells[key(r, c)] = { v })));
  return s;
}

describe('Biểu đồ', () => {
  it('tham chiếu có tên trang', () => {
    expect(absRef('Doanh thu', 1, 0, 3, 0)).toBe("'Doanh thu'!$A$2:$A$4");
    expect(absRef('Sheet1', 0, 1)).toBe('Sheet1!$B$1');
    expect(parseRef("'Doanh thu'!$A$2:$A$4")).toEqual({ sheet: 'Doanh thu', m: { r0: 1, c0: 0, r1: 3, c1: 0 } });
  });

  it('tạo từ vùng chọn: hàng đầu là tên chuỗi, cột đầu là nhãn', () => {
    const s = sample();
    const book = { sheets: [s] };
    const e = new Engine(book);
    const ch = chartFromRange(s, e, 0, { r0: 0, c0: 0, r1: 3, c1: 2 }, 'col', 'x')!;
    expect(ch.series).toHaveLength(2);
    expect(ch.series[0]).toEqual({ name: "'Doanh thu'!$B$1", cat: "'Doanh thu'!$A$2:$A$4", val: "'Doanh thu'!$B$2:$B$4" });
    const d = chartData(ch, book, e, 0);
    expect(d.cats).toEqual(['T1', 'T2', 'T3']);
    expect(d.series[1]).toEqual({ name: 'Đà Nẵng', values: [5, 7, 9] });
  });

  it('lưu vào .xlsx rồi mở lại vẫn còn biểu đồ', async () => {
    const s = sample();
    const book = { sheets: [s, emptySheet('Khác', 'b')] };
    const e = new Engine(book);
    s.charts = [
      { ...chartFromRange(s, e, 0, { r0: 0, c0: 0, r1: 3, c1: 2 }, 'line', 'c1')!, title: 'Doanh số', anchor: { r0: 1, c0: 4, r1: 16, c1: 10 } },
      { ...chartFromRange(s, e, 0, { r0: 0, c0: 0, r1: 3, c1: 1 }, 'pie', 'c2')!, anchor: { r0: 18, c0: 4, r1: 30, c1: 10 } },
    ];
    const buf = await saveXlsx(book, null, (si, r, c) => e.evaluate(si, r, c));
    const zip = await JSZip.loadAsync(buf);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('/xl/charts/chart1.xml');
    expect(ct).toContain('/xl/drawings/drawing1.xml');
    const c1 = await zip.file('xl/charts/chart1.xml')!.async('string');
    expect(c1).toContain('<c:lineChart>');
    expect(c1).toContain('<c:v>Hà Nội</c:v>');
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheetXml).toMatch(/<drawing r:id="rId\d+"\/><\/worksheet>|<drawing r:id="rId\d+"\/><(tableParts|extLst|legacyDrawing)/);

    const re = await loadXlsx(buf);
    const charts = re.book.sheets[0].charts!;
    expect(charts.map((c) => c.kind)).toEqual(['line', 'pie']);
    expect(charts[0].title).toBe('Doanh số');
    expect(charts[0].anchor).toEqual({ r0: 1, c0: 4, r1: 16, c1: 10 });
    expect(charts[0].series[1].val).toBe("'Doanh thu'!$C$2:$C$4");
    expect(re.book.sheets[1].charts).toBeUndefined();

    // Lưu lại lần nữa (workbook đã mở): không bị nhân đôi biểu đồ.
    const e2 = new Engine(re.book);
    const buf2 = await saveXlsx(re.book, re, (si, r, c) => e2.evaluate(si, r, c));
    const re2 = await loadXlsx(buf2);
    expect(re2.book.sheets[0].charts).toHaveLength(2);
  });
});
