import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { Engine, emptySheet, key } from '../src/editors/xlsx/sheetModel';
import { cfEvaluator, columnValues, filterHidden } from '../src/editors/xlsx/cf';
import { loadXlsx, saveXlsx } from '../src/editors/xlsx/xlsxIO';

function sample() {
  const s = emptySheet('Bán hàng', 'a');
  const rows: (string | number)[][] = [['Khu vực', 'Doanh số'], ['Bắc', 10], ['Nam', 50], ['Trung', 30], ['Bắc', 70]];
  rows.forEach((row, r) => row.forEach((v, c) => (s.cells[key(r, c)] = { v })));
  return s;
}

describe('Định dạng có điều kiện', () => {
  it('so sánh, chứa chữ, thang màu, thanh dữ liệu', () => {
    const s = sample();
    s.cf = [
      { range: { r0: 1, c0: 1, r1: 4, c1: 1 }, kind: 'gt', a: '40', fill: '#FFC7CE', color: '#9C0006' },
      { range: { r0: 1, c0: 0, r1: 4, c1: 0 }, kind: 'text', a: 'bắc', fill: '#C6EFCE' },
      { range: { r0: 1, c0: 1, r1: 4, c1: 1 }, kind: 'bar', colors: ['#638EC6'] },
    ];
    const book = { sheets: [s] };
    const look = cfEvaluator(s, new Engine(book), 0);
    expect(look(1, 1)?.fill).toBeUndefined();
    expect(look(2, 1)).toMatchObject({ fill: '#FFC7CE', color: '#9C0006' });
    expect(look(4, 1)?.bar?.pct).toBe(1);
    expect(look(1, 1)?.bar?.pct).toBeCloseTo(10 / 70);
    expect(look(1, 0)?.fill).toBe('#C6EFCE');
    expect(look(2, 0)).toBeNull();
    s.cf = [{ range: { r0: 1, c0: 1, r1: 4, c1: 1 }, kind: 'scale', colors: ['#000000', '#ffffff'] }];
    const sc = cfEvaluator(s, new Engine(book), 0);
    expect(sc(1, 1)?.fill).toBe('#000000');
    expect(sc(4, 1)?.fill).toBe('#ffffff');
  });

  it('lưu và mở lại giữ quy tắc, vùng lọc và hàng ẩn do lọc', async () => {
    const s = sample();
    s.cf = [
      { range: { r0: 1, c0: 1, r1: 4, c1: 1 }, kind: 'between', a: '20', b: '60', fill: '#FFEB9C' },
      { range: { r0: 1, c0: 1, r1: 4, c1: 1 }, kind: 'scale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
    ];
    s.filter = { r: 0, c0: 0, c1: 1, r1: 4, hide: { 0: ['Bắc'] } };
    const book = { sheets: [s] };
    const engine = new Engine(book);
    expect(filterHidden(s, engine, 0, s.filter)).toEqual({ 1: true, 4: true });
    expect(columnValues(s, engine, 0, s.filter, 0)).toEqual(['Bắc', 'Nam', 'Trung']);
    const buf = await saveXlsx(book, null, (si, r, c) => engine.evaluate(si, r, c));
    const xml = await (await JSZip.loadAsync(buf)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('<autoFilter ref="A1:B5"');
    expect(xml).toMatch(/<row r="2"[^>]*hidden="1"/);
    expect(xml).toContain('colorScale');
    const back = (await loadXlsx(buf)).book.sheets[0];
    expect(back.filter).toMatchObject({ r: 0, c0: 0, c1: 1, r1: 4 });
    expect(back.cf?.map((x) => x.kind)).toEqual(['between', 'scale']);
    expect(back.cf?.[0]).toMatchObject({ a: '20', b: '60', fill: '#FFEB9C' });
    expect(back.cf?.[1].colors).toHaveLength(3);
    expect(back.filter?.hide).toEqual({ 0: ['Bắc'] });
    expect(back.hiddenRows ?? {}).toEqual({});
    expect(xml).toContain('<filterColumn colId="0"><filters><filter val="Nam"/><filter val="Trung"/></filters></filterColumn>');
  });
});

import { csvToBook, detectDelimiter, parseCsv } from '../src/editors/xlsx/csv';

describe('CSV', () => {
  it('tách ngoặc kép, xuống dòng trong ô, đoán dấu phân cách', () => {
    expect(parseCsv('a,"b,c","d ""x"""\r\n1,"2\n3",4\n')).toEqual([['a', 'b,c', 'd "x"'], ['1', '2\n3', '4']]);
    expect(detectDelimiter('Tên;Tuổi\nAn;20\nBình;30')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2')).toBe('\t');
  });

  it('chuyển thành trang tính: số, số kiểu Việt Nam, giữ số 0 ở đầu', () => {
    const data = new TextEncoder().encode('﻿Tên;Điện thoại;Lương\nAn;0912345678;1.234.567,5\n').buffer as ArrayBuffer;
    const s = csvToBook(data, 'nhanvien.csv', 'x').sheets[0];
    expect(s.name).toBe('nhanvien');
    expect(s.cells[key(0, 0)].v).toBe('Tên');
    expect(s.cells[key(1, 1)].v).toBe('0912345678');
    expect(s.cells[key(1, 2)].v).toBe(1234567.5);
  });
});
