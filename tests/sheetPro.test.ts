import { describe, expect, it } from 'vitest';
import {
  Engine, emptySheet, findMatches, key, parseRange, rangeRef, replaceInCell, shiftMerges, shiftSpan,
} from '../src/editors/xlsx/sheetModel';
import { loadXlsx, saveXlsx } from '../src/editors/xlsx/xlsxIO';

describe('ô gộp', () => {
  it('đọc/ghi địa chỉ vùng', () => {
    expect(parseRange('B2:C4')).toEqual({ r0: 1, c0: 1, r1: 3, c1: 2 });
    expect(parseRange('$A$1')).toEqual({ r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(rangeRef({ r0: 1, c0: 1, r1: 3, c1: 2 })).toBe('B2:C4');
  });

  it('dịch đoạn khi chèn/xoá', () => {
    expect(shiftSpan(2, 4, 3, 2)).toEqual([2, 6]); // chèn giữa: giãn
    expect(shiftSpan(2, 4, 0, 1)).toEqual([3, 5]); // chèn phía trên: dời
    expect(shiftSpan(2, 4, 3, -1)).toEqual([2, 3]); // xoá giữa: co
    expect(shiftSpan(2, 4, 2, -3)).toBeNull(); // xoá hết
    expect(shiftSpan(5, 6, 0, -2)).toEqual([3, 4]); // xoá phía trên: dời lên
    expect(shiftMerges([{ r0: 2, c0: 0, r1: 3, c1: 0 }], 'row', 3, -1)).toEqual([]); // còn 1 ô: bỏ gộp
  });
});

describe('tìm & thay', () => {
  it('tìm không phân biệt hoa thường, theo cả giá trị hiển thị', () => {
    const s = emptySheet('S', 'a');
    s.cells[key(0, 0)] = { v: 'Hà Nội' };
    s.cells[key(1, 1)] = { v: 'hà nội mới' };
    s.cells[key(2, 0)] = { v: 1500, s: { numFmt: '#,##0' } };
    s.cells[key(3, 0)] = { f: 'A3*2' };
    const e = new Engine({ sheets: [s] });
    expect(findMatches(s, e, 0, 'HÀ NỘI')).toEqual([[0, 0], [1, 1]]);
    expect(findMatches(s, e, 0, '3000')).toEqual([[3, 0]]);
    expect(replaceInCell(s.cells[key(1, 1)], 'Hà Nội', 'Huế')).toEqual({ v: 'Huế mới', s: undefined });
    expect(replaceInCell({ v: 'abc' }, 'x', 'y')).toBeNull();
    expect(replaceInCell({ v: 'Giá 100' }, 'Giá ', '')).toEqual({ v: 100, s: undefined });
  });
});

describe('Excel nâng cao', () => {
  it('lưu rồi mở lại giữ gộp ô, cố định, viền, xuống dòng, chiều cao hàng', async () => {
    const s = emptySheet('Báo cáo', 'a');
    s.cells[key(0, 0)] = { v: 'BÁO CÁO DOANH THU', s: { bold: true, align: 'center' } };
    s.cells[key(1, 0)] = { v: 'Ghi chú dài', s: { wrap: true, border: { t: true, r: true, b: true, l: true } } };
    s.cells[key(2, 1)] = { v: 5, s: { border: { b: true } } };
    s.merges = [{ r0: 0, c0: 0, r1: 0, c1: 3 }];
    s.freeze = { rows: 1, cols: 1 };
    s.rowHeights = { 1: 40 };
    const book = { sheets: [s] };
    const e = new Engine(book);
    const buf = await saveXlsx(book, null, (si, r, c) => e.evaluate(si, r, c));
    const re = await loadXlsx(buf);
    const sh = re.book.sheets[0];
    expect(sh.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 3 }]);
    expect(sh.freeze).toEqual({ rows: 1, cols: 1 });
    expect(sh.rowHeights?.[1]).toBe(40);
    expect(sh.cells[key(1, 0)].s).toMatchObject({ wrap: true, border: { t: true, r: true, b: true, l: true } });
    expect(sh.cells[key(2, 1)].s?.border).toEqual({ t: false, r: false, b: true, l: false });

    // Bỏ gộp, bỏ cố định, bỏ viền rồi lưu đè lên workbook đã mở.
    sh.merges = [];
    sh.freeze = undefined;
    sh.cells[key(2, 1)] = { v: 5 };
    const e2 = new Engine(re.book);
    const re2 = await loadXlsx(await saveXlsx(re.book, re, (si, r, c) => e2.evaluate(si, r, c)));
    const sh2 = re2.book.sheets[0];
    expect(sh2.merges).toEqual([]);
    expect(sh2.freeze).toBeUndefined();
    expect(sh2.cells[key(2, 1)].s?.border).toBeUndefined();
    expect(sh2.cells[key(0, 0)].v).toBe('BÁO CÁO DOANH THU');
  });
});
