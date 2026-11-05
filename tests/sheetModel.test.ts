import { describe, expect, it } from 'vitest';
import { Engine, colIndex, colName, emptySheet, formatValue, key, moveFormula, parseInput, shiftFormula } from '../src/editors/xlsx/sheetModel';

describe('tên cột', () => {
  it('chuyển đổi hai chiều', () => {
    for (const [i, n] of [[0, 'A'], [25, 'Z'], [26, 'AA'], [701, 'ZZ'], [702, 'AAA']] as const) {
      expect(colName(i)).toBe(n);
      expect(colIndex(n)).toBe(i);
    }
  });
});

describe('parseInput', () => {
  it('nhận dạng số, phần trăm, công thức, chữ', () => {
    expect(parseInput('42')).toEqual({ v: 42 });
    expect(parseInput('1,234.5')).toEqual({ v: 1234.5 });
    expect(parseInput('15%')).toEqual({ v: 0.15 });
    expect(parseInput('=A1+1')).toEqual({ f: 'A1+1' });
    expect(parseInput('Xin chào')).toEqual({ v: 'Xin chào' });
    expect(parseInput('')).toEqual({ v: null });
  });
});

describe('formatValue', () => {
  it('định dạng số', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatValue(0.256, '0.0%')).toBe('25.6%');
    expect(formatValue(1500000, '#,##0 "₫"')).toBe('1,500,000 ₫');
    expect(formatValue(45658, 'dd/mm/yyyy')).toBe('01/01/2025');
    expect(formatValue(0.1 + 0.2)).toBe('0.3');
  });
});

describe('dịch công thức', () => {
  it('chèn hàng dịch tham chiếu phía dưới', () => {
    expect(shiftFormula('SUM(A1:A10)+B5', 'S', 'S', 'row', 3, 1)).toBe('SUM(A1:A11)+B6');
    expect(shiftFormula('$A$5*2', 'S', 'S', 'row', 0, 2)).toBe('$A$7*2');
  });
  it('xoá cột biến tham chiếu bị xoá thành #REF!', () => {
    expect(shiftFormula('B1+C1', 'S', 'S', 'col', 1, -1)).toBe('#REF!+B1');
  });
  it('không đụng tới trang tính khác hoặc chuỗi', () => {
    expect(shiftFormula('Other!A5+"A5"', 'S', 'S', 'row', 0, 1)).toBe('Other!A5+"A5"');
  });
  it('sao chép công thức giữ $ tuyệt đối', () => {
    expect(moveFormula('A1*$B$1+$C2', 2, 1)).toBe('B3*$B$1+$C4');
  });
});

describe('Engine', () => {
  it('tính công thức lồng nhau, liên trang và báo lỗi', () => {
    const s1 = emptySheet('Sheet1', '1');
    const s2 = emptySheet('Dữ liệu', '2');
    s1.cells[key(0, 0)] = { v: 10 };
    s1.cells[key(1, 0)] = { v: 20 };
    s1.cells[key(2, 0)] = { f: 'SUM(A1:A2)' };
    s1.cells[key(3, 0)] = { f: 'A3*2' };
    s1.cells[key(4, 0)] = { f: "'Dữ liệu'!B2+1" };
    s1.cells[key(5, 0)] = { f: '1/0' };
    s1.cells[key(6, 0)] = { f: 'A7' };
    s1.cells[key(7, 0)] = { f: 'IF(A1>5,"lớn","nhỏ")' };
    s1.cells[key(8, 0)] = { f: 'AVERAGE(A:A)' };
    s2.cells[key(1, 1)] = { v: 99 };
    const e = new Engine({ sheets: [s1, s2] });
    expect(e.evaluate(0, 2, 0)).toBe(30);
    expect(e.evaluate(0, 3, 0)).toBe(60);
    expect(e.evaluate(0, 4, 0)).toBe(100);
    expect(e.evaluate(0, 5, 0)).toEqual({ error: '#DIV/0!' });
    expect(e.evaluate(0, 6, 0)).toEqual({ error: '#CIRC!' });
    expect(e.evaluate(0, 7, 0)).toBe('lớn');
  });
});
