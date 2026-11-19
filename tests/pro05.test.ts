import { describe, expect, it } from 'vitest';
import { readOutline, type OutlineSource } from '../src/editors/pdf/outline';
import { printArea, sheetToHtml } from '../src/editors/xlsx/printSheet';
import { Engine, emptySheet, key, type Sheet } from '../src/editors/xlsx/sheetModel';
import { escapeHtml } from '../src/lib/print';

describe('Mục lục PDF', () => {
  const refs: Record<string, number> = { '10 0': 0, '20 0': 2, '30 0': 4 };
  const src = (outline: unknown): OutlineSource => ({
    getOutline: async () => outline as never,
    getDestination: async (id: string) => (id === 'chuong2' ? [{ num: 20, gen: 0 }, { name: 'XYZ' }] : null),
    getPageIndex: async (r) => {
      const v = refs[`${r.num} ${r.gen}`];
      if (v === undefined) throw new Error('bad ref');
      return v;
    },
  });

  it('làm phẳng cây mục lục, tìm trang đích qua tham chiếu, tên đích và số trang', async () => {
    const items = await readOutline(src([
      { title: 'Chương 1', dest: [{ num: 10, gen: 0 }, { name: 'Fit' }], items: [
        { title: '  Mục 1.1 ', dest: [3, { name: 'Fit' }] },
      ] },
      { title: 'Chương 2', dest: 'chuong2' },
      { title: '', dest: [{ num: 99, gen: 0 }] },
      { title: 'Liên kết web', dest: null },
    ]));
    expect(items).toEqual([
      { title: 'Chương 1', page: 0, depth: 0 },
      { title: 'Mục 1.1', page: 3, depth: 1 },
      { title: 'Chương 2', page: 2, depth: 0 },
      { title: '(không tên)', page: null, depth: 0 },
      { title: 'Liên kết web', page: null, depth: 0 },
    ]);
  });

  it('tài liệu không có mục lục', async () => {
    expect(await readOutline(src(null))).toEqual([]);
    expect(await readOutline({ ...src(null), getOutline: () => Promise.reject(new Error('x')) })).toEqual([]);
  });
});

describe('In bảng tính', () => {
  const make = (): Sheet => {
    const s = emptySheet('Doanh thu', 's1');
    s.cells[key(0, 0)] = { v: 'Tên', s: { bold: true, fill: '#ffff00' } };
    s.cells[key(0, 1)] = { v: 'Tiền' };
    s.cells[key(1, 0)] = { v: '<An & Bình>' };
    s.cells[key(1, 1)] = { v: 1234.5, s: { numFmt: '#,##0.00', border: { t: true, b: true } } };
    s.cells[key(2, 1)] = { f: 'B2*2' };
    s.cells[key(4, 0)] = { v: 'Gộp' };
    s.merges = [{ r0: 4, c0: 0, r1: 4, c1: 2 }];
    return s;
  };

  it('vùng in là vùng có dữ liệu (kể cả ô gộp) hoặc vùng chọn đã cắt phần trống', () => {
    const s = make();
    expect(printArea(s)).toEqual({ r0: 0, c0: 0, r1: 4, c1: 2 });
    expect(printArea(s, { r0: 0, c0: 0, r1: 0, c1: 0 })).toEqual({ r0: 0, c0: 0, r1: 4, c1: 2 });
    expect(printArea(s, { r0: 1, c0: 0, r1: 999, c1: 50 })).toEqual({ r0: 1, c0: 0, r1: 4, c1: 2 });
    expect(printArea(emptySheet('x', 'x'))).toBeNull();
  });

  it('dựng bảng HTML: giá trị đã tính, định dạng số, kiểu, gộp ô, hàng ẩn, thoát ký tự HTML', () => {
    const s = make();
    const book = { sheets: [s] };
    const html = sheetToHtml(s, new Engine(book), 0, printArea(s)!, { defaultWidth: 96, hidden: { 3: true }, title: 'Doanh thu' });
    expect(html).toContain('print-sheet-title">Doanh thu<');
    expect(html).toContain('&lt;An &amp; Bình&gt;');
    expect(html).not.toContain('<An');
    expect(html).toContain(">2469</td>");
    expect(html).toContain(">1,234.50</td>");
    expect(html).toContain('font-weight:700');
    expect(html).toContain('background:#ffff00');
    expect(html).toContain('border-top:1px solid #000');
    expect(html).toContain('colspan="3"');
    expect((html.match(/<tr>/g) ?? []).length).toBe(4); // hàng 4 bị ẩn
    expect((html.match(/<col /g) ?? []).length).toBe(3);
  });

  it('định dạng có điều kiện được áp vào bản in', () => {
    const s = make();
    const html = sheetToHtml(s, new Engine({ sheets: [s] }), 0, { r0: 1, c0: 1, r1: 1, c1: 1 }, {
      defaultWidth: 96, look: (r, c) => (r === 1 && c === 1 ? { fill: '#00ff00', color: '#ff0000' } : null),
    });
    expect(html).toContain('background:#00ff00');
    expect(html).toContain('color:#ff0000');
  });

  it('escapeHtml', () => {
    expect(escapeHtml(`a<b>"c"&'d'`)).toBe('a&lt;b&gt;&quot;c&quot;&amp;&#39;d&#39;');
  });
});
