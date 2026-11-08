import { describe, expect, it } from 'vitest';
import { guessFont, previewFamily, fontFile } from '../src/editors/pdf/fonts';

describe('Chọn phông gần giống phông gốc', () => {
  it('đoán từ tên phông', () => {
    expect(guessFont('BAAAAA+TimesNewRomanPS-BoldItalicMT')).toEqual({ family: 'serif', bold: true, italic: true });
    expect(guessFont('ArialMT')).toEqual({ family: 'sans', bold: false, italic: false });
    expect(guessFont('LiberationSans-Italic')).toMatchObject({ family: 'sans', italic: true });
    expect(guessFont('CourierNewPS-BoldMT')).toMatchObject({ family: 'mono', bold: true });
    expect(guessFont(undefined, { family: 'serif' }).family).toBe('serif');
    expect(fontFile({ family: 'serif', bold: true })).toBe('fonts/DejaVuSerif-Bold.ttf');
  });
  it('phông đơn cách thiếu dấu tiếng Việt thì dùng không chân', () => {
    expect(previewFamily({ family: 'mono' }, 'ABC-123')).toBe('mono');
    expect(previewFamily({ family: 'mono' }, 'Mã số')).toBe('sans');
  });
});
