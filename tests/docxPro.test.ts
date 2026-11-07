import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { Packer } from 'docx';
import { buildDocument, halfPoints } from '../src/editors/docx/docxExport';
import { getSchema } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { Node as PMNode } from '@tiptap/pm/model';
import { findInDoc } from '../src/editors/docx/extensions';

describe('Word nâng cao', () => {
  it('đổi cỡ chữ sang nửa điểm', () => {
    expect(halfPoints('12pt')).toBe(24);
    expect(halfPoints('16px')).toBe(24);
    expect(halfPoints('abc')).toBeUndefined();
  });

  it('xuất phông, cỡ chữ, chỉ số, ngắt trang, khổ ngang và số trang', async () => {
    const doc = buildDocument({
      type: 'doc',
      content: [
        { type: 'paragraph', content: [
          { type: 'text', text: 'Tiêu đề', marks: [{ type: 'textStyle', attrs: { fontFamily: 'Times New Roman', fontSize: '16pt' } }] },
          { type: 'text', text: '2', marks: [{ type: 'superscript' }] },
          { type: 'text', text: 'Trang chủ', marks: [{ type: 'link', attrs: { href: 'https://example.com' } }] },
        ] },
        { type: 'pageBreak' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Trang hai' }] },
      ],
    }, { landscape: true, pageNumbers: true });
    const zip = await JSZip.loadAsync(await Packer.toBuffer(doc));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('w:ascii="Times New Roman"');
    expect(xml).toContain('<w:sz w:val="32"/>');
    expect(xml).toContain('w:vertAlign w:val="superscript"');
    expect(xml).toContain('w:type="page"');
    expect(xml).toContain('w:orient="landscape"');
    const footers = Object.keys(zip.files).filter((n) => /footer\d*\.xml$/.test(n));
    const footer = await zip.file(footers[0])!.async('string');
    expect(footer).toContain('PAGE');
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('https://example.com');
  });

  it('tìm chữ qua nhiều đoạn định dạng, không phân biệt hoa thường', () => {
    const schema = getSchema([StarterKit]);
    const doc = PMNode.fromJSON(schema, {
      type: 'doc',
      content: [
        { type: 'paragraph', content: [{ type: 'text', text: 'Thành phố ' }, { type: 'text', text: 'Hà', marks: [{ type: 'bold' }] }, { type: 'text', text: ' Nội đẹp' }] },
        { type: 'paragraph', content: [{ type: 'text', text: 'hà nội và HÀ NỘI' }] },
      ],
    });
    const found = findInDoc(doc, 'hà nội');
    expect(found.length).toBe(3);
    expect(doc.textBetween(found[0].from, found[0].to)).toBe('Hà Nội');
    expect(doc.textBetween(found[2].from, found[2].to)).toBe('HÀ NỘI');
  });
});
