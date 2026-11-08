import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { Packer } from 'docx';
import { formattingToHtml, headerFooterText, markFormatting } from '../src/editors/docx/docxImport';
import { buildDocument } from '../src/editors/docx/docxExport';

describe('Đọc định dạng Word mà mammoth bỏ qua', () => {
  it('màu chữ, tô nền, căn lề, giãn dòng, màu ô', () => {
    const xml = '<w:body><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>' +
      '<w:r><w:rPr><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr><w:t>Đỏ</w:t></w:r><w:r><w:t xml:space="preserve"> thường</w:t></w:r></w:p>' +
      '<w:tbl><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>Ô</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>';
    const marked = markFormatting(xml);
    // Mô phỏng HTML do mammoth tạo từ chữ đã đánh dấu.
    const text = [...marked.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]);
    const html = formattingToHtml(`<p>${text[0]}${text[1]}</p><table><tr><td><p>${text[2]}</p></td></tr></table>`);
    expect(html).toBe('<p style="text-align: center; line-height: 1.5"><span style="color: #C00000"><mark data-color="#FFFF00" style="background-color: #FFFF00">Đỏ</mark></span> thường</p>' +
      '<table><tr><td style="background-color: #FFF2CC"><p>Ô</p></td></tr></table>');
  });

  it('chữ đầu/chân trang bỏ đoạn số trang', () => {
    const xml = '<w:ftr><w:p><w:r><w:t>Công ty ABC</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Trang </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>';
    expect(headerFooterText(xml)).toBe('Công ty ABC');
  });

  it('xuất giãn dòng, màu ô, đầu/chân trang, mục lục', async () => {
    const doc = buildDocument({
      type: 'doc', content: [
        { type: 'tableOfContents' },
        { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Phần 1' }] },
        { type: 'paragraph', attrs: { lineHeight: '1.5' }, content: [{ type: 'text', text: 'A' }] },
        { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', attrs: { backgroundColor: '#fff2cc' }, content: [{ type: 'paragraph' }] }] }] },
      ],
    }, { landscape: false, pageNumbers: true, header: 'Đầu trang ABC', footer: 'Chân trang XYZ' });
    const zip = await JSZip.loadAsync(await Packer.toBuffer(doc));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toMatch(/w:line="360"/);
    expect(xml).toContain('w:fill="FFF2CC"');
    expect(xml).toContain('TOC \\h \\o');
    const parts = Object.keys(zip.files);
    const header = await zip.file(parts.find((p) => /word\/header\d*\.xml/.test(p))!)!.async('string');
    const footer = await zip.file(parts.find((p) => /word\/footer\d*\.xml/.test(p))!)!.async('string');
    expect(header).toContain('Đầu trang ABC');
    expect(footer).toContain('Chân trang XYZ');
    expect(footer).toContain('PAGE');
    expect(await zip.file('word/settings.xml')!.async('string')).toContain('updateFields');
  });
});
