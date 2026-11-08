// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { TextSelection } from '@tiptap/pm/state';
import JSZip from 'jszip';
import { Packer } from 'docx';
import { TrackChanges, listChanges, markTrackedRuns, trackedMarkersToHtml } from '../src/editors/docx/trackChanges';
import { buildDocument, PMNode } from '../src/editors/docx/docxExport';

const mk = (html = '<p>Xin chào thế giới</p>') => {
  const ed = new Editor({ extensions: [StarterKit, TrackChanges], content: html });
  ed.commands.setTrackChanges(true);
  return ed;
};
const sel = (ed: Editor, a: number, b = a) => ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, a, b)));
const html = (ed: Editor) => ed.getHTML().replace(/ (data-author|data-date|title)="[^"]*"/g, '');

describe('Theo dõi thay đổi', () => {
  it('gõ thêm được đánh dấu chèn', () => {
    const ed = mk();
    sel(ed, 4);
    ed.commands.insertContent('X');
    ed.commands.insertContent('Y');
    expect(html(ed)).toBe('<p>Xin<ins data-tc="" class="tc-ins">XY</ins> chào thế giới</p>');
  });

  it('xoá chữ gốc thì giữ lại và gạch ngang; Backspace đưa con trỏ về trước', () => {
    const ed = mk();
    sel(ed, 4);
    ed.view.dispatch(ed.state.tr.delete(3, 4)); // Backspace xoá "n"
    expect(html(ed)).toBe('<p>Xi<del data-tc="" class="tc-del">n</del> chào thế giới</p>');
    expect(ed.state.selection.from).toBe(3);
  });

  it('gõ đè vùng chọn: phần cũ bị gạch, phần mới được chèn; xoá chữ mình vừa chèn thì xoá hẳn', () => {
    const ed = mk();
    sel(ed, 5, 9); // "chào"
    ed.commands.insertContent('hello');
    expect(html(ed)).toBe('<p>Xin <del data-tc="" class="tc-del">chào</del><ins data-tc="" class="tc-ins">hello</ins> thế giới</p>');
    const end = ed.state.selection.from;
    ed.view.dispatch(ed.state.tr.delete(end - 2, end));
    expect(html(ed)).toContain('<ins data-tc="" class="tc-ins">hel</ins>');
    expect(listChanges(ed.state.doc).map((c) => c.kind)).toEqual(['deletion', 'insertion']);
  });

  it('Backspace liên tiếp bỏ qua phần đã xoá, xoá trọn ký tự có dấu tổ hợp', () => {
    const ed = mk('<p>cùng ky\u0300.</p>');
    sel(ed, ed.state.doc.content.size - 1);
    for (let i = 0; i < 3; i++) ed.commands.keyboardShortcut('Backspace');
    expect(html(ed)).toBe('<p>cùng <del data-tc="" class="tc-del">ky\u0300.</del></p>');
    expect(ed.state.selection.from).toBe(6);
    ed.commands.keyboardShortcut('Delete'); // nhảy qua phần đã xoá rồi hết đoạn: không làm gì thêm
    expect(listChanges(ed.state.doc)).toHaveLength(1);
  });

  it('chấp nhận và từ chối', () => {
    const ed = mk();
    sel(ed, 5, 9);
    ed.commands.insertContent('hello');
    const copy = ed.getJSON();
    ed.commands.acceptAllChanges();
    expect(html(ed)).toBe('<p>Xin hello thế giới</p>');
    ed.commands.setContent(copy);
    ed.commands.rejectAllChanges();
    expect(html(ed)).toBe('<p>Xin chào thế giới</p>');
    // Hoàn tác đưa về trạng thái trước khi gõ (một bước).
  });

  it('tắt theo dõi: gõ bình thường', () => {
    const ed = mk();
    ed.commands.setTrackChanges(false);
    sel(ed, 4);
    ed.view.dispatch(ed.state.tr.delete(3, 4));
    expect(html(ed)).toBe('<p>Xi chào thế giới</p>');
  });

  it('xuất .docx với w:ins và w:del', async () => {
    const ed = mk();
    sel(ed, 5, 9);
    ed.commands.insertContent('hello');
    const doc = buildDocument(ed.getJSON() as PMNode, { landscape: false, pageNumbers: false, track: true });
    const zip = await JSZip.loadAsync(await Packer.toBuffer(doc));
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:ins [^>]*w:author="[^"]+"[^>]*>.*hello/);
    expect(xml).toMatch(/<w:del [^>]*>.*<w:delText[^>]*>chào<\/w:delText>/);
    expect(await zip.file('word/settings.xml')!.async('string')).toContain('w:trackRevisions');
  });

  it('đọc lại w:ins/w:del từ document.xml', () => {
    const xml = '<w:p><w:r><w:t>A</w:t></w:r><w:ins w:id="1" w:author="B"><w:r><w:t xml:space="preserve">mới</w:t></w:r></w:ins>' +
      '<w:del w:id="2" w:author="B"><w:r><w:delText>cũ</w:delText></w:r></w:del><w:r><w:rPr><w:del w:id="3"/></w:rPr></w:r></w:p>';
    const out = markTrackedRuns(xml);
    expect(out).toContain('<w:t xml:space="preserve">mới</w:t>');
    expect(out).toContain('<w:t>cũ</w:t>');
    expect(out).not.toContain('<w:ins w:id="1"');
    expect(out).toContain('<w:del w:id="3"/>');
    expect(trackedMarkersToHtml('<p>ab</p>')).toBe('<p><ins data-tc="">a</ins><del data-tc="">b</del></p>');
  });
});
