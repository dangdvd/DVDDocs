import { useEffect, useMemo, useRef, useState } from 'react';
import { EditorContent, useEditor, Editor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import Underline from '@tiptap/extension-underline';
import TextAlign from '@tiptap/extension-text-align';
import Table from '@tiptap/extension-table';
import TableRow from '@tiptap/extension-table-row';
import TableCell from '@tiptap/extension-table-cell';
import TableHeader from '@tiptap/extension-table-header';
import Image from '@tiptap/extension-image';
import Highlight from '@tiptap/extension-highlight';
import TextStyle from '@tiptap/extension-text-style';
import Color from '@tiptap/extension-color';
import FontFamily from '@tiptap/extension-font-family';
import Link from '@tiptap/extension-link';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import JSZip from 'jszip';
import { OpenedFile, saveFile } from '../../lib/fileio';
import { EditorShell, Sep, ToolButton } from '../../components/EditorShell';
import { exportDocx, PageSetup, PMNode } from './docxExport';
import { FontSize, LineHeight, PageBreak, SearchHighlight, TableOfContents, cellBackground, findInDoc, searchKey } from './extensions';
import { formattingToHtml, headerFooterText, markFormatting } from './docxImport';
import { TrackChanges, currentAuthor, listChanges, markTrackedRuns, trackKey, trackedMarkersToHtml } from './trackChanges';

const FONTS = ['Arial', 'Calibri', 'Cambria', 'Times New Roman', 'Tahoma', 'Verdana', 'Georgia', 'Courier New'];
const SIZES = [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

interface MammothRun { font?: string; fontSize?: number; children: { type: string; value?: string }[] }
interface MammothTransforms { run(fn: (r: MammothRun) => MammothRun): unknown }

async function docxToHtml(data: ArrayBuffer): Promise<{ html: string; warnings: string[] }> {
  const mammoth = await import('mammoth');
  // Giữ các thay đổi được theo dõi (chèn/xoá) của tệp gốc.
  try {
    const zip = await JSZip.loadAsync(data);
    const xml = await zip.file('word/document.xml')?.async('string');
    if (xml) {
      // Giữ màu chữ, tô nền, căn lề, giãn dòng, màu ô và các thay đổi được theo dõi.
      zip.file('word/document.xml', markTrackedRuns(markFormatting(xml)));
      data = await zip.generateAsync({ type: 'arraybuffer' });
    }
  } catch { /* tệp lạ: để mammoth tự xử lý */ }
  const transforms = ((mammoth as unknown as { default?: { transforms?: MammothTransforms } }).default?.transforms
    ?? (mammoth as unknown as { transforms: MammothTransforms }).transforms);
  const res = await mammoth.convertToHtml(
    { arrayBuffer: data },
    {
      styleMap: [
        'u => u', 'strike => s', "p[style-name='Title'] => h1:fresh", "p[style-name='Subtitle'] => h2:fresh",
        "br[type='page'] => hr.page-break",
      ],
      // mammoth bỏ phông và cỡ chữ của từng đoạn chữ: đánh dấu chúng bằng ký tự riêng rồi đổi thành <span style> bên dưới.
      transformDocument: transforms.run((run: MammothRun) => {
        const st: string[] = [];
        if (run.font) st.push(`font-family: ${run.font.replace(/[^\p{L}\p{N} \-]/gu, '')}`);
        if (run.fontSize) st.push(`font-size: ${run.fontSize}pt`);
        if (!st.length || run.children.some((c) => c.type !== 'text' && c.type !== 'tab')) return run;
        return { ...run, children: [{ type: 'text', value: `\uE000${st.join('; ')}\uE001` }, ...run.children, { type: 'text', value: '\uE002' }] };
      }),
    } as Parameters<typeof mammoth.convertToHtml>[1],
  );
  // mammoth đặt ngắt trang trong một đoạn riêng: đưa ra ngoài để thành khối ngắt trang.
  let html = res.value
    .replace(/<p>\s*<hr class="page-break" \/>\s*<\/p>/g, '<hr class="page-break" />')
    .replace(/\uE000([^\uE001<>"]*)\uE001/g, '<span style="$1">')
    .replace(/\uE002/g, '</span>');
  html = formattingToHtml(trackedMarkersToHtml(html));
  return { html, warnings: res.messages.filter((m) => m.type === 'warning').map((m) => m.message) };
}

/** Đọc khổ giấy (ngang/dọc) và việc có đánh số trang ở chân trang từ tệp .docx gốc. */
async function readPageSetup(data: ArrayBuffer): Promise<PageSetup> {
  try {
    const zip = await JSZip.loadAsync(data);
    const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';
    const sz = /<w:pgSz\b[^>]*>/.exec(xml)?.[0] ?? '';
    const w = Number(/w:w="(\d+)"/.exec(sz)?.[1] ?? 0), h = Number(/w:h="(\d+)"/.exec(sz)?.[1] ?? 0);
    const landscape = /w:orient="landscape"/.test(sz) || (w > 0 && w > h);
    let pageNumbers = false;
    for (const name of Object.keys(zip.files).filter((n) => /^word\/footer\d*\.xml$/.test(n))) {
      if (/\bPAGE\b/.test(await zip.file(name)!.async('string'))) pageNumbers = true;
    }
    // Chữ đầu/chân trang mặc định của mục đầu tiên.
    const rels = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
    const partText = async (kind: 'header' | 'footer') => {
      const id = new RegExp(`<w:${kind}Reference\\b[^>]*w:type="default"[^>]*r:id="([^"]+)"|<w:${kind}Reference\\b[^>]*r:id="([^"]+)"[^>]*w:type="default"`).exec(xml);
      const rid = id?.[1] ?? id?.[2];
      const target = rid ? new RegExp(`<Relationship\\b[^>]*Id="${rid}"[^>]*Target="([^"]+)"|<Relationship\\b[^>]*Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels) : null;
      const file = target ? `word/${(target[1] ?? target[2]).replace(/^\/?word\//, '')}` : null;
      const part = file ? await zip.file(file)?.async('string') : undefined;
      return part ? headerFooterText(part) : '';
    };
    const header = await partText('header');
    const footer = await partText('footer');
    const settings = (await zip.file('word/settings.xml')?.async('string')) ?? '';
    const track = /<w:trackRevisions\b(?![^>]*w:val="(false|0|off)")/.test(settings);
    return { landscape, pageNumbers, track, header, footer };
  } catch {
    return { landscape: false, pageNumbers: false };
  }
}

export default function DocxEditor({ file, onClose }: { file: OpenedFile; onClose: () => void }) {
  const isNew = file.data.byteLength === 0;
  const [mode, setMode] = useState<'edit' | 'original'>('edit');
  const [dirty, setDirty] = useState(false);
  const [busy, setBusy] = useState<string | null>(isNew ? null : 'Đang mở…');
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [path, setPath] = useState(file.path);
  const [name, setName] = useState(file.name);
  const [zoom, setZoom] = useState(1);
  const [, force] = useState(0);
  const [setup, setSetup] = useState<PageSetup>({ landscape: false, pageNumbers: false });
  const [hfEdit, setHfEdit] = useState(false);
  const [find, setFind] = useState<{ q: string; repl: string; idx: number; replace: boolean } | null>(null);
  const originalRef = useRef<HTMLDivElement>(null);

  const editor = useEditor({
    extensions: [
      StarterKit,
      Underline,
      TextStyle,
      Color,
      FontFamily,
      FontSize,
      Subscript,
      Superscript,
      Link.configure({ openOnClick: false, autolink: true, HTMLAttributes: { rel: 'noopener noreferrer', target: null } }),
      PageBreak,
      LineHeight,
      TableOfContents,
      SearchHighlight,
      TrackChanges,
      Highlight.configure({ multicolor: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Table.configure({ resizable: true }),
      TableRow,
      TableHeader.extend({ addAttributes() { return { ...this.parent?.(), ...cellBackground }; } }),
      TableCell.extend({ addAttributes() { return { ...this.parent?.(), ...cellBackground }; } }),
      Image.configure({ allowBase64: true }),
    ],
    content: isNew ? '<h1>Tài liệu mới</h1><p></p>' : '',
    onUpdate: () => setDirty(true),
    onSelectionUpdate: () => force((n) => n + 1),
    onTransaction: () => force((n) => n + 1),
  });

  useEffect(() => {
    if (!editor || isNew) return;
    const setupP = readPageSetup(file.data);
    setupP.then(setSetup);
    docxToHtml(file.data)
      .then(async ({ html, warnings }) => {
        editor.commands.setContent(html || '<p></p>', false);
        if ((await setupP).track) editor.commands.setTrackChanges(true);
        if (warnings.length) setNotice('Một số định dạng phức tạp (màu chữ, căn lề, header/footer…) chỉ hiển thị đầy đủ ở chế độ "Bố cục gốc".');
      })
      .catch((e) => setError(`Không mở được tệp Word: ${e.message ?? e}`))
      .finally(() => setBusy(null));
  }, [editor]);

  useEffect(() => {
    if (mode !== 'original' || !originalRef.current || isNew) return;
    const el = originalRef.current;
    el.innerHTML = '';
    import('docx-preview')
      .then(({ renderAsync }) => renderAsync(file.data, el, undefined, { inWrapper: true, ignoreLastRenderedPageBreak: false }))
      .catch((e) => setError(`Không hiển thị được bố cục gốc: ${e.message ?? e}`));
  }, [mode]);

  const save = async (saveAs = false) => {
    if (!editor) return;
    setBusy('Đang lưu…');
    try {
      const out = await exportDocx(editor.getJSON() as PMNode, { ...setup, track: !!trackKey.getState(editor.state)?.on });
      const finalName = name.toLowerCase().endsWith('.docx') ? name : `${name}.docx`;
      const p = await saveFile({ name: finalName, kind: 'docx', path }, out, saveAs || (isNew && !path));
      if (p) { setPath(p); setName(p.split(/[\\/]/).pop() ?? finalName); }
      setDirty(false);
    } catch (e) {
      setError(`Lưu thất bại: ${(e as Error).message}`);
    } finally {
      setBusy(null);
    }
  };

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') { e.preventDefault(); window.print(); }
      if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'f' || e.key.toLowerCase() === 'h') && mode === 'edit') {
        e.preventDefault();
        setFind((f) => ({ q: f?.q ?? '', repl: f?.repl ?? '', idx: 0, replace: e.key.toLowerCase() === 'h' || !!f?.replace }));
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  const insertImage = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/png,image/jpeg,image/gif';
    input.onchange = () => {
      const f = input.files?.[0];
      if (!f || !editor) return;
      const r = new FileReader();
      r.onload = () => editor.chain().focus().setImage({ src: r.result as string }).run();
      r.readAsDataURL(f);
    };
    input.click();
  };

  // ---- Tìm & thay thế ----
  const doc = editor?.state.doc;
  const matches = useMemo(() => (doc && find?.q ? findInDoc(doc, find.q) : []), [doc, find?.q]);
  const curIdx = matches.length ? Math.min(find?.idx ?? 0, matches.length - 1) : 0;
  useEffect(() => {
    if (!editor) return;
    editor.view.dispatch(editor.state.tr.setMeta(searchKey, { query: find?.q ?? '', index: curIdx }));
    requestAnimationFrame(() => document.querySelector('.search-hit.cur')?.scrollIntoView({ block: 'center', behavior: 'smooth' }));
  }, [editor, find?.q, curIdx]);
  const step = (d: number) => find && matches.length && setFind({ ...find, idx: (curIdx + d + matches.length) % matches.length });
  const replaceOne = () => {
    if (!editor || !find || !matches.length) return;
    const m = matches[curIdx];
    editor.view.dispatch(editor.state.tr.insertText(find.repl, m.from, m.to));
  };
  const replaceAll = () => {
    if (!editor || !find || !matches.length) return;
    const tr = editor.state.tr;
    for (const m of [...matches].reverse()) tr.insertText(find.repl, m.from, m.to);
    editor.view.dispatch(tr);
  };
  const closeFind = () => {
    setFind(null);
    editor?.commands.focus();
  };

  const changes = useMemo(() => (doc ? listChanges(doc).length : 0), [doc]);
  const tracking = !!(editor && trackKey.getState(editor.state)?.on);
  const text = editor?.state.doc.textContent ?? '';
  const words = text.trim() ? text.trim().split(/\s+/).length : 0;

  return (
    <EditorShell
      kind="docx"
      title={name}
      dirty={dirty}
      onClose={onClose}
      onSave={() => save()}
      onSaveAs={() => save(true)}
      busy={busy}
      error={error}
      onDismissError={() => setError(null)}
      toolbar={editor && <DocxToolbar editor={editor} mode={mode} setMode={setMode} canOriginal={!isNew}
        insertImage={insertImage} zoom={zoom} setZoom={setZoom} setup={setup}
        setSetup={(p) => { setSetup(p); setDirty(true); }}
        findOpen={!!find} toggleFind={() => (find ? closeFind() : setFind({ q: '', repl: '', idx: 0, replace: true }))}
        tracking={tracking} changes={changes} openHeaderFooter={() => setHfEdit(true)} />}
      status={`${words} từ · ${text.length} ký tự${tracking ? ' · Đang theo dõi thay đổi' : ''}${changes ? ` · ${changes} thay đổi` : ''}${mode === 'original' ? ' · Chế độ xem bố cục gốc (chỉ đọc)' : ''}`}
    >
      {find && mode === 'edit' && (
        <form className="xl-find" onSubmit={(e) => { e.preventDefault(); step(1); }}>
          <input autoFocus placeholder="Tìm trong văn bản…" value={find.q}
            onChange={(e) => setFind({ ...find, q: e.target.value, idx: 0 })}
            onKeyDown={(e) => {
              if (e.key === 'Escape') closeFind();
              if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
            }} />
          <span className="xl-find-count">{find.q ? (matches.length ? `${curIdx + 1}/${matches.length}` : 'Không thấy') : ''}</span>
          <button type="button" className="btn" onClick={() => step(-1)} disabled={!matches.length} title="Kết quả trước">‹</button>
          <button type="button" className="btn" onClick={() => step(1)} disabled={!matches.length} title="Kết quả sau">›</button>
          {find.replace && (
            <>
              <input placeholder="Thay bằng…" value={find.repl} onChange={(e) => setFind({ ...find, repl: e.target.value })}
                onKeyDown={(e) => { if (e.key === 'Escape') closeFind(); if (e.key === 'Enter') { e.preventDefault(); replaceOne(); } }} />
              <button type="button" className="btn" onClick={replaceOne} disabled={!matches.length}>Thay</button>
              <button type="button" className="btn" onClick={replaceAll} disabled={!matches.length}>Thay tất cả</button>
            </>
          )}
          <button type="button" className="icon-btn" onClick={closeFind} title="Đóng">✕</button>
        </form>
      )}
      {notice && mode === 'edit' && (
        <div className="notice-bar">{notice}<button className="icon-btn" onClick={() => setNotice(null)}>✕</button></div>
      )}
      <div className="docx-scroll">
        {mode === 'edit' ? (
          <div className={`docx-page${setup.landscape ? ' landscape' : ''}`} style={{ zoom }}>
            {setup.landscape && <style>{'@page { size: A4 landscape; }'}</style>}
            {(setup.header || hfEdit) && (
              <div className="docx-hf top" onClick={() => setHfEdit(true)} title="Bấm để sửa đầu trang">{setup.header || 'Đầu trang'}</div>
            )}
            <EditorContent editor={editor} />
            {(setup.footer || setup.pageNumbers) && (
              <div className="docx-hf bottom" onClick={() => setHfEdit(true)} title="Bấm để sửa chân trang">
                {setup.footer}{setup.footer && setup.pageNumbers ? '\n' : ''}{setup.pageNumbers ? 'Trang x / y' : ''}
              </div>
            )}
          </div>
        ) : (
          <div className="docx-original" ref={originalRef} />
        )}
      </div>
      {hfEdit && (
        <div className="modal-backdrop" onClick={() => setHfEdit(false)}>
          <form className="modal hf-dialog" onClick={(e) => e.stopPropagation()} onSubmit={(e) => {
            e.preventDefault();
            const f = new FormData(e.currentTarget);
            setSetup({ ...setup, header: String(f.get('header') ?? ''), footer: String(f.get('footer') ?? ''), pageNumbers: f.get('pn') === 'on' });
            setDirty(true);
            setHfEdit(false);
          }}>
            <h3>Đầu trang và chân trang</h3>
            <label>Đầu trang<textarea name="header" rows={2} defaultValue={setup.header ?? ''} placeholder="Ví dụ: CÔNG TY ABC – Báo cáo nội bộ" /></label>
            <label>Chân trang<textarea name="footer" rows={2} defaultValue={setup.footer ?? ''} placeholder="Ví dụ: Địa chỉ, số điện thoại…" /></label>
            <label className="row"><input type="checkbox" name="pn" defaultChecked={setup.pageNumbers} /> Đánh số trang ("Trang x / y")</label>
            <div className="modal-actions">
              <button type="button" className="btn" onClick={() => setHfEdit(false)}>Huỷ</button>
              <button type="submit" className="btn primary">Xong</button>
            </div>
          </form>
        </div>
      )}
    </EditorShell>
  );
}

function DocxToolbar({ editor, mode, setMode, canOriginal, insertImage, zoom, setZoom, setup, setSetup, findOpen, toggleFind, tracking, changes, openHeaderFooter }: {
  editor: Editor; mode: 'edit' | 'original'; setMode(m: 'edit' | 'original'): void; canOriginal: boolean;
  insertImage(): void; zoom: number; setZoom(fn: (z: number) => number): void;
  setup: PageSetup; setSetup(p: PageSetup): void; findOpen: boolean; toggleFind(): void; tracking: boolean; changes: number;
  openHeaderFooter(): void;
}) {
  const toggleTrack = () => {
    if (!tracking) {
      try {
        if (!localStorage.getItem('dvd-author')) {
          const who = prompt('Tên của bạn (hiện bên cạnh các thay đổi)', currentAuthor());
          if (who?.trim()) localStorage.setItem('dvd-author', who.trim());
        }
      } catch { /* không lưu được tên: dùng tên mặc định */ }
    }
    editor.chain().focus().setTrackChanges(!tracking).run();
  };
  const ts = editor.getAttributes('textStyle');
  const family = (ts.fontFamily as string | undefined)?.replace(/["']/g, '').split(',')[0].trim() ?? '';
  const sizePt = ts.fontSize ? Math.round(parseFloat(ts.fontSize) * (String(ts.fontSize).endsWith('px') ? 0.75 : 1)) : '';
  const setLink = () => {
    const prev = editor.getAttributes('link').href as string | undefined;
    const url = prompt('Địa chỉ liên kết (để trống để bỏ liên kết)', prev ?? 'https://');
    if (url === null) return;
    if (!url.trim() || url.trim() === 'https://') return void editor.chain().focus().extendMarkRange('link').unsetLink().run();
    const href = /^(https?:|mailto:|tel:)/i.test(url.trim()) ? url.trim() : `https://${url.trim()}`;
    if (editor.state.selection.empty && !editor.isActive('link')) {
      editor.chain().focus().insertContent({ type: 'text', text: url.trim(), marks: [{ type: 'link', attrs: { href } }] }).run();
    } else editor.chain().focus().extendMarkRange('link').setLink({ href }).run();
  };
  const c = () => editor.chain().focus();
  const level = [1, 2, 3, 4].find((l) => editor.isActive('heading', { level: l }));
  const inTable = editor.isActive('table');
  const editing = mode === 'edit';
  return (
    <>
      {editing && (
        <>
          <ToolButton icon="↶" label="Hoàn tác" onClick={() => c().undo().run()} disabled={!editor.can().undo()} />
          <ToolButton icon="↷" label="Làm lại" onClick={() => c().redo().run()} disabled={!editor.can().redo()} />
          <Sep />
          <label className="tb-field" title="Kiểu đoạn">
            <select value={level ?? 0} onChange={(e) => {
              const v = Number(e.target.value);
              v ? c().setHeading({ level: v as 1 | 2 | 3 | 4 }).run() : c().setParagraph().run();
            }}>
              <option value={0}>Văn bản thường</option>
              <option value={1}>Tiêu đề 1</option>
              <option value={2}>Tiêu đề 2</option>
              <option value={3}>Tiêu đề 3</option>
              <option value={4}>Tiêu đề 4</option>
            </select>
          </label>
          <label className="tb-field" title="Phông chữ">
            <select value={FONTS.includes(family) ? family : ''} onChange={(e) => {
              e.target.value ? c().setFontFamily(e.target.value).run() : c().unsetFontFamily().run();
            }}>
              <option value="">Phông mặc định</option>
              {FONTS.map((f) => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
            </select>
          </label>
          <label className="tb-field" title="Cỡ chữ">
            <select value={sizePt} onChange={(e) => {
              e.target.value ? c().setFontSize(`${e.target.value}pt`).run() : c().unsetFontSize().run();
            }}>
              <option value="">Cỡ</option>
              {SIZES.map((n) => <option key={n} value={n}>{n}</option>)}
              {sizePt && !SIZES.includes(sizePt) && <option value={sizePt}>{sizePt}</option>}
            </select>
          </label>
          <ToolButton icon={<b>B</b>} label="Đậm" active={editor.isActive('bold')} onClick={() => c().toggleBold().run()} />
          <ToolButton icon={<i>I</i>} label="Nghiêng" active={editor.isActive('italic')} onClick={() => c().toggleItalic().run()} />
          <ToolButton icon={<u>U</u>} label="Gạch chân" active={editor.isActive('underline')} onClick={() => c().toggleUnderline().run()} />
          <ToolButton icon={<s>S</s>} label="Gạch ngang" active={editor.isActive('strike')} onClick={() => c().toggleStrike().run()} />
          <ToolButton icon={<span>x<sup>2</sup></span>} label="Chỉ số trên" active={editor.isActive('superscript')} onClick={() => c().toggleSuperscript().run()} />
          <ToolButton icon={<span>x<sub>2</sub></span>} label="Chỉ số dưới" active={editor.isActive('subscript')} onClick={() => c().toggleSubscript().run()} />
          <label className="tb-field" title="Màu chữ">
            <span className="tb-mini">A</span>
            <input type="color" value={editor.getAttributes('textStyle').color ?? '#000000'}
              onChange={(e) => c().setColor(e.target.value).run()} />
          </label>
          <label className="tb-field" title="Tô nền chữ">
            <span className="tb-mini">▮</span>
            <input type="color" defaultValue="#fff59d" onChange={(e) => c().setHighlight({ color: e.target.value }).run()} />
          </label>
          <ToolButton icon="⌫" label="Xoá định dạng" onClick={() => c().unsetAllMarks().clearNodes().run()} />
          <Sep />
          <ToolButton icon="⇤" label="Căn trái" active={editor.isActive({ textAlign: 'left' })} onClick={() => c().setTextAlign('left').run()} />
          <ToolButton icon="↔" label="Căn giữa" active={editor.isActive({ textAlign: 'center' })} onClick={() => c().setTextAlign('center').run()} />
          <ToolButton icon="⇥" label="Căn phải" active={editor.isActive({ textAlign: 'right' })} onClick={() => c().setTextAlign('right').run()} />
          <ToolButton icon="☰" label="Căn đều" active={editor.isActive({ textAlign: 'justify' })} onClick={() => c().setTextAlign('justify').run()} />
          <label className="tb-field" title="Giãn dòng">
            <select value={(editor.getAttributes('paragraph').lineHeight ?? editor.getAttributes('heading').lineHeight ?? '') as string}
              onChange={(e) => c().setLineHeight(e.target.value || null).run()}>
              <option value="">Giãn dòng</option>
              {['1', '1.15', '1.5', '2', '2.5', '3'].map((v) => <option key={v} value={v}>{v.replace('.', ',')}</option>)}
            </select>
          </label>
          <Sep />
          <ToolButton icon="•" label="Dấu đầu dòng" active={editor.isActive('bulletList')} onClick={() => c().toggleBulletList().run()} />
          <ToolButton icon="1." label="Đánh số" active={editor.isActive('orderedList')} onClick={() => c().toggleOrderedList().run()} />
          <ToolButton icon="→" label="Thụt vào" onClick={() => c().sinkListItem('listItem').run()} disabled={!editor.can().sinkListItem('listItem')} />
          <ToolButton icon="←" label="Thụt ra" onClick={() => c().liftListItem('listItem').run()} disabled={!editor.can().liftListItem('listItem')} />
          <ToolButton icon="❝" label="Trích dẫn" active={editor.isActive('blockquote')} onClick={() => c().toggleBlockquote().run()} />
          <Sep />
          <ToolButton icon="▦" label="Chèn bảng" onClick={() => c().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()} />
          {inTable && (
            <>
              <ToolButton icon="+↓" label="Thêm hàng" onClick={() => c().addRowAfter().run()} />
              <ToolButton icon="+→" label="Thêm cột" onClick={() => c().addColumnAfter().run()} />
              <ToolButton icon="−↓" label="Xoá hàng" onClick={() => c().deleteRow().run()} />
              <ToolButton icon="−→" label="Xoá cột" onClick={() => c().deleteColumn().run()} />
              <ToolButton icon="⊞" label="Gộp/tách ô" onClick={() => c().mergeOrSplit().run()} />
              <label className="tb-field" title="Màu nền ô">
                <span className="tb-mini">▦</span>
                <input type="color" defaultValue="#fff2cc" onChange={(e) => c().setCellAttribute('backgroundColor', e.target.value).run()} />
              </label>
              <ToolButton icon="∅" label="Bỏ màu ô" onClick={() => c().setCellAttribute('backgroundColor', null).run()} />
              <ToolButton icon="✕▦" label="Xoá bảng" onClick={() => c().deleteTable().run()} />
            </>
          )}
          <ToolButton icon="🖼" label="Chèn ảnh" onClick={insertImage} />
          <ToolButton icon="🔗" label="Liên kết" active={editor.isActive('link')} onClick={setLink} />
          <ToolButton icon="―" label="Đường kẻ" onClick={() => c().setHorizontalRule().run()} />
          <ToolButton icon="⤓" label="Ngắt trang" onClick={() => c().setPageBreak().run()} />
          <ToolButton icon="≣" label="Mục lục" onClick={() => c().insertTableOfContents().run()} />
          <Sep />
          <ToolButton icon="📝" label="Theo dõi thay đổi" active={tracking} onClick={toggleTrack} />
          {(changes > 0 || tracking) && (
            <>
              <ToolButton icon="‹" label="Thay đổi trước" onClick={() => c().goToChange(-1).run()} disabled={!changes} />
              <ToolButton icon="›" label="Thay đổi sau" onClick={() => c().goToChange(1).run()} disabled={!changes} />
              <ToolButton icon="✓" label="Chấp nhận" onClick={() => c().acceptChange().goToChange(1).run()} disabled={!changes} />
              <ToolButton icon="✗" label="Từ chối" onClick={() => c().rejectChange().goToChange(1).run()} disabled={!changes} />
              <ToolButton icon="✓✓" label="Chấp nhận tất cả" onClick={() => c().acceptAllChanges().run()} disabled={!changes} />
              <ToolButton icon="✗✗" label="Từ chối tất cả" onClick={() => c().rejectAllChanges().run()} disabled={!changes} />
            </>
          )}
          <Sep />
          <ToolButton icon="🔍" label="Tìm & thay" active={findOpen} onClick={toggleFind} />
          <label className="tb-field" title="Hướng trang">
            <select value={setup.landscape ? 'l' : 'p'} onChange={(e) => setSetup({ ...setup, landscape: e.target.value === 'l' })}>
              <option value="p">A4 dọc</option>
              <option value="l">A4 ngang</option>
            </select>
          </label>
          <ToolButton icon="#" label="Số trang" active={setup.pageNumbers} onClick={() => setSetup({ ...setup, pageNumbers: !setup.pageNumbers })} />
          <ToolButton icon="▭" label="Đầu/chân trang" active={!!(setup.header || setup.footer)} onClick={openHeaderFooter} />
          <Sep />
          <ToolButton icon="−" label="Thu nhỏ" onClick={() => setZoom((z) => Math.max(0.5, +(z - 0.1).toFixed(2)))} />
          <span className="tb-zoom">{Math.round(zoom * 100)}%</span>
          <ToolButton icon="+" label="Phóng to" onClick={() => setZoom((z) => Math.min(2.5, +(z + 0.1).toFixed(2)))} />
          <ToolButton icon="⎙" label="In / PDF" onClick={() => window.print()} />
        </>
      )}
      {canOriginal && (
        <>
          <Sep />
          <ToolButton icon="✎" label="Soạn thảo" active={mode === 'edit'} onClick={() => setMode('edit')} />
          <ToolButton icon="👁" label="Bố cục gốc" active={mode === 'original'} onClick={() => setMode('original')} />
        </>
      )}
    </>
  );
}
