import { useEffect, useRef, useState } from 'react';
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
import { OpenedFile, saveFile } from '../../lib/fileio';
import { EditorShell, Sep, ToolButton } from '../../components/EditorShell';
import { exportDocx, PMNode } from './docxExport';

async function docxToHtml(data: ArrayBuffer): Promise<{ html: string; warnings: string[] }> {
  const mammoth = await import('mammoth');
  const res = await mammoth.convertToHtml(
    { arrayBuffer: data },
    { styleMap: ['u => u', 'strike => s', "p[style-name='Title'] => h1:fresh", "p[style-name='Subtitle'] => h2:fresh"] },
  );
  return { html: res.value, warnings: res.messages.filter((m) => m.type === 'warning').map((m) => m.message) };
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
  const originalRef = useRef<HTMLDivElement>(null);

  const editor = useEditor({
    extensions: [
      StarterKit,
      Underline,
      TextStyle,
      Color,
      Highlight.configure({ multicolor: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Table.configure({ resizable: true }),
      TableRow,
      TableHeader,
      TableCell,
      Image.configure({ allowBase64: true }),
    ],
    content: isNew ? '<h1>Tài liệu mới</h1><p></p>' : '',
    onUpdate: () => setDirty(true),
    onSelectionUpdate: () => force((n) => n + 1),
    onTransaction: () => force((n) => n + 1),
  });

  useEffect(() => {
    if (!editor || isNew) return;
    docxToHtml(file.data)
      .then(({ html, warnings }) => {
        editor.commands.setContent(html || '<p></p>', false);
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
      const out = await exportDocx(editor.getJSON() as PMNode);
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
        insertImage={insertImage} zoom={zoom} setZoom={setZoom} />}
      status={`${words} từ · ${text.length} ký tự${mode === 'original' ? ' · Chế độ xem bố cục gốc (chỉ đọc)' : ''}`}
    >
      {notice && mode === 'edit' && (
        <div className="notice-bar">{notice}<button className="icon-btn" onClick={() => setNotice(null)}>✕</button></div>
      )}
      <div className="docx-scroll">
        {mode === 'edit' ? (
          <div className="docx-page" style={{ zoom }}>
            <EditorContent editor={editor} />
          </div>
        ) : (
          <div className="docx-original" ref={originalRef} />
        )}
      </div>
    </EditorShell>
  );
}

function DocxToolbar({ editor, mode, setMode, canOriginal, insertImage, zoom, setZoom }: {
  editor: Editor; mode: 'edit' | 'original'; setMode(m: 'edit' | 'original'): void; canOriginal: boolean;
  insertImage(): void; zoom: number; setZoom(fn: (z: number) => number): void;
}) {
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
          <ToolButton icon={<b>B</b>} label="Đậm" active={editor.isActive('bold')} onClick={() => c().toggleBold().run()} />
          <ToolButton icon={<i>I</i>} label="Nghiêng" active={editor.isActive('italic')} onClick={() => c().toggleItalic().run()} />
          <ToolButton icon={<u>U</u>} label="Gạch chân" active={editor.isActive('underline')} onClick={() => c().toggleUnderline().run()} />
          <ToolButton icon={<s>S</s>} label="Gạch ngang" active={editor.isActive('strike')} onClick={() => c().toggleStrike().run()} />
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
              <ToolButton icon="✕▦" label="Xoá bảng" onClick={() => c().deleteTable().run()} />
            </>
          )}
          <ToolButton icon="🖼" label="Chèn ảnh" onClick={insertImage} />
          <ToolButton icon="―" label="Đường kẻ" onClick={() => c().setHorizontalRule().run()} />
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
