import { useEffect, useState, type ReactNode } from 'react';
import type { DocKind } from '../lib/fileio';
import type { DraftControl } from '../lib/drafts';
import { ShortcutsDialog } from './ShortcutsDialog';

const KIND_LABEL: Record<DocKind, string> = { pdf: 'PDF', docx: 'Word', xlsx: 'Excel' };

export function EditorShell(props: {
  kind: DocKind;
  title: string;
  dirty: boolean;
  busy?: string | null;
  error?: string | null;
  onDismissError?: () => void;
  onClose(): void;
  onSave(): void;
  onSaveAs(): void;
  /** Bản nháp tự động lưu của tài liệu (xoá khi người dùng chọn đóng không lưu). */
  draft?: DraftControl;
  /** Tài liệu chỉ đọc (ví dụ PDF có mật khẩu): ẩn nút Lưu. */
  readOnly?: boolean;
  toolbar: ReactNode;
  status?: ReactNode;
  children: ReactNode;
}) {
  const [help, setHelp] = useState(false);
  const close = () => {
    if (props.dirty) {
      if (!confirm('Tài liệu có thay đổi chưa lưu. Đóng và bỏ các thay đổi này?')) return;
      props.draft?.discard();
    }
    props.onClose();
  };

  // Đóng thẻ / tải lại trang / đóng cửa sổ khi còn thay đổi chưa lưu: hỏi lại (bản máy tính hiện hộp thoại riêng).
  useEffect(() => {
    if (!props.dirty) return;
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      void props.draft?.flush();
      e.preventDefault();
      e.returnValue = '';
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => window.removeEventListener('beforeunload', onBeforeUnload);
  }, [props.dirty, props.draft]);

  // F1 hoặc Ctrl+/ : bảng phím tắt.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'F1' || ((e.ctrlKey || e.metaKey) && e.key === '/')) { e.preventDefault(); setHelp((v) => !v); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  return (
    <div className={`editor kind-${props.kind}`}>
      <header className="editor-header">
        <button className="icon-btn" onClick={close} title="Về trang chủ" aria-label="Về trang chủ">‹</button>
        <span className={`badge badge-${props.kind}`}>{KIND_LABEL[props.kind]}</span>
        <h1 className="editor-title" title={props.title}>
          {props.title}
          {props.dirty && <span className="dirty" title="Chưa lưu"> •</span>}
        </h1>
        <span style={{ flex: 1 }} />
        {props.busy && <span className="busy" role="status">{props.busy}</span>}
        <button className="icon-btn help-btn" onClick={() => setHelp(true)} title="Phím tắt (F1)" aria-label="Phím tắt">?</button>
        {props.readOnly ? <span className="readonly-tag" title="Tài liệu này chỉ xem được, không lưu được thay đổi">Chỉ đọc</span> : (
          <>
            <button className="btn" onClick={props.onSaveAs}>Lưu thành…</button>
            <button className="btn primary" onClick={props.onSave}>Lưu</button>
          </>
        )}
      </header>
      <div className="toolbar" role="toolbar" aria-label="Thanh công cụ">{props.toolbar}</div>
      {props.error && (
        <div className="error-bar" role="alert">
          {props.error}
          <button className="icon-btn" onClick={props.onDismissError} aria-label="Đóng thông báo">✕</button>
        </div>
      )}
      <main className="editor-main">{props.children}</main>
      {props.status && <footer className="statusbar" aria-live="polite">{props.status}</footer>}
      {help && <ShortcutsDialog kind={props.kind} onClose={() => setHelp(false)} />}
    </div>
  );
}

export function ToolButton(p: { icon: ReactNode; label: string; onClick(): void; active?: boolean; disabled?: boolean }) {
  return (
    <button
      className={`tool-btn${p.active ? ' active' : ''}`}
      onClick={p.onClick}
      disabled={p.disabled}
      title={p.label}
      aria-label={p.label}
      aria-pressed={p.active}
      onMouseDown={(e) => e.preventDefault()}
    >
      <span className="tool-icon">{p.icon}</span>
      <span className="tool-label">{p.label}</span>
    </button>
  );
}

export const Sep = () => <span className="tb-sep" />;
