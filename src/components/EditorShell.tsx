import type { ReactNode } from 'react';
import type { DocKind } from '../lib/fileio';

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
  toolbar: ReactNode;
  status?: ReactNode;
  children: ReactNode;
}) {
  const close = () => {
    if (props.dirty && !confirm('Tài liệu có thay đổi chưa lưu. Vẫn đóng?')) return;
    props.onClose();
  };
  return (
    <div className={`editor kind-${props.kind}`}>
      <header className="editor-header">
        <button className="icon-btn" onClick={close} title="Về trang chủ">‹</button>
        <span className={`badge badge-${props.kind}`}>{KIND_LABEL[props.kind]}</span>
        <h1 className="editor-title" title={props.title}>
          {props.title}
          {props.dirty && <span className="dirty" title="Chưa lưu"> •</span>}
        </h1>
        <span style={{ flex: 1 }} />
        {props.busy && <span className="busy">{props.busy}</span>}
        <button className="btn" onClick={props.onSaveAs}>Lưu thành…</button>
        <button className="btn primary" onClick={props.onSave}>Lưu</button>
      </header>
      <div className="toolbar">{props.toolbar}</div>
      {props.error && (
        <div className="error-bar" role="alert">
          {props.error}
          <button className="icon-btn" onClick={props.onDismissError}>✕</button>
        </div>
      )}
      <main className="editor-main">{props.children}</main>
      {props.status && <footer className="statusbar">{props.status}</footer>}
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
      aria-pressed={p.active}
      onMouseDown={(e) => e.preventDefault()}
    >
      <span className="tool-icon">{p.icon}</span>
      <span className="tool-label">{p.label}</span>
    </button>
  );
}

export const Sep = () => <span className="tb-sep" />;
