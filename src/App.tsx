import { lazy, Suspense, useEffect, useRef, useState } from 'react';
import { DocKind, OpenedFile, desktopToOpened, fileToOpened, pickFile, platform } from './lib/fileio';
import { RecentEntry, addRecent, listRecent, loadRecent, removeRecent } from './lib/recent';
import { DraftEntry, deleteDraft, hasUnsaved, listDrafts, loadDraft } from './lib/drafts';
import { ErrorBoundary } from './components/ErrorBoundary';

// Mỗi trình soạn thảo được tải riêng để màn hình chính mở nhanh.
const PdfEditor = lazy(() => import('./editors/pdf/PdfEditor'));
const DocxEditor = lazy(() => import('./editors/docx/DocxEditor'));
const XlsxEditor = lazy(() => import('./editors/xlsx/XlsxEditor'));

const KIND_INFO: Record<DocKind, { label: string; desc: string; newName: string }> = {
  docx: { label: 'Word', desc: 'Văn bản .docx', newName: 'Tài liệu mới.docx' },
  xlsx: { label: 'Excel', desc: 'Bảng tính .xlsx', newName: 'Bảng tính mới.xlsx' },
  pdf: { label: 'PDF', desc: 'Tài liệu .pdf', newName: 'PDF mới.pdf' },
};

function fmtSize(n: number) {
  return n > 1e6 ? `${(n / 1e6).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1e3))} KB`;
}

export default function App() {
  const [file, setFile] = useState<OpenedFile | null>(null);
  const [recent, setRecent] = useState<RecentEntry[]>([]);
  const [drafts, setDrafts] = useState<DraftEntry[]>([]);
  const [session, setSession] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [drag, setDrag] = useState(false);

  useEffect(() => {
    if (file) return;
    listRecent().then(setRecent);
    // Bản nháp được ghi khi trình soạn thảo đóng lại: đọc lại sau một nhịp.
    const t = setTimeout(() => listDrafts().then(setDrafts), 300);
    listDrafts().then(setDrafts);
    return () => clearTimeout(t);
  }, [file]);

  const open = (f: OpenedFile | null) => {
    if (!f) return;
    setError(null);
    if (!f.restored) addRecent(f);
    setSession((n) => n + 1);
    setFile(f);
  };
  const openRef = useRef(open);
  openRef.current = open;

  // Bản máy tính: tệp được mở bằng nhấp đúp / "Mở bằng DVDDocs" / kéo vào biểu tượng ứng dụng.
  useEffect(() => {
    const desk = window.dvdDesktop;
    if (!desk) return;
    const take = (f: { name: string; path: string; data: ArrayBuffer } | null) => {
      try {
        const o = desktopToOpened(f);
        if (!o) return;
        if (hasUnsaved() && !confirm(`Mở “${o.name}”? Tài liệu đang sửa có thay đổi chưa lưu (vẫn được giữ thành bản nháp để khôi phục).`)) return;
        openRef.current(o);
      } catch (e) {
        setError((e as Error).message);
      }
    };
    desk.getLaunchFile?.().then(take).catch(() => {});
    return desk.onOpenFile?.(take);
  }, []);

  const handle = (p: Promise<OpenedFile | null>) => p.then(open).catch((e) => setError((e as Error).message));

  const createNew = async (kind: DocKind) => {
    const data = kind === 'pdf' ? ((await (await import('./editors/pdf/pdfOps')).createBlankPdf()).slice().buffer as ArrayBuffer) : new ArrayBuffer(0);
    setSession((n) => n + 1);
    setFile({ name: KIND_INFO[kind].newName, kind, data });
  };

  if (file) {
    const props = { file, onClose: () => setFile(null) };
    return (
      <ErrorBoundary onHome={() => setFile(null)} key={session}>
        <Suspense fallback={<div className="loading full">Đang tải trình soạn thảo…</div>}>
          {file.kind === 'pdf' && <PdfEditor {...props} />}
          {file.kind === 'docx' && <DocxEditor {...props} />}
          {file.kind === 'xlsx' && <XlsxEditor {...props} />}
        </Suspense>
      </ErrorBoundary>
    );
  }

  return (
    <div
      className={`home${drag ? ' dragging' : ''}`}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={(e) => { if (e.currentTarget === e.target) setDrag(false); }}
      onDrop={(e) => {
        e.preventDefault();
        setDrag(false);
        const f = e.dataTransfer.files[0];
        if (f) handle(fileToOpened(f));
      }}
    >
      <header className="home-header">
        <img src="./icon.svg" alt="" width={40} height={40} />
        <div>
          <h1>DVDDocs</h1>
          <p className="muted">Xem và chỉnh sửa PDF, Word, Excel ngay trên thiết bị của bạn</p>
        </div>
      </header>

      <section className="home-open">
        <button className="btn primary big" onClick={() => handle(pickFile())}>Mở tệp…</button>
        <span className="muted">{platform() === 'desktop' || (platform() === 'web' && !matchMedia('(pointer: coarse)').matches) ? 'hoặc kéo thả tệp .pdf, .docx, .xlsx, .csv vào đây' : 'Hỗ trợ .pdf, .docx, .xlsx, .csv'}</span>
      </section>
      {error && <div className="error-bar" role="alert">{error}</div>}

      {drafts.length > 0 && (
        <section className="drafts" aria-labelledby="drafts-title">
          <h2 id="drafts-title">Khôi phục tài liệu chưa lưu</h2>
          <p className="muted small">Các thay đổi này chưa được lưu vào tệp khi ứng dụng bị đóng. Mở để xem lại rồi bấm Lưu, hoặc bỏ đi.</p>
          <ul className="recent">
            {drafts.map((d) => (
              <li key={d.id}>
                <button className="recent-item" onClick={() => handle(loadDraft(d.id))}>
                  <span className={`badge badge-${d.kind}`}>{KIND_INFO[d.kind].label}</span>
                  <span className="recent-name">{d.name}</span>
                  <span className="muted">Bản nháp · {fmtSize(d.size)} · {new Date(d.savedAt).toLocaleString('vi-VN')}</span>
                </button>
                <button className="icon-btn" title="Bỏ bản nháp" aria-label={`Bỏ bản nháp ${d.name}`} onClick={() => {
                  if (confirm(`Bỏ bản nháp chưa lưu của “${d.name}”? Không thể hoàn tác.`)) deleteDraft(d.id).then(() => listDrafts().then(setDrafts));
                }}>✕</button>
              </li>
            ))}
          </ul>
        </section>
      )}

      <h2>Tạo mới</h2>
      <section className="new-grid">
        {(['docx', 'xlsx', 'pdf'] as DocKind[]).map((k) => (
          <button key={k} className={`new-card kind-${k}`} onClick={() => createNew(k)}>
            <span className={`badge badge-${k}`}>{KIND_INFO[k].label}</span>
            <span>{KIND_INFO[k].desc} trống</span>
          </button>
        ))}
      </section>

      <h2>Gần đây</h2>
      {recent.length === 0 ? (
        <p className="muted">Chưa có tệp nào. Các tệp bạn mở sẽ hiện ở đây (lưu trên chính thiết bị này).</p>
      ) : (
        <ul className="recent">
          {recent.map((r) => (
            <li key={r.id}>
              <button className="recent-item" onClick={() => handle(loadRecent(r.id))}>
                <span className={`badge badge-${r.kind}`}>{KIND_INFO[r.kind].label}</span>
                <span className="recent-name">{r.name}</span>
                <span className="muted">{fmtSize(r.size)} · {new Date(r.openedAt).toLocaleString('vi-VN')}</span>
              </button>
              <button className="icon-btn" title="Xoá khỏi danh sách" aria-label={`Xoá ${r.name} khỏi danh sách`} onClick={() => removeRecent(r.id).then(() => listRecent().then(setRecent))}>✕</button>
            </li>
          ))}
        </ul>
      )}
      <footer className="home-footer muted">Mọi tệp được xử lý ngay trên máy, không tải lên máy chủ nào.</footer>
    </div>
  );
}
