// Tự động lưu bản nháp của tài liệu đang sửa vào IndexedDB của thiết bị, để khôi phục khi ứng dụng
// bị đóng đột ngột, hết pin, trình duyệt tải lại trang hoặc iOS tắt ứng dụng chạy nền.
import { useEffect, useRef } from 'react';
import type { DocKind, OpenedFile } from './fileio';
import { tx } from './db';
import { uid } from './uid';

export interface DraftEntry { id: string; name: string; kind: DocKind; path?: string; size: number; savedAt: number }

/** Bản nháp lớn hơn mức này không được lưu (tránh làm đầy bộ nhớ của trình duyệt). */
export const MAX_DRAFT_BYTES = 60 * 1024 * 1024;
/** Thời gian chờ sau thay đổi đầu tiên trước khi ghi bản nháp. */
export const AUTOSAVE_DELAY = 8000;

export async function listDrafts(): Promise<DraftEntry[]> {
  try {
    const all = await tx<(DraftEntry & { data: ArrayBuffer })[]>('drafts', 'readonly', (s) => s.getAll());
    return all.map(({ data: _d, ...e }) => e).sort((a, b) => b.savedAt - a.savedAt);
  } catch {
    return [];
  }
}

export async function putDraft(e: DraftEntry & { data: ArrayBuffer }) {
  await tx('drafts', 'readwrite', (s) => s.put(e));
}

export async function loadDraft(id: string): Promise<OpenedFile | null> {
  const e = await tx<(DraftEntry & { data: ArrayBuffer }) | undefined>('drafts', 'readonly', (s) => s.get(id));
  return e ? { name: e.name, kind: e.kind, data: e.data, path: e.path, draftId: e.id, restored: true } : null;
}

export async function deleteDraft(id: string) {
  try { await tx('drafts', 'readwrite', (s) => s.delete(id)); } catch { /* không có IndexedDB: bỏ qua */ }
}

// Trạng thái "có thay đổi chưa lưu" của trình soạn thảo đang mở, để màn hình chính và bản máy tính hỏi trước khi đóng.
let unsaved = false;
export const setUnsaved = (v: boolean) => { unsaved = v; };
export const hasUnsaved = () => unsaved;

export interface DraftControl {
  id: string;
  /** Người dùng chọn bỏ thay đổi: xoá bản nháp và không ghi lại nữa. */
  discard(): void;
  /** Ghi bản nháp ngay (khi ứng dụng chuyển xuống nền…). */
  flush(): Promise<void>;
}

/**
 * Tự động lưu bản nháp khi tài liệu có thay đổi chưa lưu.
 * - `snapshot` tạo nội dung tệp hiện tại (giống khi bấm Lưu); trả null nếu chưa sẵn sàng.
 * - `deps`: các giá trị đổi khi nội dung đổi (số phần tử phải cố định).
 * Khi đã lưu (dirty = false) thì bản nháp bị xoá.
 */
export function useDraftAutosave(opts: {
  file: OpenedFile;
  name: string;
  path?: string;
  dirty: boolean;
  snapshot: () => Promise<ArrayBuffer | Uint8Array | null> | null;
  deps: unknown[];
}): DraftControl {
  const latest = useRef(opts);
  latest.current = opts;
  const state = useRef({ id: opts.file.draftId ?? `draft:${uid()}`, timer: 0 as ReturnType<typeof setTimeout> | 0, running: false, again: false, discarded: false, written: !!opts.file.draftId });

  const flush = async () => {
    const st = state.current;
    if (st.timer) { clearTimeout(st.timer); st.timer = 0; }
    if (st.discarded || !latest.current.dirty) return;
    if (st.running) { st.again = true; return; }
    st.running = true;
    try {
      const out = await latest.current.snapshot();
      if (!out || st.discarded || !latest.current.dirty) return;
      const data = out instanceof Uint8Array ? (out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer) : out;
      if (!data.byteLength || data.byteLength > MAX_DRAFT_BYTES) return;
      const { file, name, path } = latest.current;
      await putDraft({ id: st.id, name, kind: file.kind, path, size: data.byteLength, savedAt: Date.now(), data });
      st.written = true;
      // Vừa lưu xong tệp trong lúc đang ghi bản nháp: bản nháp không còn cần nữa.
      if (!latest.current.dirty || st.discarded) await deleteDraft(st.id);
    } catch {
      /* hết bộ nhớ, trình duyệt chặn IndexedDB, tài liệu đang dở…: lần sau thử lại */
    } finally {
      st.running = false;
      if (st.again) { st.again = false; schedule(); }
    }
  };

  const schedule = () => {
    const st = state.current;
    if (!st.timer && !st.discarded) st.timer = setTimeout(() => { st.timer = 0; void flush(); }, AUTOSAVE_DELAY);
  };

  useEffect(() => {
    setUnsaved(opts.dirty);
    const st = state.current;
    if (opts.dirty) schedule();
    else {
      if (st.timer) { clearTimeout(st.timer); st.timer = 0; }
      if (st.written) { st.written = false; void deleteDraft(st.id); }
    }
  }, [opts.dirty, ...opts.deps]);

  useEffect(() => {
    // Ứng dụng xuống nền (chuyển ứng dụng trên điện thoại, đóng thẻ…): ghi ngay.
    const onHide = () => { if (document.visibilityState === 'hidden') void flush(); };
    const onPageHide = () => void flush();
    document.addEventListener('visibilitychange', onHide);
    window.addEventListener('pagehide', onPageHide);
    return () => {
      document.removeEventListener('visibilitychange', onHide);
      window.removeEventListener('pagehide', onPageHide);
      setUnsaved(false);
      // Trình soạn thảo bị gỡ khi còn thay đổi chưa lưu (lỗi giao diện…): cố ghi bản nháp lần cuối.
      void flush();
    };
  }, []);

  return {
    id: state.current.id,
    discard: () => {
      const st = state.current;
      st.discarded = true;
      if (st.timer) { clearTimeout(st.timer); st.timer = 0; }
      setUnsaved(false);
      void deleteDraft(st.id);
    },
    flush,
  };
}
