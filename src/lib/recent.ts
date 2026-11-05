// Lưu các tệp mở gần đây trong IndexedDB của thiết bị (không gửi đi đâu cả).
import type { DocKind, OpenedFile } from './fileio';

export interface RecentEntry { id: string; name: string; kind: DocKind; size: number; openedAt: number; path?: string }

const DB = 'dvddocs';
const STORE = 'recent';
const MAX = 12;
const MAX_BYTES = 30 * 1024 * 1024;

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function tx<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return open().then((db) => new Promise<T>((resolve, reject) => {
    const req = fn(db.transaction(STORE, mode).objectStore(STORE));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
}

export async function listRecent(): Promise<RecentEntry[]> {
  try {
    const all = await tx<(RecentEntry & { data: ArrayBuffer })[]>('readonly', (s) => s.getAll());
    return all.map(({ data: _d, ...e }) => e).sort((a, b) => b.openedAt - a.openedAt);
  } catch {
    return [];
  }
}

export async function addRecent(f: OpenedFile) {
  if (!f.data.byteLength || f.data.byteLength > MAX_BYTES) return;
  try {
    const id = `${f.kind}:${f.path ?? f.name}`;
    await tx('readwrite', (s) => s.put({ id, name: f.name, kind: f.kind, size: f.data.byteLength, openedAt: Date.now(), path: f.path, data: f.data.slice(0) }));
    const list = await listRecent();
    for (const old of list.slice(MAX)) await tx('readwrite', (s) => s.delete(old.id));
  } catch {
    /* bộ nhớ đầy hoặc trình duyệt chặn IndexedDB: bỏ qua */
  }
}

export async function loadRecent(id: string): Promise<OpenedFile | null> {
  const e = await tx<(RecentEntry & { data: ArrayBuffer }) | undefined>('readonly', (s) => s.get(id));
  return e ? { name: e.name, kind: e.kind, data: e.data, path: e.path } : null;
}

export async function removeRecent(id: string) {
  await tx('readwrite', (s) => s.delete(id));
}
