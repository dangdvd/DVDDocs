import 'fake-indexeddb/auto';
import { expect, it } from 'vitest';

it('nâng cấp cơ sở dữ liệu cũ (chỉ có tệp gần đây) mà không mất dữ liệu', async () => {
  // Cơ sở dữ liệu phiên bản 1 của các bản trước 0.5.
  await new Promise<void>((resolve, reject) => {
    const req = indexedDB.open('dvddocs', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('recent', { keyPath: 'id' }).put({ id: 'pdf:cu.pdf', name: 'cu.pdf', kind: 'pdf', size: 1, openedAt: 1, data: new ArrayBuffer(1) });
    req.onsuccess = () => { req.result.close(); resolve(); };
    req.onerror = () => reject(req.error);
  });
  const { listRecent } = await import('../src/lib/recent');
  const { listDrafts, putDraft } = await import('../src/lib/drafts');
  expect((await listRecent()).map((r) => r.name)).toEqual(['cu.pdf']);
  await putDraft({ id: 'draft:x', name: 'x.docx', kind: 'docx', size: 1, savedAt: 1, data: new ArrayBuffer(1) });
  expect((await listDrafts()).length).toBe(1);
});
