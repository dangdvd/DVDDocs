import 'fake-indexeddb/auto';
import { describe, expect, it } from 'vitest';
import { deleteDraft, listDrafts, loadDraft, putDraft } from '../src/lib/drafts';
import { addRecent, listRecent } from '../src/lib/recent';

describe('Bản nháp tự động lưu', () => {
  it('ghi, liệt kê (mới nhất trước), mở lại và xoá bản nháp', async () => {
    const data = new Uint8Array([1, 2, 3]).buffer;
    await putDraft({ id: 'draft:a', name: 'Báo cáo.docx', kind: 'docx', size: 3, savedAt: 100, data });
    await putDraft({ id: 'draft:b', name: 'Số liệu.xlsx', kind: 'xlsx', path: '/tmp/Số liệu.xlsx', size: 3, savedAt: 200, data });
    const list = await listDrafts();
    expect(list.map((d) => d.id)).toEqual(['draft:b', 'draft:a']);
    expect(list[0]).not.toHaveProperty('data');
    const f = await loadDraft('draft:b');
    expect(f).toMatchObject({ name: 'Số liệu.xlsx', kind: 'xlsx', path: '/tmp/Số liệu.xlsx', draftId: 'draft:b', restored: true });
    expect(new Uint8Array(f!.data)).toEqual(new Uint8Array([1, 2, 3]));
    await deleteDraft('draft:b');
    expect((await listDrafts()).map((d) => d.id)).toEqual(['draft:a']);
    expect(await loadDraft('draft:b')).toBeNull();
  });

  it('kho tệp gần đây vẫn hoạt động cùng cơ sở dữ liệu mới', async () => {
    await addRecent({ name: 'a.pdf', kind: 'pdf', data: new Uint8Array([9]).buffer });
    expect((await listRecent()).map((r) => r.name)).toContain('a.pdf');
  });
});
