// Mục lục (bookmark) của PDF: đọc cây mục lục và tìm trang đích của từng mục.

export interface OutlineItem { title: string; page: number | null; depth: number }

interface RawOutline { title: string; dest: string | unknown[] | null; items?: RawOutline[] }

/** Phần cần dùng của PDFDocumentProxy (để kiểm thử được mà không cần pdf.js thật). */
export interface OutlineSource {
  getOutline(): Promise<RawOutline[] | null>;
  getDestination(id: string): Promise<unknown[] | null>;
  getPageIndex(ref: { num: number; gen: number }): Promise<number>;
}

const MAX_ITEMS = 2000;

/** Đọc mục lục thành danh sách phẳng (kèm độ sâu), mỗi mục có chỉ số trang đích (từ 0) nếu tìm được. */
export async function readOutline(doc: OutlineSource): Promise<OutlineItem[]> {
  const root = await doc.getOutline().catch(() => null);
  if (!root?.length) return [];
  const out: OutlineItem[] = [];
  const pageOf = async (dest: RawOutline['dest']): Promise<number | null> => {
    try {
      const arr = typeof dest === 'string' ? await doc.getDestination(dest) : dest;
      const target = Array.isArray(arr) ? arr[0] : null;
      if (typeof target === 'number') return target; // một số tệp ghi thẳng số trang
      if (target && typeof target === 'object' && 'num' in target) return await doc.getPageIndex(target as { num: number; gen: number });
    } catch { /* đích hỏng: mục vẫn hiện nhưng không bấm được */ }
    return null;
  };
  const walk = async (items: RawOutline[], depth: number) => {
    for (const it of items) {
      if (out.length >= MAX_ITEMS) return;
      out.push({ title: (it.title ?? '').trim() || '(không tên)', page: await pageOf(it.dest), depth });
      if (it.items?.length) await walk(it.items, depth + 1);
    }
  };
  await walk(root, 0);
  return out;
}
