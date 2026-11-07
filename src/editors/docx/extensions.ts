// Phần mở rộng TipTap tự viết: cỡ chữ, ngắt trang, tô sáng kết quả tìm kiếm.
import { Extension, Node } from '@tiptap/core';
import type { Node as PMNode } from '@tiptap/pm/model';
import { NodeSelection, Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    fontSize: {
      setFontSize: (size: string) => ReturnType;
      unsetFontSize: () => ReturnType;
    };
    pageBreak: {
      setPageBreak: () => ReturnType;
    };
  }
}

/** Cỡ chữ (ví dụ "12pt") lưu trong dấu textStyle, giống cách màu chữ được lưu. */
export const FontSize = Extension.create({
  name: 'fontSize',
  addGlobalAttributes() {
    return [{
      types: ['textStyle'],
      attributes: {
        fontSize: {
          default: null,
          parseHTML: (el) => (el as HTMLElement).style.fontSize || null,
          renderHTML: (attrs) => (attrs.fontSize ? { style: `font-size: ${attrs.fontSize}` } : {}),
        },
      },
    }];
  },
  addCommands() {
    return {
      setFontSize: (size) => ({ chain }) => chain().setMark('textStyle', { fontSize: size }).run(),
      unsetFontSize: () => ({ chain }) => chain().setMark('textStyle', { fontSize: null }).removeEmptyTextStyle().run(),
    };
  },
});

/** Ngắt trang: khi in/xuất PDF và khi lưu .docx, nội dung sau nó sang trang mới. */
export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  parseHTML: () => [{ tag: 'hr.page-break', priority: 60 }, { tag: 'div[data-page-break]' }],
  renderHTML: () => ['div', { 'data-page-break': '', class: 'page-break' }],
  addCommands() {
    return {
      // Giống lệnh chèn đường kẻ của TipTap: chèn rồi đặt con trỏ vào đoạn phía sau để gõ tiếp không ghi đè.
      setPageBreak: () => ({ chain, state }) => {
        const { selection } = state;
        const { $from, $to } = selection;
        const c = chain();
        if ($from.parentOffset === 0) c.insertContentAt({ from: Math.max($from.pos - 1, 0), to: $to.pos }, { type: this.name });
        else if (selection instanceof NodeSelection) c.insertContentAt($to.pos, { type: this.name });
        else c.insertContent({ type: this.name });
        return c.command(({ tr, dispatch }) => {
          if (!dispatch) return true;
          const { $to: end } = tr.selection;
          if (end.nodeAfter?.isTextblock) tr.setSelection(TextSelection.create(tr.doc, end.pos + 1));
          else {
            // Sau ngắt trang là bảng/ảnh hoặc hết tài liệu: thêm một đoạn trống để gõ, tránh ghi đè lên khối đó.
            const para = state.schema.nodes.paragraph.create();
            tr.insert(end.pos, para);
            tr.setSelection(TextSelection.create(tr.doc, end.pos + 1));
          }
          tr.scrollIntoView();
          return true;
        }).run();
      },
    };
  },
  addKeyboardShortcuts() {
    return { 'Mod-Enter': () => this.editor.commands.setPageBreak() };
  },
});

export interface Match { from: number; to: number }

const fold = (t: string) => t.toLocaleLowerCase('vi');

/** Tìm mọi chỗ xuất hiện (không phân biệt hoa thường), kể cả khi cụm từ trải qua nhiều đoạn định dạng khác nhau. */
export function findInDoc(doc: PMNode, query: string): Match[] {
  const q = fold(query.normalize('NFC'));
  const out: Match[] = [];
  if (!q) return out;
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    let text = '';
    const map: number[] = [];
    node.forEach((child, offset) => {
      if (child.isText) {
        for (let i = 0; i < child.text!.length; i++) map.push(pos + 1 + offset + i);
        text += child.text;
      } else {
        map.push(pos + 1 + offset);
        text += '￼';
      }
    });
    const t = fold(text);
    let i = t.indexOf(q);
    while (i >= 0) {
      out.push({ from: map[i], to: map[i + q.length - 1] + 1 });
      i = t.indexOf(q, i + q.length);
    }
    return false;
  });
  return out;
}

export const searchKey = new PluginKey<{ query: string; index: number }>('dvd-search');

/** Tô sáng các kết quả tìm kiếm; kết quả đang chọn đậm hơn. Cập nhật bằng tr.setMeta(searchKey, {query, index}). */
export const SearchHighlight = Extension.create({
  name: 'searchHighlight',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: searchKey,
      state: {
        init: () => ({ query: '', index: 0 }),
        apply: (tr, value) => tr.getMeta(searchKey) ?? value,
      },
      props: {
        decorations(state) {
          const s = searchKey.getState(state);
          if (!s?.query) return null;
          const found = findInDoc(state.doc, s.query);
          return DecorationSet.create(state.doc, found.map((m, i) =>
            Decoration.inline(m.from, m.to, { class: i === s.index ? 'search-hit cur' : 'search-hit' })));
        },
      },
    })];
  },
});
