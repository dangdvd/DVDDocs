// Theo dõi thay đổi (giống "Track Changes" của Word): khi bật, chữ gõ thêm được đánh dấu "chèn",
// chữ bị xoá không mất mà được đánh dấu "xoá" (gạch ngang đỏ). Có thể chấp nhận hoặc từ chối từng chỗ hay tất cả.
import { Extension, Mark, mergeAttributes } from '@tiptap/core';
import type { Node as PMNode, Mark as PMMark, MarkType } from '@tiptap/pm/model';
import { EditorState, Plugin, PluginKey, TextSelection, Transaction } from '@tiptap/pm/state';
import { ReplaceStep } from '@tiptap/pm/transform';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    trackChanges: {
      setTrackChanges: (on: boolean) => ReturnType;
      acceptAllChanges: () => ReturnType;
      rejectAllChanges: () => ReturnType;
      acceptChange: () => ReturnType;
      rejectChange: () => ReturnType;
      goToChange: (dir: 1 | -1) => ReturnType;
    };
  }
}

const changeAttrs = () => ({
  author: { default: null, parseHTML: (el: HTMLElement) => el.getAttribute('data-author'), renderHTML: (a: Record<string, unknown>) => (a.author ? { 'data-author': a.author } : {}) },
  date: { default: null, parseHTML: (el: HTMLElement) => el.getAttribute('data-date'), renderHTML: (a: Record<string, unknown>) => (a.date ? { 'data-date': a.date } : {}) },
});

/** Chữ được chèn khi đang theo dõi. */
export const Insertion = Mark.create({
  name: 'insertion',
  // Không tự lan sang chữ gõ kế bên: khi đang theo dõi, chữ mới được gắn dấu ở appendTransaction.
  inclusive: false,
  excludes: 'deletion',
  priority: 1100,
  addAttributes: changeAttrs,
  parseHTML: () => [{ tag: 'ins[data-tc]', priority: 60 }],
  renderHTML: ({ HTMLAttributes }) => ['ins', mergeAttributes(HTMLAttributes, { 'data-tc': '', class: 'tc-ins', title: titleOf('Đã chèn', HTMLAttributes) }), 0],
});

/** Chữ bị xoá khi đang theo dõi (vẫn còn trong tài liệu cho tới khi chấp nhận). */
export const Deletion = Mark.create({
  name: 'deletion',
  inclusive: false,
  excludes: 'insertion',
  priority: 1100,
  addAttributes: changeAttrs,
  parseHTML: () => [{ tag: 'del[data-tc]', priority: 60 }],
  renderHTML: ({ HTMLAttributes }) => ['del', mergeAttributes(HTMLAttributes, { 'data-tc': '', class: 'tc-del', title: titleOf('Đã xoá', HTMLAttributes) }), 0],
});

function titleOf(what: string, a: Record<string, unknown>) {
  const who = a['data-author'] ? ` bởi ${a['data-author']}` : '';
  const d = a['data-date'] ? new Date(String(a['data-date'])) : null;
  const when = d && !Number.isNaN(d.getTime()) ? ` lúc ${d.toLocaleString('vi-VN')}` : '';
  return `${what}${who}${when}`;
}

export const trackKey = new PluginKey<{ on: boolean }>('dvd-track');
const SKIP = 'dvd-track-skip';

export function currentAuthor() {
  try { return localStorage.getItem('dvd-author') || 'Người dùng'; } catch { return 'Người dùng'; }
}

export interface Change { kind: 'insertion' | 'deletion'; from: number; to: number }

/** Mọi đoạn chèn/xoá liền nhau trong tài liệu. */
export function listChanges(doc: PMNode): Change[] {
  const out: Change[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    for (const kind of ['insertion', 'deletion'] as const) {
      if (!node.marks.some((m) => m.type.name === kind)) continue;
      const last = out[out.length - 1];
      if (last && last.kind === kind && last.to === pos) last.to = pos + node.nodeSize;
      else out.push({ kind, from: pos, to: pos + node.nodeSize });
    }
    return false;
  });
  return out;
}

/** Áp dụng quyết định cho các thay đổi: chấp nhận (giữ chèn, bỏ xoá) hoặc từ chối (bỏ chèn, giữ xoá). */
function resolve(tr: Transaction, changes: Change[], accept: boolean, ins: MarkType, del: MarkType) {
  // Làm từ cuối lên để vị trí phía trước không bị lệch.
  for (const c of [...changes].sort((a, b) => b.from - a.from)) {
    const remove = (c.kind === 'insertion') !== accept;
    if (remove) tr.delete(c.from, c.to);
    else tr.removeMark(c.from, c.to, c.kind === 'insertion' ? ins : del);
  }
  return tr.setMeta(SKIP, true);
}

const hasMark = (node: PMNode, t: MarkType) => node.marks.some((m) => m.type === t);

/** Có chữ nào (không phải chữ mình vừa chèn) trong vùng [from, to) của doc không. */
function hasTrackableText(doc: PMNode, from: number, to: number, ins: MarkType) {
  let found = false;
  doc.nodesBetween(from, to, (n) => {
    if (found) return false;
    if (n.isText && !hasMark(n, ins)) found = true;
    else if (n.isLeaf && !n.isText && n.isInline) found = true; // ảnh, xuống dòng mềm
    return !found;
  });
  return found;
}

function track(tr: Transaction, oldState: EditorState, newState: EditorState): Transaction | null {
  const schema = newState.schema;
  const ins = schema.marks.insertion, del = schema.marks.deletion;
  // Dùng lại dấu của đoạn thay đổi liền kề (cùng người sửa) để các lần gõ nối tiếp gộp thành một thay đổi.
  const mark = (t: MarkType, doc?: PMNode, a?: number, b?: number): PMMark => {
    const author = currentAuthor();
    if (doc) {
      for (const n of [a !== undefined ? doc.resolve(a).nodeBefore : null, b !== undefined ? doc.resolve(b).nodeAfter : null]) {
        const m = n?.marks.find((x) => x.type === t && x.attrs.author === author);
        if (m) return m;
      }
    }
    return t.create({ author, date: new Date().toISOString() });
  };
  const steps = tr.steps;
  if (!steps.length || !steps.every((s) => s instanceof ReplaceStep)) return null;

  if (steps.length === 1) {
    const step = steps[0] as ReplaceStep & { from: number; to: number; slice: import('@tiptap/pm/model').Slice };
    const before = oldState.doc;
    const { from, to, slice } = step;
    const deletes = to > from && hasTrackableText(before, from, to, ins);
    const out = newState.tr;
    if (!deletes) {
      // Chỉ chèn (hoặc xoá chữ do chính mình chèn): đánh dấu phần mới là "chèn".
      const end = from + slice.size;
      if (slice.size === 0) return null;
      let need = false;
      newState.doc.nodesBetween(from, end, (n) => { if (n.isInline && !hasMark(n, ins)) need = true; });
      if (!need) return null;
      out.removeMark(from, end, del).addMark(from, end, mark(ins, newState.doc, from, end));
      return out.setMeta(SKIP, true).setMeta('addToHistory', tr.getMeta('addToHistory'));
    }
    // Có xoá chữ gốc: hoàn lại bước xoá, đánh dấu "xoá", rồi chèn phần mới phía sau.
    out.step(step.invert(before));
    let insEnd = to;
    if (slice.size) {
      const size0 = out.doc.content.size;
      out.replace(to, to, slice);
      insEnd = to + (out.doc.content.size - size0);
      out.removeMark(to, insEnd, del).addMark(to, insEnd, mark(ins, out.doc, undefined, insEnd));
    }
    // Chữ do chính mình chèn trong vùng xoá thì xoá hẳn; còn lại gắn dấu "xoá".
    const own: [number, number][] = [];
    out.doc.nodesBetween(from, to, (n, pos) => {
      if (n.isText && hasMark(n, ins)) own.push([Math.max(pos, from), Math.min(pos + n.nodeSize, to)]);
    });
    out.addMark(from, to, mark(del, out.doc, from, to));
    let shift = 0;
    for (const [a, b] of own.reverse()) { out.delete(a, b); shift += b - a; }
    const delEnd = to - shift;
    insEnd -= shift;
    // Đặt con trỏ: Backspace → trước phần bị xoá; Delete/chọn vùng → sau phần bị xoá (hoặc sau phần mới gõ).
    const sel = oldState.selection;
    let caret = insEnd;
    if (!slice.size && sel.empty && sel.head === to) caret = from;
    out.setSelection(TextSelection.create(out.doc, Math.min(caret, out.doc.content.size), Math.min(caret, out.doc.content.size)));
    void delEnd;
    return out.setMeta(SKIP, true);
  }

  // Nhiều bước (dán phức tạp…): chỉ đánh dấu phần chèn, không giữ phần bị xoá.
  const out = newState.tr;
  let changed = false;
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1);
    map.forEach((_a, _b, start, end) => {
      const a = rest.map(start, -1), b = rest.map(end, 1);
      if (b > a) { out.removeMark(a, b, del).addMark(a, b, mark(ins)); changed = true; }
    });
  });
  return changed ? out.setMeta(SKIP, true) : null;
}

export const TrackChanges = Extension.create({
  name: 'trackChanges',
  addExtensions: () => [Insertion, Deletion],
  addProseMirrorPlugins() {
    return [new Plugin({
      key: trackKey,
      state: {
        init: () => ({ on: false }),
        apply: (tr, v) => (tr.getMeta(trackKey) as { on: boolean } | undefined) ?? v,
      },
      appendTransaction(trs, oldState, newState) {
        if (trs.length !== 1) return null;
        const tr = trs[0];
        if (!tr.docChanged || tr.getMeta(SKIP) || tr.getMeta('history$') || tr.getMeta(trackKey)) return null;
        if (!trackKey.getState(newState)?.on) return null;
        // Nạp lại cả tài liệu (setContent): không theo dõi.
        const first = tr.steps[0] as ReplaceStep & { from?: number; to?: number };
        if (tr.steps.length === 1 && first.from === 0 && first.to === oldState.doc.content.size) return null;
        return track(tr, oldState, newState);
      },
    })];
  },
  addKeyboardShortcuts() {
    // Tự xử lý Backspace/Delete khi đang theo dõi: bỏ qua chữ đã đánh dấu xoá và xoá trọn một ký tự (kể cả dấu tiếng Việt tổ hợp).
    const del = (dir: -1 | 1) => () => {
      // Đọc ngay vùng chọn của trình duyệt (vừa bấm/chạm) trước khi xử lý, tránh dùng vùng chọn cũ.
      (this.editor.view as unknown as { domObserver?: { flush(): void } }).domObserver?.flush();
      const { state, view } = this.editor;
      if (!trackKey.getState(state)?.on) return false;
      const sel = state.selection;
      if (!sel.empty) return false;
      const $p = sel.$from;
      const parent = $p.parent;
      if (!parent.isTextblock) return false;
      const isDel = (n: PMNode | null | undefined) => !!n?.marks.some((m) => m.type.name === 'deletion');
      let pos = sel.from;
      // Nhảy qua phần đã bị xoá liền kề.
      for (;;) {
        const $q = state.doc.resolve(pos);
        const n = dir < 0 ? $q.nodeBefore : $q.nodeAfter;
        if (!n || !isDel(n)) break;
        pos += dir * (n.isText ? Math.min(n.nodeSize, dir < 0 ? $q.textOffset || n.nodeSize : n.nodeSize - $q.textOffset) : n.nodeSize);
      }
      const $q = state.doc.resolve(pos);
      const n = dir < 0 ? $q.nodeBefore : $q.nodeAfter;
      if (!n) {
        // Đầu/cuối đoạn: chỉ cho nối với đoạn văn liền kề (không theo dõi được dấu xuống đoạn);
        // nếu bên kia là bảng, ảnh… thì không làm gì để tránh xoá cả khối.
        if (pos !== sel.from) { view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, pos))); return true; }
        const d = $q.depth;
        const idx = $q.index(d - 1);
        const sib = dir < 0 ? (idx > 0 ? $q.node(d - 1).child(idx - 1) : null) : (idx + 1 < $q.node(d - 1).childCount ? $q.node(d - 1).child(idx + 1) : null);
        return !(sib && sib.isTextblock);
      }
      let len = 1;
      if (n.isText) {
        const text = n.text!;
        const off = dir < 0 ? ($q.textOffset || text.length) : $q.textOffset;
        // Một "ký tự" theo cách người dùng thấy: cặp thay thế UTF-16 và các dấu tổ hợp đi kèm.
        if (dir < 0) {
          let i = off - 1;
          while (i > 0 && (/\p{M}/u.test(text[i]) || /[\uDC00-\uDFFF]/.test(text[i]))) i--;
          len = off - i;
        } else {
          let i = off + 1;
          if (/[\uD800-\uDBFF]/.test(text[off])) i++;
          while (i < text.length && /\p{M}/u.test(text[i])) i++;
          len = i - off;
        }
      }
      const from = dir < 0 ? pos - len : pos, to = dir < 0 ? pos : pos + len;
      const tr = state.tr.setSelection(TextSelection.create(state.doc, dir < 0 ? to : from)).delete(from, to);
      view.dispatch(tr);
      return true;
    };
    return { Backspace: del(-1), Delete: del(1) };
  },
  addCommands() {
    const range = (state: EditorState, kinds?: Change['kind'][]) => {
      const { from, to } = state.selection;
      return listChanges(state.doc).filter((c) => (!kinds || kinds.includes(c.kind)) && c.from <= to && c.to >= from);
    };
    return {
      setTrackChanges: (on) => ({ tr, dispatch }) => { if (dispatch) tr.setMeta(trackKey, { on }); return true; },
      acceptAllChanges: () => ({ state, tr, dispatch }) => {
        const all = listChanges(state.doc);
        if (!all.length) return false;
        if (dispatch) resolve(tr, all, true, state.schema.marks.insertion, state.schema.marks.deletion);
        return true;
      },
      rejectAllChanges: () => ({ state, tr, dispatch }) => {
        const all = listChanges(state.doc);
        if (!all.length) return false;
        if (dispatch) resolve(tr, all, false, state.schema.marks.insertion, state.schema.marks.deletion);
        return true;
      },
      acceptChange: () => ({ state, tr, dispatch }) => {
        const cs = range(state);
        if (!cs.length) return false;
        if (dispatch) resolve(tr, cs, true, state.schema.marks.insertion, state.schema.marks.deletion);
        return true;
      },
      rejectChange: () => ({ state, tr, dispatch }) => {
        const cs = range(state);
        if (!cs.length) return false;
        if (dispatch) resolve(tr, cs, false, state.schema.marks.insertion, state.schema.marks.deletion);
        return true;
      },
      goToChange: (dir) => ({ state, tr, dispatch }) => {
        const all = listChanges(state.doc);
        if (!all.length) return false;
        const { from, to } = state.selection;
        const next = dir > 0 ? all.find((c) => c.from >= to) ?? all[0] : [...all].reverse().find((c) => c.to <= from) ?? all[all.length - 1];
        if (dispatch) tr.setSelection(TextSelection.create(tr.doc, next.from, next.to)).scrollIntoView();
        return true;
      },
    };
  },
});

/**
 * Chuẩn bị document.xml cho mammoth: mammoth bỏ chữ bị xoá và không đánh dấu chữ chèn,
 * nên ta bọc chữ của từng đoạn chèn/xoá bằng ký tự riêng rồi đổi thành <ins>/<del> sau khi chuyển HTML.
 */
export function markTrackedRuns(xml: string): string {
  if (!/<w:(ins|del)\b/.test(xml)) return xml;
  const wrap = (body: string, open: string, close: string, textTag: RegExp) =>
    body.replace(textTag, (_m, attrs: string | undefined, text: string) => `<w:t${attrs ?? ''}>${open}${text}${close}</w:t>`);
  // Bỏ lớp bọc <w:ins>/<w:del> (mammoth bỏ qua cả khối <w:del>), giữ các đoạn chữ bên trong.
  // Bỏ qua dạng tự đóng <w:ins .../> (đánh dấu ký tự xuống đoạn).
  xml = xml.replace(/<w:ins\b[^>]*?(?<!\/)>([\s\S]*?)<\/w:ins>/g, (_m, body: string) =>
    wrap(body, '', '', /<w:t(\s[^>]*)?>([^<]*)<\/w:t>/g));
  xml = xml.replace(/<w:del\b[^>]*?(?<!\/)>([\s\S]*?)<\/w:del>/g, (_m, body: string) =>
    wrap(body, '', '', /<w:delText(\s[^>]*)?>([^<]*)<\/w:delText>/g));
  return xml;
}

/** Đổi các ký tự đánh dấu (sau khi mammoth chuyển sang HTML) thành thẻ <ins>/<del>. */
export function trackedMarkersToHtml(html: string): string {
  return html
    .replace(//g, '<ins data-tc="">').replace(//g, '</ins>')
    .replace(//g, '<del data-tc="">').replace(//g, '</del>');
}
