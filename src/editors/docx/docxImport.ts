// Bổ sung những định dạng mammoth bỏ qua khi đọc .docx: màu chữ, tô nền chữ, căn lề và giãn dòng của đoạn,
// màu nền ô bảng; đọc chữ ở đầu trang / chân trang.
// Cách làm: chèn ký tự đánh dấu (vùng dùng riêng của Unicode) vào chữ trong document.xml trước khi mammoth đọc,
// rồi đổi các dấu đó thành thuộc tính HTML sau khi có HTML.

const HL: Record<string, string> = {
  yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
  darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
  darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
};

const ALIGN: Record<string, string> = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', left: 'left', start: 'left' };

/** Chèn `marker` vào đầu chữ của thẻ <w:t> đầu tiên trong đoạn xml[from, to). Trả về xml mới (hoặc như cũ). */
function markFirstText(xml: string, from: number, to: number, marker: string): string {
  const re = /<w:t(\s[^>]*)?>/g;
  re.lastIndex = from;
  const m = re.exec(xml);
  if (!m || m.index >= to) return xml;
  const at = m.index + m[0].length;
  return xml.slice(0, at) + marker + xml.slice(at);
}

/** Vị trí kết thúc của phần tử bắt đầu tại `start` (có tính lồng nhau cùng tên). */
function elementEnd(xml: string, start: number, name: string): number {
  const re = new RegExp(`<${name}(?=[\\s>/])[^>]*?(/?)>|</${name}>`, 'g');
  re.lastIndex = start;
  let depth = 0;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[0].startsWith('</')) depth--;
    else if (m[1] !== '/') depth++;
    if (depth === 0) return m.index + m[0].length;
  }
  return xml.length;
}

/** Đánh dấu định dạng trong document.xml (chạy trước mammoth). */
export function markFormatting(xml: string): string {
  // 1) Từng đoạn chữ (<w:r>): màu chữ và tô nền chữ.
  xml = xml.replace(/<w:r(?=[\s>])[^>]*>([\s\S]*?)<\/w:r>/g, (run, body: string) => {
    const rpr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(body)?.[1] ?? '';
    const color = /<w:color\b[^>]*w:val="([0-9A-Fa-f]{6})"/.exec(rpr)?.[1];
    const hlName = /<w:highlight\b[^>]*w:val="(\w+)"/.exec(rpr)?.[1];
    const shd = /<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/.exec(rpr)?.[1];
    const hl = (hlName && HL[hlName]) || (shd && shd.toUpperCase() !== 'FFFFFF' ? shd : undefined);
    const useColor = color && color.toUpperCase() !== '000000';
    if (!useColor && !hl) return run;
    const open = `${useColor ? `${color!.toUpperCase()}` : ''}${hl ? `${hl.toUpperCase()}` : ''}`;
    const close = `${hl ? '' : ''}${useColor ? '' : ''}`;
    return run.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g, (_m, a: string, t: string, b: string) => `${a}${open}${t}${close}${b}`);
  });
  // 2) Đoạn văn: căn lề và giãn dòng (chèn vào chữ đầu tiên của đoạn).
  const paras: { at: number; end: number; marker: string }[] = [];
  for (const m of xml.matchAll(/<w:p(?=[\s>])[^>]*>\s*<w:pPr>([\s\S]*?)<\/w:pPr>/g)) {
    const ppr = m[1];
    const jc = /<w:jc\b[^>]*w:val="(\w+)"/.exec(ppr)?.[1];
    const align = jc ? ALIGN[jc] ?? '' : '';
    const sp = /<w:spacing\b[^>]*>/.exec(ppr)?.[0] ?? '';
    const line = Number(/w:line="(\d+)"/.exec(sp)?.[1] ?? 0);
    const rule = /w:lineRule="(\w+)"/.exec(sp)?.[1] ?? 'auto';
    const lh = line && rule === 'auto' && line !== 240 ? +(line / 240).toFixed(2) : 0;
    if (!align && !lh) continue;
    const start = m.index!;
    paras.push({ at: start + m[0].length, end: elementEnd(xml, start, 'w:p'), marker: `${align}|${lh || ''}` });
  }
  // 3) Ô bảng: màu nền.
  const cells: { at: number; end: number; marker: string }[] = [];
  for (const m of xml.matchAll(/<w:tc(?=[\s>])[^>]*>\s*<w:tcPr>([\s\S]*?)<\/w:tcPr>/g)) {
    const fill = /<w:shd\b[^>]*w:fill="([0-9A-Fa-f]{6})"/.exec(m[1])?.[1];
    if (!fill || fill.toUpperCase() === 'FFFFFF') continue;
    cells.push({ at: m.index! + m[0].length, end: elementEnd(xml, m.index!, 'w:tc'), marker: `${fill.toUpperCase()}` });
  }
  // Chèn từ cuối lên để vị trí phía trước không đổi.
  for (const e of [...paras, ...cells].sort((a, b) => b.at - a.at)) xml = markFirstText(xml, e.at, e.end, e.marker);
  return xml;
}

/** Đổi các dấu đã chèn thành HTML (chạy sau mammoth). */
export function formattingToHtml(html: string): string {
  html = html
    .replace(/([0-9A-F]{6})/g, '<span style="color: #$1">').replace(//g, '</span>')
    .replace(/([0-9A-F]{6})/g, '<mark data-color="#$1" style="background-color: #$1">').replace(//g, '</mark>');
  // Thuộc tính đoạn: tìm dấu nằm trong đoạn (<p>/<h1-6>) gần nhất phía trước.
  html = html.replace(/<(p|h[1-6])((?:\s[^>]*)?)>((?:(?!<\/?(?:p|h[1-6])[\s>])[\s\S])*?)(\w*)\|([\d.]*)/g,
    (_m, tag: string, attrs: string, before: string, align: string, lh: string) => {
      const st = [align && `text-align: ${align}`, lh && `line-height: ${lh}`].filter(Boolean).join('; ');
      return `<${tag}${attrs}${st ? ` style="${st}"` : ''}>${before}`;
    });
  html = html.replace(/<(td|th)((?:\s[^>]*)?)>((?:(?!<\/?t[dh][\s>])[\s\S])*?)([0-9A-F]{6})/g,
    (_m, tag: string, attrs: string, before: string, fill: string) => `<${tag}${attrs} style="background-color: #${fill}">${before}`);
  // Dấu còn sót (ví dụ đoạn trong danh sách): bỏ đi.
  return html.replace(/[^]*|[^]*/g, '');
}

/** Chữ thường của một phần đầu/chân trang (bỏ đoạn chứa số trang tự động). */
export function headerFooterText(xml: string): string {
  const out: string[] = [];
  for (const m of xml.matchAll(/<w:p(?=[\s>])[\s\S]*?<\/w:p>/g)) {
    const p = m[0];
    if (/\b(PAGE|NUMPAGES|SECTIONPAGES)\b/.test(p)) continue;
    // Bỏ kết quả của trường (field) để không lấy số đã tính sẵn.
    const clean = p.replace(/<w:fldChar[^>]*w:fldCharType="separate"[^>]*\/>[\s\S]*?<w:fldChar[^>]*w:fldCharType="end"[^>]*\/>/g, '')
      .replace(/<w:instrText[\s\S]*?<\/w:instrText>/g, '');
    const text = [...clean.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g)]
      .map((t) => (t[0] === '<w:tab/>' ? '\t' : t[1])).join('')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
    out.push(text);
  }
  while (out.length && !out[out.length - 1].trim()) out.pop();
  while (out.length && !out[0].trim()) out.shift();
  return out.join('\n');
}
