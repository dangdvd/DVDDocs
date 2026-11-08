// Định dạng có điều kiện và lọc dữ liệu: mô hình, cách tính, và chuyển đổi với ExcelJS.
import type ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { sheetParts } from './charts';
import type { Engine, Merge, Sheet } from './sheetModel';
import { formatValue, key, parseRange, rangeRef } from './sheetModel';

export type CfKind = 'gt' | 'lt' | 'ge' | 'le' | 'eq' | 'ne' | 'between' | 'text' | 'scale' | 'bar' | 'other';

export interface CfRule {
  range: Merge;
  kind: CfKind;
  /** Giá trị so sánh (số hoặc chữ); với "between" là cận dưới. */
  a?: string;
  /** Cận trên của "between". */
  b?: string;
  fill?: string;
  color?: string;
  /** Thang màu (2 hoặc 3 màu, từ nhỏ tới lớn) hoặc màu thanh dữ liệu (1 màu). */
  colors?: string[];
  /** Quy tắc gốc không hỗ trợ vẽ: giữ nguyên để ghi lại. */
  raw?: unknown;
}

export const CF_KINDS: [CfKind, string][] = [
  ['gt', 'Lớn hơn'], ['lt', 'Nhỏ hơn'], ['ge', 'Lớn hơn hoặc bằng'], ['le', 'Nhỏ hơn hoặc bằng'], ['eq', 'Bằng'], ['ne', 'Khác'],
  ['between', 'Trong khoảng'], ['text', 'Chứa chữ'], ['scale', 'Thang màu'], ['bar', 'Thanh dữ liệu'],
];

export interface Filter {
  /** Hàng tiêu đề (có nút lọc) và các cột của vùng lọc. Dữ liệu từ hàng r + 1 tới r1. */
  r: number; c0: number; c1: number; r1: number;
  /** Cột → các giá trị (chữ hiển thị) bị ẩn. */
  hide: Record<number, string[]>;
}

export interface CfLook { fill?: string; color?: string; bar?: { pct: number; color: string } }

const inRange = (m: Merge, r: number, c: number) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1;
const num = (v: unknown) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v)) ? Number(v) : null);

/** Trộn màu a → b theo t ∈ [0, 1]. */
function mix(a: string, b: string, t: number) {
  const p = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const x = p(a), y = p(b);
  return `#${x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('')}`;
}

/** Bộ tính định dạng có điều kiện cho một trang (ghi nhớ giá trị nhỏ/lớn nhất của từng vùng thang màu). */
export function cfEvaluator(sheet: Sheet, engine: Engine, si: number) {
  const rules = (sheet.cf ?? []).filter((x) => x.kind !== 'other');
  const extent = new Map<CfRule, { min: number; max: number }>();
  const ext = (rule: CfRule) => {
    let e = extent.get(rule);
    if (!e) {
      let min = Infinity, max = -Infinity;
      for (let r = rule.range.r0; r <= Math.min(rule.range.r1, rule.range.r0 + 5000); r++)
        for (let c = rule.range.c0; c <= rule.range.c1; c++) {
          const v = num(engine.valueAt(si, r, c));
          if (v !== null) { min = Math.min(min, v); max = Math.max(max, v); }
        }
      e = { min, max };
      extent.set(rule, e);
    }
    return e;
  };
  return (r: number, c: number): CfLook | null => {
    let out: CfLook | null = null;
    for (const rule of rules) {
      if (!inRange(rule.range, r, c)) continue;
      const v = engine.valueAt(si, r, c);
      const n = num(v);
      const a = num(rule.a), b = num(rule.b);
      let hit = false;
      switch (rule.kind) {
        case 'gt': hit = n !== null && a !== null && n > a; break;
        case 'lt': hit = n !== null && a !== null && n < a; break;
        case 'ge': hit = n !== null && a !== null && n >= a; break;
        case 'le': hit = n !== null && a !== null && n <= a; break;
        case 'eq': hit = a !== null ? n === a : String(v ?? '').toLowerCase() === String(rule.a ?? '').replace(/^"|"$/g, '').toLowerCase(); break;
        case 'ne': hit = a !== null ? n !== a : String(v ?? '').toLowerCase() !== String(rule.a ?? '').replace(/^"|"$/g, '').toLowerCase(); break;
        case 'between': hit = n !== null && a !== null && b !== null && n >= Math.min(a, b) && n <= Math.max(a, b); break;
        case 'text': hit = !!rule.a && formatValue(v, sheet.cells[key(r, c)]?.s?.numFmt).toLowerCase().includes(rule.a.toLowerCase()); break;
        case 'scale': {
          if (n === null) break;
          const { min, max } = ext(rule);
          const t = max > min ? (n - min) / (max - min) : 0.5;
          const cs = rule.colors?.length ? rule.colors : ['#F8696B', '#63BE7B'];
          const fill = cs.length >= 3 ? (t < 0.5 ? mix(cs[0], cs[1], t * 2) : mix(cs[1], cs[2], (t - 0.5) * 2)) : mix(cs[0], cs[cs.length - 1], t);
          out = { ...(out ?? {}), fill };
          continue;
        }
        case 'bar': {
          if (n === null) break;
          const { min, max } = ext(rule);
          const lo = Math.min(0, min);
          out = { ...(out ?? {}), bar: { pct: max > lo ? Math.max(0, Math.min(1, (n - lo) / (max - lo))) : 1, color: rule.colors?.[0] ?? '#638EC6' } };
          continue;
        }
      }
      if (hit) out = { ...(out ?? {}), fill: rule.fill ?? out?.fill, color: rule.color ?? out?.color };
    }
    return out;
  };
}

/** Các hàng bị ẩn do lọc. */
export function filterHidden(sheet: Sheet, engine: Engine, si: number, f: Filter): Record<number, boolean> {
  const hidden: Record<number, boolean> = {};
  const cols = Object.entries(f.hide).filter(([, vals]) => vals.length).map(([c, vals]) => [Number(c), new Set(vals)] as const);
  if (!cols.length) return hidden;
  for (let r = f.r + 1; r <= f.r1; r++) {
    for (const [c, vals] of cols) {
      const shown = formatValue(engine.valueAt(si, r, c), sheet.cells[key(r, c)]?.s?.numFmt);
      if (vals.has(shown)) { hidden[r] = true; break; }
    }
  }
  return hidden;
}

/** Các giá trị (chữ hiển thị) khác nhau trong một cột của vùng lọc, đã sắp xếp. */
export function columnValues(sheet: Sheet, engine: Engine, si: number, f: Filter, c: number): string[] {
  const set = new Set<string>();
  for (let r = f.r + 1; r <= f.r1; r++) set.add(formatValue(engine.valueAt(si, r, c), sheet.cells[key(r, c)]?.s?.numFmt));
  return [...set].sort((a, b) => {
    const x = Number(a.replace(/[^\d.-]/g, '')), y = Number(b.replace(/[^\d.-]/g, ''));
    return a && b && !Number.isNaN(x) && !Number.isNaN(y) && /\d/.test(a) && /\d/.test(b) ? x - y : a.localeCompare(b, 'vi');
  });
}

// ---------------- ExcelJS ----------------

const OPS: Record<string, CfKind> = {
  greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le', equal: 'eq', notEqual: 'ne', between: 'between',
};
const OPS_BACK = Object.fromEntries(Object.entries(OPS).map(([k, v]) => [v, k])) as Record<string, string>;
const argb = (c?: { argb?: string }) => (c?.argb && c.argb.length >= 6 ? `#${c.argb.slice(-6).toUpperCase()}` : undefined);
const toArgb = (hex?: string) => (hex ? `FF${hex.replace('#', '').toUpperCase()}` : undefined);

type XRule = Record<string, any>;

export function readCf(ws: ExcelJS.Worksheet): CfRule[] {
  const out: CfRule[] = [];
  const list = (ws as unknown as { conditionalFormattings?: { ref: string; rules: XRule[] }[] }).conditionalFormattings ?? [];
  for (const cf of list) {
    for (const ref of String(cf.ref).split(/\s+/).filter(Boolean)) {
      const range = parseRange(ref);
      if (!range) continue;
      for (const x of cf.rules ?? []) {
        const st = x.style ?? {};
        const fill = argb(st.fill?.bgColor) ?? argb(st.fill?.fgColor) ?? (st.fill ? '#FFC7CE' : undefined);
        const color = argb(st.font?.color) ?? (st.fill && !st.font ? undefined : st.font ? '#9C0006' : undefined);
        if (x.type === 'cellIs' && OPS[x.operator]) {
          out.push({ range, kind: OPS[x.operator], a: x.formulae?.[0]?.toString(), b: x.formulae?.[1]?.toString(), fill, color });
        } else if (x.type === 'containsText' && x.text) {
          out.push({ range, kind: 'text', a: String(x.text), fill, color });
        } else if (x.type === 'colorScale' && Array.isArray(x.color)) {
          out.push({ range, kind: 'scale', colors: x.color.map((c: { argb?: string }) => argb(c) ?? '#FFFFFF') });
        } else if (x.type === 'dataBar') {
          out.push({ range, kind: 'bar', colors: [argb(x.color) ?? '#638EC6'], raw: x });
        } else {
          out.push({ range, kind: 'other', raw: x });
        }
      }
    }
  }
  return out;
}

export function writeCf(ws: ExcelJS.Worksheet, rules: CfRule[] | undefined) {
  const out: { ref: string; rules: XRule[] }[] = [];
  let priority = 1;
  for (const r of rules ?? []) {
    let x: XRule;
    const style = { ...(r.fill ? { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: toArgb(r.fill) } } } : {}), ...(r.color ? { font: { color: { argb: toArgb(r.color) } } } : {}) };
    if (r.kind === 'other' || (r.kind === 'bar' && r.raw)) x = { ...(r.raw as XRule) };
    else if (r.kind === 'text') x = { type: 'containsText', operator: 'containsText', text: r.a ?? '', style };
    else if (r.kind === 'scale') {
      const cs = r.colors?.length ? r.colors : ['#F8696B', '#63BE7B'];
      x = {
        type: 'colorScale',
        cfvo: cs.length >= 3 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }],
        color: cs.map((c) => ({ argb: toArgb(c) })),
      };
    } else if (r.kind === 'bar') x = { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: toArgb(r.colors?.[0] ?? '#638EC6') } };
    else x = { type: 'cellIs', operator: OPS_BACK[r.kind], formulae: [r.a ?? '0', ...(r.kind === 'between' ? [r.b ?? '0'] : [])], style };
    x.priority = priority++;
    out.push({ ref: rangeRef(r.range), rules: [x] });
  }
  (ws as unknown as { conditionalFormattings: unknown }).conditionalFormattings = out;
}

// ---------------- Tiêu chí lọc trong <autoFilter> (ExcelJS chỉ ghi vùng lọc) ----------------

const xesc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const xunesc = (t: string) => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** Phần tử <autoFilter> đầy đủ: mỗi cột có giá trị bị ẩn ghi danh sách giá trị còn hiện (như Excel). */
export function autoFilterXml(f: Filter, sheet: Sheet, engine: Engine, si: number): string {
  let cols = '';
  for (const [c, hide] of Object.entries(f.hide)) {
    if (!hide.length) continue;
    const h = new Set(hide);
    const shown = columnValues(sheet, engine, si, f, Number(c)).filter((v) => !h.has(v));
    const blank = shown.includes('') ? ' blank="1"' : '';
    cols += `<filterColumn colId="${Number(c) - f.c0}"><filters${blank}>${shown.filter((v) => v !== '').map((v) => `<filter val="${xesc(v)}"/>`).join('')}</filters></filterColumn>`;
  }
  const ref = `${rangeRef({ r0: f.r, c0: f.c0, r1: f.r1, c1: f.c1 })}`;
  return cols ? `<autoFilter ref="${ref}">${cols}</autoFilter>` : `<autoFilter ref="${ref}"/>`;
}

/** Thay <autoFilter> trong các trang tính đã ghi bằng phần tử đầy đủ. */
export async function writeAutoFilters(buf: ArrayBuffer, filters: Map<string, string>): Promise<ArrayBuffer> {
  if (![...filters.values()].some((x) => x.includes('<filterColumn'))) return buf;
  const zip = await JSZip.loadAsync(buf);
  for (const part of await sheetParts(zip)) {
    const el = filters.get(part.name);
    if (!el?.includes('<filterColumn')) continue;
    const xml = await zip.file(part.path)?.async('string');
    if (!xml) continue;
    zip.file(part.path, xml.replace(/<autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/autoFilter>)/, el));
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}

/** Tiêu chí lọc đọc từ tệp: cột (tính từ đầu vùng) → giá trị còn hiện; `unsupported` khi có kiểu lọc khác (lọc theo điều kiện, top 10…). */
export interface FilterCriteria { shown: Record<number, { vals: string[]; blank: boolean }>; unsupported: boolean }

export async function readAutoFilters(data: ArrayBuffer): Promise<Map<string, FilterCriteria>> {
  const out = new Map<string, FilterCriteria>();
  const zip = await JSZip.loadAsync(data);
  for (const part of await sheetParts(zip)) {
    const xml = await zip.file(part.path)?.async('string');
    const af = xml && /<(?:\w+:)?autoFilter\b[^>]*>([\s\S]*?)<\/(?:\w+:)?autoFilter>/.exec(xml)?.[1];
    if (!af) continue;
    const fc: FilterCriteria = { shown: {}, unsupported: false };
    for (const m of af.matchAll(/<(?:\w+:)?filterColumn\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?filterColumn>/g)) {
      const col = Number(/colId="(\d+)"/.exec(m[1])?.[1] ?? -1);
      const fs = /<(?:\w+:)?filters\b([^>]*)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?filters>)/.exec(m[2]);
      if (col < 0 || !fs || /dateGroupItem/.test(fs[2] ?? '')) { fc.unsupported = true; continue; }
      const vals = [...(fs[2] ?? '').matchAll(/<(?:\w+:)?filter\b[^>]*\bval="([^"]*)"/g)].map((x) => xunesc(x[1]));
      fc.shown[col] = { vals, blank: /blank="(1|true)"/.test(fs[1]) };
    }
    out.set(part.name, fc);
  }
  return out;
}
