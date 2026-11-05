import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import * as pdfjs from 'pdfjs-dist';
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { OpenedFile, pickFile, saveFile } from '../../lib/fileio';
import { Annot, applyAnnotations, deletePage, insertBlankPage, mergePdf, movePage, rotatePage } from './pdfOps';
import { SignaturePad } from './SignaturePad';
import { EditorShell, ToolButton, Sep } from '../../components/EditorShell';
import { uid } from '../../lib/uid';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

// Phông chữ dùng khi xem trước chữ chèn vào PDF, trùng với phông được nhúng khi lưu (hỗ trợ tiếng Việt).
if (typeof FontFace !== 'undefined') {
  const face = new FontFace('DejaVu Sans', `url(${new URL('fonts/DejaVuSans.ttf', document.baseURI).href})`);
  face.load().then((f) => document.fonts.add(f)).catch(() => {});
}

type Tool = 'select' | 'text' | 'highlight' | 'rect' | 'whiteout' | 'ink' | 'place';
type Size = { w: number; h: number };

const MIN_ZOOM = 0.2, MAX_ZOOM = 5;
const clampZoom = (z: number) => +Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z)).toFixed(3);

interface Snapshot { bytes: Uint8Array; annots: Annot[] }

const TOOLS: { id: Tool; label: string; icon: string; hint: string }[] = [
  { id: 'select', label: 'Chọn', icon: '↖', hint: 'Chọn, kéo để di chuyển, Delete để xoá' },
  { id: 'text', label: 'Chữ', icon: 'T', hint: 'Bấm vào trang để thêm chữ' },
  { id: 'highlight', label: 'Tô sáng', icon: '▮', hint: 'Kéo để tô sáng' },
  { id: 'rect', label: 'Khung', icon: '▭', hint: 'Kéo để vẽ khung' },
  { id: 'whiteout', label: 'Che trắng', icon: '⬜', hint: 'Kéo để che nội dung cũ, rồi gõ chữ mới lên' },
  { id: 'ink', label: 'Bút vẽ', icon: '✎', hint: 'Kéo để vẽ tự do' },
];

export default function PdfEditor({ file, onClose }: { file: OpenedFile; onClose: () => void }) {
  const [bytes, setBytes] = useState(() => new Uint8Array(file.data));
  const [doc, setDoc] = useState<PDFDocumentProxy | null>(null);
  const [sizes, setSizes] = useState<Size[]>([]);
  const [annots, setAnnots] = useState<Annot[]>([]);
  const [history, setHistory] = useState<Snapshot[]>([]);
  const [future, setFuture] = useState<Snapshot[]>([]);
  const [tool, setTool] = useState<Tool>('select');
  const [color, setColor] = useState('#d1242f');
  const [fontSize, setFontSize] = useState(14);
  const [zoom, setZoom] = useState(1.2);
  const [selected, setSelected] = useState<string | null>(null);
  const [pending, setPending] = useState<{ dataUrl: string; w: number; h: number } | null>(null);
  const [showSign, setShowSign] = useState(false);
  const [showThumbs, setShowThumbs] = useState(() => window.innerWidth > 800);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);
  const [path, setPath] = useState(file.path);
  const [query, setQuery] = useState('');
  const [hits, setHits] = useState<number[] | null>(null);
  const [hitIdx, setHitIdx] = useState(0);
  const scrollRef = useRef<HTMLDivElement>(null);
  const pagesRef = useRef<HTMLDivElement>(null);
  const fitted = useRef(false);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  /** Điểm nội dung cần giữ nguyên dưới ngón tay/con trỏ sau khi đổi tỉ lệ. */
  const anchor = useRef<{ cx: number; cy: number; vx: number; vy: number; from: number } | null>(null);

  const fitZoom = (s: Size[] = sizes) => {
    const w = scrollRef.current?.clientWidth ?? window.innerWidth;
    const maxW = Math.max(...s.map((p) => p.w), 1);
    const pad = w < 700 ? 8 : 40;
    return clampZoom((w - pad) / maxW);
  };

  /** Đổi tỉ lệ, giữ điểm (vx, vy) trong khung nhìn đứng yên (mặc định là tâm khung). */
  const zoomTo = (next: number, vx?: number, vy?: number) => {
    const el = scrollRef.current;
    const z = clampZoom(next);
    if (el) {
      const px = vx ?? el.clientWidth / 2, py = vy ?? el.clientHeight / 2;
      anchor.current = { cx: el.scrollLeft + px, cy: el.scrollTop + py, vx: px, vy: py, from: zoomRef.current };
    }
    setZoom(z);
  };

  useLayoutEffect(() => {
    const el = scrollRef.current, a = anchor.current;
    if (!el || !a) return;
    anchor.current = null;
    const k = zoom / a.from;
    el.scrollLeft = a.cx * k - a.vx;
    el.scrollTop = a.cy * k - a.vy;
  }, [zoom]);

  // Chụm hai ngón để phóng to/thu nhỏ (điện thoại, máy tính bảng) và Ctrl + lăn chuột (máy tính).
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    let pinch: { d0: number; mx: number; my: number; scale: number } | null = null;
    const dist = (t: TouchList) => Math.hypot(t[0].clientX - t[1].clientX, t[0].clientY - t[1].clientY);
    const mid = (t: TouchList) => {
      const r = el.getBoundingClientRect();
      return { x: (t[0].clientX + t[1].clientX) / 2 - r.left, y: (t[0].clientY + t[1].clientY) / 2 - r.top };
    };
    const onStart = (e: TouchEvent) => {
      if (e.touches.length !== 2) return;
      const m = mid(e.touches);
      pinch = { d0: dist(e.touches), mx: m.x, my: m.y, scale: 1 };
      const pages = pagesRef.current;
      if (pages) pages.style.transformOrigin = `${el.scrollLeft + m.x}px ${el.scrollTop + m.y}px`;
    };
    const onMove = (e: TouchEvent) => {
      if (!pinch || e.touches.length !== 2) return;
      e.preventDefault();
      const z0 = zoomRef.current;
      pinch.scale = clampZoom(z0 * (dist(e.touches) / pinch.d0)) / z0;
      if (pagesRef.current) pagesRef.current.style.transform = `scale(${pinch.scale})`;
    };
    const onEnd = () => {
      if (!pinch) return;
      const p = pinch;
      pinch = null;
      if (pagesRef.current) pagesRef.current.style.transform = '';
      if (Math.abs(p.scale - 1) > 0.01) zoomTo(zoomRef.current * p.scale, p.mx, p.my);
    };
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      const r = el.getBoundingClientRect();
      zoomTo(zoomRef.current * Math.exp(-e.deltaY * 0.002), e.clientX - r.left, e.clientY - r.top);
    };
    el.addEventListener('touchstart', onStart, { passive: true });
    el.addEventListener('touchmove', onMove, { passive: false });
    el.addEventListener('touchend', onEnd);
    el.addEventListener('touchcancel', onEnd);
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => {
      el.removeEventListener('touchstart', onStart);
      el.removeEventListener('touchmove', onMove);
      el.removeEventListener('touchend', onEnd);
      el.removeEventListener('touchcancel', onEnd);
      el.removeEventListener('wheel', onWheel);
    };
  }, []);

  // Nạp tài liệu vào pdf.js mỗi khi nội dung thay đổi.
  useEffect(() => {
    let cancelled = false;
    const task = pdfjs.getDocument({ data: bytes.slice() });
    task.promise
      .then(async (d) => {
        const s: Size[] = [];
        for (let i = 1; i <= d.numPages; i++) {
          const vp = (await d.getPage(i)).getViewport({ scale: 1 });
          s.push({ w: vp.width, h: vp.height });
        }
        if (cancelled) return d.destroy();
        if (!fitted.current) {
          fitted.current = true;
          setZoom(fitZoom(s));
        }
        setSizes(s);
        setDoc((old) => {
          old?.destroy();
          return d;
        });
      })
      .catch((e) => !cancelled && setError(`Không mở được PDF: ${e.message ?? e}`));
    return () => {
      cancelled = true;
    };
  }, [bytes]);

  const snapshot = useCallback(() => {
    setHistory((h) => [...h.slice(-29), { bytes, annots }]);
    setFuture([]);
    setDirty(true);
  }, [bytes, annots]);

  const undo = () => {
    const last = history[history.length - 1];
    if (!last) return;
    setFuture((f) => [{ bytes, annots }, ...f]);
    setHistory((h) => h.slice(0, -1));
    if (last.bytes !== bytes) setBytes(last.bytes);
    setAnnots(last.annots);
  };
  const redo = () => {
    const next = future[0];
    if (!next) return;
    setHistory((h) => [...h, { bytes, annots }]);
    setFuture((f) => f.slice(1));
    if (next.bytes !== bytes) setBytes(next.bytes);
    setAnnots(next.annots);
  };

  const run = async (label: string, fn: () => Promise<void>) => {
    setBusy(label);
    setError(null);
    try {
      await fn();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(null);
    }
  };

  /** Thao tác trang: làm phẳng chú thích hiện có trước rồi mới biến đổi trang. */
  const pageOp = (label: string, op: (b: Uint8Array) => Promise<Uint8Array>) =>
    run(label, async () => {
      const flat = annots.length ? await applyAnnotations(bytes, annots) : bytes;
      const out = await op(flat);
      snapshot();
      setAnnots([]);
      setSelected(null);
      setBytes(out);
    });

  const save = (saveAs = false) =>
    run('Đang lưu…', async () => {
      const out = await applyAnnotations(bytes, annots);
      const p = await saveFile({ name: file.name, kind: 'pdf', path }, out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength) as ArrayBuffer, saveAs);
      if (p) setPath(p);
      setDirty(false);
    });

  const updateAnnot = (id: string, patch: Partial<Annot>, record = false) => {
    if (record) snapshot();
    setAnnots((list) => list.map((a) => (a.id === id ? ({ ...a, ...patch } as Annot) : a)));
  };
  const addAnnot = (a: Annot) => {
    snapshot();
    setAnnots((list) => [...list, a]);
  };
  const removeSelected = () => {
    if (!selected) return;
    snapshot();
    setAnnots((list) => list.filter((a) => a.id !== selected));
    setSelected(null);
  };

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const typing = (e.target as HTMLElement).closest('textarea,input');
      const mod = e.ctrlKey || e.metaKey;
      if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
      else if (typing) return;
      else if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
      else if (mod && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); }
      else if ((e.key === 'Delete' || e.key === 'Backspace') && selected) { e.preventDefault(); removeSelected(); }
      else if (e.key === 'Escape') { setSelected(null); setPending(null); setTool('select'); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  });

  const scrollToPage = (i: number) => {
    scrollRef.current?.querySelector(`[data-page="${i}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  const search = async () => {
    if (!doc || !query.trim()) return setHits(null);
    const q = query.trim().toLowerCase();
    const found: number[] = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const tc = await (await doc.getPage(i)).getTextContent();
      const text = tc.items.map((it) => ('str' in it ? it.str : '')).join(' ').toLowerCase();
      if (text.includes(q)) found.push(i - 1);
    }
    setHits(found);
    setHitIdx(0);
    if (found.length) scrollToPage(found[0]);
  };

  const pickImage = () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/png,image/jpeg';
    input.onchange = () => {
      const f = input.files?.[0];
      if (!f) return;
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;
        const img = new Image();
        img.onload = () => {
          const w = Math.min(220, img.width);
          setPending({ dataUrl, w, h: (img.height / img.width) * w });
          setTool('place');
        };
        img.src = dataUrl;
      };
      reader.readAsDataURL(f);
    };
    input.click();
  };

  const selAnnot = annots.find((a) => a.id === selected);
  const hint = tool === 'place' ? 'Bấm vào trang để đặt ảnh/chữ ký' : TOOLS.find((t) => t.id === tool)?.hint;

  return (
    <EditorShell
      kind="pdf"
      title={file.name}
      dirty={dirty}
      onClose={onClose}
      onSave={() => save()}
      onSaveAs={() => save(true)}
      busy={busy}
      error={error}
      onDismissError={() => setError(null)}
      toolbar={
        <>
          <ToolButton icon="↶" label="Hoàn tác" onClick={undo} disabled={!history.length} />
          <ToolButton icon="↷" label="Làm lại" onClick={redo} disabled={!future.length} />
          <Sep />
          {TOOLS.map((t) => (
            <ToolButton key={t.id} icon={t.icon} label={t.label} active={tool === t.id} onClick={() => { setTool(t.id); setPending(null); }} />
          ))}
          <ToolButton icon="✍" label="Chữ ký" onClick={() => setShowSign(true)} active={tool === 'place' && !!pending} />
          <ToolButton icon="🖼" label="Ảnh" onClick={pickImage} />
          <Sep />
          <label className="tb-field" title="Màu">
            <input type="color" value={color} onChange={(e) => {
              setColor(e.target.value);
              if (selAnnot && 'color' in selAnnot) updateAnnot(selAnnot.id, { color: e.target.value }, true);
            }} />
          </label>
          <label className="tb-field" title="Cỡ chữ">
            <select value={fontSize} onChange={(e) => {
              const v = Number(e.target.value);
              setFontSize(v);
              if (selAnnot?.type === 'text') updateAnnot(selAnnot.id, { size: v }, true);
            }}>
              {[8, 10, 11, 12, 14, 16, 18, 24, 32, 48].map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </label>
          {selected && <ToolButton icon="🗑" label="Xoá" onClick={removeSelected} />}
          <Sep />
          <ToolButton icon="−" label="Thu nhỏ" onClick={() => zoomTo(zoom / 1.25)} />
          <span className="tb-zoom">{Math.round(zoom * 100)}%</span>
          <ToolButton icon="+" label="Phóng to" onClick={() => zoomTo(zoom * 1.25)} />
          <ToolButton icon="⇔" label="Vừa khung" onClick={() => zoomTo(fitZoom())} />
          <Sep />
          <ToolButton icon="▤" label="Trang" active={showThumbs} onClick={() => setShowThumbs((v) => !v)} />
          <ToolButton icon="⊕" label="Gộp PDF" onClick={() => run('Đang gộp…', async () => {
            const other = await pickFile();
            if (!other) return;
            if (other.kind !== 'pdf') throw new Error('Hãy chọn một tệp PDF để gộp.');
            await pageOp('Đang gộp…', (b) => mergePdf(b, new Uint8Array(other.data)));
          })} />
          <form className="tb-search" onSubmit={(e) => { e.preventDefault(); search(); }}>
            <input placeholder="Tìm trong PDF…" value={query} onChange={(e) => setQuery(e.target.value)} />
            {hits && (
              <button type="button" className="tb-hits" onClick={() => {
                if (!hits.length) return;
                const n = (hitIdx + 1) % hits.length;
                setHitIdx(n);
                scrollToPage(hits[n]);
              }}>
                {hits.length ? `Trang ${hits[hitIdx] + 1} (${hitIdx + 1}/${hits.length}) ›` : 'Không thấy'}
              </button>
            )}
          </form>
        </>
      }
      status={`${sizes.length} trang · ${hint ?? ''}`}
    >
      <div className="pdf-body">
        {showThumbs && doc && (
          <aside className="pdf-thumbs">
            {sizes.map((s, i) => (
              <div key={i} className="thumb">
                <button className="thumb-img" onClick={() => scrollToPage(i)} title={`Trang ${i + 1}`}>
                  <PdfCanvas doc={doc} index={i} zoom={110 / s.w} size={s} />
                </button>
                <div className="thumb-actions">
                  <span>{i + 1}</span>
                  <button title="Xoay trái" onClick={() => pageOp('Đang xoay…', (b) => rotatePage(b, i, -90))}>⟲</button>
                  <button title="Xoay phải" onClick={() => pageOp('Đang xoay…', (b) => rotatePage(b, i, 90))}>⟳</button>
                  <button title="Lên" disabled={i === 0} onClick={() => pageOp('Đang di chuyển…', (b) => movePage(b, i, i - 1))}>↑</button>
                  <button title="Xuống" disabled={i === sizes.length - 1} onClick={() => pageOp('Đang di chuyển…', (b) => movePage(b, i, i + 1))}>↓</button>
                  <button title="Chèn trang trắng phía sau" onClick={() => pageOp('Đang chèn…', (b) => insertBlankPage(b, i + 1))}>＋</button>
                  <button title="Xoá trang" onClick={() => {
                    if (confirm(`Xoá trang ${i + 1}?`)) pageOp('Đang xoá…', (b) => deletePage(b, i));
                  }}>✕</button>
                </div>
              </div>
            ))}
          </aside>
        )}
        <div className="pdf-scroll" ref={scrollRef} onPointerDown={(e) => { if (e.target === e.currentTarget || e.target === pagesRef.current) setSelected(null); }}>
          {!doc && !error && <div className="loading">Đang mở tài liệu…</div>}
          <div className="pdf-pages" ref={pagesRef}>
          {doc && sizes.map((s, i) => (
            <div key={i} className="pdf-page" data-page={i} style={{ width: s.w * zoom, height: s.h * zoom }}>
              <PdfCanvas doc={doc} index={i} zoom={zoom} size={s} />
              <AnnotLayer
                page={i}
                size={s}
                zoom={zoom}
                tool={tool}
                color={color}
                fontSize={fontSize}
                annots={annots}
                selected={selected}
                pending={pending}
                onSelect={setSelected}
                onAdd={(a) => { addAnnot(a); setSelected(a.id); if (a.type === 'image') { setPending(null); setTool('select'); } if (a.type === 'text') setTool('select'); }}
                onUpdate={updateAnnot}
                onBeginChange={snapshot}
              />
            </div>
          ))}
          </div>
        </div>
        {doc && (
          <div className="zoom-float" role="group" aria-label="Thu phóng">
            <button onClick={() => zoomTo(zoom / 1.25)} title="Thu nhỏ" aria-label="Thu nhỏ">−</button>
            <button className="zoom-pct" onClick={() => zoomTo(fitZoom())} title="Vừa khung">{Math.round(zoom * 100)}%</button>
            <button onClick={() => zoomTo(zoom * 1.25)} title="Phóng to" aria-label="Phóng to">+</button>
          </div>
        )}
      </div>
      {showSign && (
        <SignaturePad
          onCancel={() => setShowSign(false)}
          onDone={(dataUrl, w, h) => {
            setShowSign(false);
            setPending({ dataUrl, w, h });
            setTool('place');
          }}
        />
      )}
    </EditorShell>
  );
}

/** Số điểm ảnh tối đa cho canvas cả trang. iOS giới hạn mỗi canvas ~16,7 triệu điểm ảnh và tổng bộ nhớ canvas,
 * vượt quá thì canvas trắng trơn. Khi phóng to lớn, trang được vẽ mờ ở mức này rồi phủ thêm một canvas nét
 * chỉ cho phần đang nhìn thấy. */
const MAX_CANVAS_PIXELS = 4_000_000;

type Region = { x: number; y: number; w: number; h: number };

/** Vẽ một trang PDF lên canvas khi trang xuất hiện trong khung nhìn. */
function PdfCanvas({ doc, index, zoom, size }: { doc: PDFDocumentProxy; index: number; zoom: number; size: Size }) {
  const ref = useRef<HTMLCanvasElement>(null);
  const detailRef = useRef<HTMLCanvasElement>(null);
  const [visible, setVisible] = useState(false);
  const [region, setRegion] = useState<Region | null>(null);
  const [shown, setShown] = useState<(Region & { zoom: number }) | null>(null);

  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const want = zoom * dpr;
  const baseScale = Math.min(want, Math.sqrt(MAX_CANVAS_PIXELS / (size.w * size.h)));
  const needDetail = baseScale < want * 0.98;

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const io = new IntersectionObserver(([e]) => setVisible(e.isIntersecting), { rootMargin: '600px' });
    io.observe(el);
    return () => {
      io.disconnect();
      // Giải phóng bộ nhớ canvas ngay (Safari giữ lại đến khi GC).
      el.width = el.height = 0;
      if (detailRef.current) detailRef.current.width = detailRef.current.height = 0;
    };
  }, []);

  // Canvas cả trang, độ phân giải bị giới hạn.
  useEffect(() => {
    if (!visible) return;
    let task: RenderTask | null = null;
    let cancelled = false;
    doc.getPage(index + 1).then((page) => {
      if (cancelled || !ref.current) return;
      const vp = page.getViewport({ scale: baseScale });
      const canvas = ref.current;
      canvas.width = Math.floor(vp.width);
      canvas.height = Math.floor(vp.height);
      task = page.render({ canvasContext: canvas.getContext('2d')!, viewport: vp });
      task.promise.catch(() => {});
    });
    return () => {
      cancelled = true;
      task?.cancel();
    };
  }, [doc, index, baseScale, visible]);

  // Theo dõi phần trang đang nằm trong khung cuộn để vẽ nét phần đó.
  useEffect(() => {
    if (!visible || !needDetail) {
      setRegion(null);
      return;
    }
    const canvas = ref.current!;
    const scroller = canvas.closest('.pdf-scroll') as HTMLElement | null;
    if (!scroller) return;
    let timer = 0;
    const measure = () => {
      const pr = canvas.getBoundingClientRect();
      const sr = scroller.getBoundingClientRect();
      // Chừa thêm 1/4 khung nhìn mỗi phía để cuộn nhẹ không lộ phần mờ.
      const mx = sr.width / 4, my = sr.height / 4;
      const x0 = Math.max(0, sr.left - mx - pr.left), y0 = Math.max(0, sr.top - my - pr.top);
      const x1 = Math.min(pr.width, sr.right + mx - pr.left), y1 = Math.min(pr.height, sr.bottom + my - pr.top);
      if (x1 <= x0 || y1 <= y0) return setRegion(null);
      const r = { x: Math.floor(x0), y: Math.floor(y0), w: Math.ceil(x1 - Math.floor(x0)), h: Math.ceil(y1 - Math.floor(y0)) };
      setRegion((old) => (old && old.x === r.x && old.y === r.y && old.w === r.w && old.h === r.h ? old : r));
    };
    const later = () => {
      clearTimeout(timer);
      timer = window.setTimeout(measure, 120);
    };
    measure();
    scroller.addEventListener('scroll', later, { passive: true });
    window.addEventListener('resize', later);
    return () => {
      clearTimeout(timer);
      scroller.removeEventListener('scroll', later);
      window.removeEventListener('resize', later);
    };
  }, [visible, needDetail, zoom]);

  // Canvas nét cho phần đang nhìn thấy.
  useEffect(() => {
    if (!region || !needDetail) {
      setShown(null);
      return;
    }
    let task: RenderTask | null = null;
    let cancelled = false;
    doc.getPage(index + 1).then((page) => {
      const canvas = detailRef.current;
      if (cancelled || !canvas) return;
      const vp = page.getViewport({ scale: want, offsetX: -region.x * dpr, offsetY: -region.y * dpr });
      // Vẽ vào canvas tạm rồi mới chép sang, để không bị nháy trắng khi đang vẽ.
      const tmp = document.createElement('canvas');
      tmp.width = Math.ceil(region.w * dpr);
      tmp.height = Math.ceil(region.h * dpr);
      task = page.render({ canvasContext: tmp.getContext('2d')!, viewport: vp });
      task.promise
        .then(() => {
          if (cancelled || !detailRef.current) return;
          canvas.width = tmp.width;
          canvas.height = tmp.height;
          canvas.getContext('2d')!.drawImage(tmp, 0, 0);
          setShown({ ...region, zoom });
        })
        .catch(() => {})
        .finally(() => {
          tmp.width = tmp.height = 0;
        });
    });
    return () => {
      cancelled = true;
      task?.cancel();
    };
  }, [doc, index, region, want, needDetail]);

  const detailOk = needDetail && shown && shown.zoom === zoom;
  return (
    <>
      <canvas ref={ref} className="pdf-canvas" style={{ width: size.w * zoom, height: size.h * zoom }} />
      <canvas
        ref={detailRef}
        className="pdf-canvas pdf-detail"
        style={detailOk ? { left: shown.x, top: shown.y, width: shown.w, height: shown.h } : { display: 'none' }}
      />
    </>
  );
}

interface LayerProps {
  page: number; size: Size; zoom: number; tool: Tool; color: string; fontSize: number;
  annots: Annot[]; selected: string | null; pending: { dataUrl: string; w: number; h: number } | null;
  onSelect(id: string | null): void; onAdd(a: Annot): void;
  onUpdate(id: string, patch: Partial<Annot>, record?: boolean): void; onBeginChange(): void;
}

function AnnotLayer(p: LayerProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [draft, setDraft] = useState<Annot | null>(null);
  const drag = useRef<{ id: string; sx: number; sy: number; orig: Annot; mode: 'move' | 'resize'; moved: boolean } | null>(null);
  const mine = useMemo(() => p.annots.filter((a) => a.page === p.page), [p.annots, p.page]);

  const pt = (e: React.PointerEvent): [number, number] => {
    const r = ref.current!.getBoundingClientRect();
    return [(e.clientX - r.left) / p.zoom, (e.clientY - r.top) / p.zoom];
  };

  const onDown = (e: React.PointerEvent) => {
    if (e.button !== 0) return;
    const [x, y] = pt(e);
    const target = (e.target as HTMLElement).closest('[data-annot]') as HTMLElement | null;
    if (p.tool === 'select') {
      if (!target) return p.onSelect(null);
      const id = target.dataset.annot!;
      const orig = p.annots.find((a) => a.id === id)!;
      p.onSelect(id);
      if ((e.target as HTMLElement).tagName === 'TEXTAREA') return;
      drag.current = { id, sx: x, sy: y, orig, mode: (e.target as HTMLElement).dataset.handle ? 'resize' : 'move', moved: false };
      ref.current!.setPointerCapture(e.pointerId);
      return;
    }
    e.preventDefault();
    if (p.tool === 'text') {
      p.onAdd({ id: uid(), page: p.page, type: 'text', x, y: y - p.fontSize * 0.6, text: '', size: p.fontSize, color: p.color });
      return;
    }
    if (p.tool === 'place' && p.pending) {
      const { dataUrl, w, h } = p.pending;
      p.onAdd({ id: uid(), page: p.page, type: 'image', x: x - w / 2, y: y - h / 2, w, h, dataUrl });
      return;
    }
    ref.current!.setPointerCapture(e.pointerId);
    if (p.tool === 'ink') setDraft({ id: uid(), page: p.page, type: 'ink', points: [[x, y]], color: p.color, width: 2 });
    else if (p.tool === 'highlight' || p.tool === 'rect' || p.tool === 'whiteout')
      setDraft({ id: uid(), page: p.page, type: p.tool, x, y, w: 0, h: 0, color: p.tool === 'highlight' ? '#ffd400' : p.color, ...{ ox: x, oy: y } } as Annot);
  };

  const onMove = (e: React.PointerEvent) => {
    const [x, y] = pt(e);
    if (drag.current) {
      const d = drag.current;
      const dx = x - d.sx, dy = y - d.sy;
      if (!d.moved) {
        if (Math.abs(dx) + Math.abs(dy) < 2) return;
        d.moved = true;
        p.onBeginChange();
      }
      const o = d.orig;
      if (d.mode === 'resize' && 'w' in o) {
        p.onUpdate(o.id, { w: Math.max(8, o.w + dx), h: o.type === 'image' ? Math.max(8, o.h * ((o.w + dx) / o.w)) : Math.max(4, o.h + dy) });
      } else if (o.type === 'ink') {
        p.onUpdate(o.id, { points: o.points.map(([px, py]) => [px + dx, py + dy] as [number, number]) });
      } else {
        p.onUpdate(o.id, { x: (o as { x: number }).x + dx, y: (o as { y: number }).y + dy });
      }
      return;
    }
    if (!draft) return;
    if (draft.type === 'ink') setDraft({ ...draft, points: [...draft.points, [x, y]] });
    else if ('w' in draft) {
      const { ox, oy } = draft as unknown as { ox: number; oy: number };
      setDraft({ ...draft, x: Math.min(ox, x), y: Math.min(oy, y), w: Math.abs(x - ox), h: Math.abs(y - oy) });
    }
  };

  const onUp = () => {
    drag.current = null;
    if (!draft) return;
    const ok = draft.type === 'ink' ? draft.points.length > 1 : 'w' in draft && draft.w > 3 && draft.h > 3;
    if (ok) {
      const { ox, oy, ...clean } = draft as unknown as Annot & { ox?: number; oy?: number };
      void ox; void oy;
      p.onAdd(clean as Annot);
    }
    setDraft(null);
  };

  const z = p.zoom;
  const all = draft ? [...mine, draft] : mine;
  return (
    <div
      ref={ref}
      className={`annot-layer tool-${p.tool}`}
      onPointerDown={onDown}
      onPointerMove={onMove}
      onPointerUp={onUp}
      onPointerCancel={onUp}
    >
      <svg width={p.size.w * z} height={p.size.h * z} viewBox={`0 0 ${p.size.w} ${p.size.h}`} className="annot-svg">
        {all.map((a) => {
          if (a.type === 'ink') {
            const d = a.points.map(([x, y], i) => `${i ? 'L' : 'M'}${x.toFixed(1)} ${y.toFixed(1)}`).join(' ');
            return (
              <g key={a.id} data-annot={a.id} className={a.id === p.selected ? 'sel' : ''}>
                <path d={d} stroke="transparent" strokeWidth={12} fill="none" />
                <path d={d} stroke={a.color} strokeWidth={a.width} fill="none" strokeLinecap="round" strokeLinejoin="round" />
              </g>
            );
          }
          if (a.type === 'highlight' || a.type === 'rect' || a.type === 'whiteout') {
            return (
              <rect
                key={a.id}
                data-annot={a.id}
                className={a.id === p.selected ? 'sel' : ''}
                x={a.x} y={a.y} width={a.w} height={a.h}
                fill={a.type === 'rect' ? 'transparent' : a.type === 'whiteout' ? '#fff' : a.color}
                fillOpacity={a.type === 'highlight' ? 0.35 : 1}
                stroke={a.type === 'rect' ? a.color : a.type === 'whiteout' ? '#ddd' : 'none'}
                strokeWidth={a.type === 'rect' ? 1.5 : 0.5}
              />
            );
          }
          return null;
        })}
      </svg>
      {mine.map((a) => {
        if (a.type === 'text') {
          return (
            <div
              key={a.id}
              data-annot={a.id}
              className={`annot-text ${a.id === p.selected ? 'sel' : ''}`}
              style={{ left: a.x * z, top: a.y * z, fontSize: a.size * z, color: a.color }}
            >
              <textarea
                autoFocus={!a.text}
                value={a.text}
                rows={Math.max(1, a.text.split('\n').length)}
                placeholder="Nhập chữ…"
                style={{ width: `${Math.max(4, ...a.text.split('\n').map((l) => l.length)) * 0.62 + 1.2}em` }}
                onFocus={() => p.onSelect(a.id)}
                onChange={(e) => p.onUpdate(a.id, { text: e.target.value })}
              />
            </div>
          );
        }
        if (a.type === 'image') {
          return (
            <div
              key={a.id}
              data-annot={a.id}
              className={`annot-img ${a.id === p.selected ? 'sel' : ''}`}
              style={{ left: a.x * z, top: a.y * z, width: a.w * z, height: a.h * z }}
            >
              <img src={a.dataUrl} alt="" draggable={false} />
              {a.id === p.selected && <span className="handle" data-handle="1" />}
            </div>
          );
        }
        if ((a.type === 'rect' || a.type === 'highlight' || a.type === 'whiteout') && a.id === p.selected) {
          return (
            <div key={a.id + 'h'} data-annot={a.id} className="annot-box-handle"
              style={{ left: (a.x + a.w) * z - 6, top: (a.y + a.h) * z - 6 }}>
              <span className="handle" data-handle="1" />
            </div>
          );
        }
        return null;
      })}
    </div>
  );
}
