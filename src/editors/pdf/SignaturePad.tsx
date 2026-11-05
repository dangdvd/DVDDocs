import { useEffect, useRef, useState } from 'react';

/** Hộp thoại ký tên bằng ngón tay / chuột / bút. Trả về PNG nền trong suốt đã cắt sát nét. */
export function SignaturePad({ onDone, onCancel }: { onDone(dataUrl: string, w: number, h: number): void; onCancel(): void }) {
  const ref = useRef<HTMLCanvasElement>(null);
  const drawing = useRef(false);
  const [empty, setEmpty] = useState(true);
  const [color, setColor] = useState('#0b3d91');

  useEffect(() => {
    const c = ref.current!;
    const dpr = window.devicePixelRatio || 1;
    c.width = c.clientWidth * dpr;
    c.height = c.clientHeight * dpr;
    const ctx = c.getContext('2d')!;
    ctx.scale(dpr, dpr);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.lineWidth = 2.5;
  }, []);

  const pos = (e: React.PointerEvent) => {
    const r = ref.current!.getBoundingClientRect();
    return [e.clientX - r.left, e.clientY - r.top] as const;
  };

  const clear = () => {
    const c = ref.current!;
    c.getContext('2d')!.clearRect(0, 0, c.width, c.height);
    setEmpty(true);
  };

  const done = () => {
    const c = ref.current!;
    const ctx = c.getContext('2d')!;
    const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height);
    let minX = width, minY = height, maxX = 0, maxY = 0;
    for (let y = 0; y < height; y++)
      for (let x = 0; x < width; x++)
        if (data[(y * width + x) * 4 + 3] > 0) {
          if (x < minX) minX = x; if (x > maxX) maxX = x;
          if (y < minY) minY = y; if (y > maxY) maxY = y;
        }
    if (maxX <= minX) return;
    const pad = 6;
    const w = maxX - minX + pad * 2, h = maxY - minY + pad * 2;
    const out = document.createElement('canvas');
    out.width = w; out.height = h;
    out.getContext('2d')!.drawImage(c, minX - pad, minY - pad, w, h, 0, 0, w, h);
    const targetW = 160;
    onDone(out.toDataURL('image/png'), targetW, (h / w) * targetW);
  };

  return (
    <div className="modal-backdrop" onPointerDown={(e) => e.target === e.currentTarget && onCancel()}>
      <div className="modal">
        <h3>Ký tên</h3>
        <p className="muted">Ký vào khung bên dưới bằng chuột, ngón tay hoặc bút.</p>
        <canvas
          ref={ref}
          className="sign-canvas"
          onPointerDown={(e) => {
            drawing.current = true;
            ref.current!.setPointerCapture(e.pointerId);
            const ctx = ref.current!.getContext('2d')!;
            ctx.strokeStyle = color;
            ctx.beginPath();
            ctx.moveTo(...pos(e));
          }}
          onPointerMove={(e) => {
            if (!drawing.current) return;
            const ctx = ref.current!.getContext('2d')!;
            ctx.lineTo(...pos(e));
            ctx.stroke();
            setEmpty(false);
          }}
          onPointerUp={() => (drawing.current = false)}
        />
        <div className="modal-actions">
          <input type="color" value={color} onChange={(e) => setColor(e.target.value)} title="Màu mực" />
          <button className="btn" onClick={clear}>Xoá</button>
          <span style={{ flex: 1 }} />
          <button className="btn" onClick={onCancel}>Huỷ</button>
          <button className="btn primary" disabled={empty} onClick={done}>Dùng chữ ký</button>
        </div>
      </div>
    </div>
  );
}
