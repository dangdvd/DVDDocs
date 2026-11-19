import { useEffect } from 'react';
import type { DocKind } from '../lib/fileio';

const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);
const MOD = isMac ? '⌘' : 'Ctrl';

const COMMON: [string, string][] = [
  [`${MOD}+S`, 'Lưu'],
  [`${MOD}+Z`, 'Hoàn tác'],
  [`${MOD}+Y / ${MOD}+Shift+Z`, 'Làm lại'],
  ['F1 / ' + `${MOD}+/`, 'Mở bảng phím tắt này'],
];

export const SHORTCUTS: Record<DocKind, [string, string][]> = {
  docx: [
    [`${MOD}+B / I / U`, 'Đậm / nghiêng / gạch chân'],
    [`${MOD}+F`, 'Tìm'],
    [`${MOD}+H`, 'Tìm & thay thế'],
    [`${MOD}+Enter`, 'Ngắt trang'],
    [`${MOD}+P`, 'In / xuất PDF'],
    [`${MOD}+Shift+7 / 8`, 'Danh sách đánh số / dấu đầu dòng'],
    ['Tab / Shift+Tab', 'Thụt vào / thụt ra trong danh sách'],
  ],
  xlsx: [
    ['Phím mũi tên', 'Di chuyển ô (giữ Shift để chọn vùng)'],
    ['Enter / Tab', 'Xác nhận và sang ô dưới / bên phải'],
    ['F2', 'Sửa ô đang chọn'],
    ['Delete', 'Xoá nội dung vùng chọn'],
    [`${MOD}+B / I / U`, 'Đậm / nghiêng / gạch chân'],
    [`${MOD}+C / X / V`, 'Sao chép / cắt / dán (cả với Excel, Google Sheets)'],
    [`${MOD}+D / ${MOD}+R`, 'Điền xuống / điền sang phải'],
    [`${MOD}+A`, 'Chọn cả trang tính'],
    [`${MOD}+F / ${MOD}+H`, 'Tìm / Tìm & thay thế'],
    [`${MOD}+P`, 'In / xuất PDF'],
    ['Alt+Enter', 'Xuống dòng trong ô'],
  ],
  pdf: [
    ['Delete', 'Xoá chú thích đang chọn'],
    ['Esc', 'Bỏ chọn, về công cụ Chọn'],
    [`${MOD}+P`, 'In / xuất PDF'],
    ['Chụm hai ngón / Ctrl+cuộn chuột', 'Phóng to, thu nhỏ'],
  ],
};

const TITLE: Record<DocKind, string> = { docx: 'Word', xlsx: 'Excel', pdf: 'PDF' };

export function ShortcutsDialog({ kind, onClose }: { kind: DocKind; onClose(): void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [onClose]);
  const rows = [...COMMON, ...SHORTCUTS[kind]];
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal shortcuts" role="dialog" aria-modal="true" aria-labelledby="sc-title" onClick={(e) => e.stopPropagation()}>
        <h3 id="sc-title">Phím tắt – {TITLE[kind]}</h3>
        <table>
          <tbody>
            {rows.map(([k, d]) => (
              <tr key={k}><td><kbd>{k}</kbd></td><td>{d}</td></tr>
            ))}
          </tbody>
        </table>
        <p className="muted small">Thay đổi được tự động lưu thành bản nháp trên thiết bị; nếu ứng dụng bị đóng đột ngột, mở lại để khôi phục ở trang chủ.</p>
        <div className="modal-actions" style={{ justifyContent: 'flex-end' }}>
          <button className="btn primary" autoFocus onClick={onClose}>Đóng</button>
        </div>
      </div>
    </div>
  );
}
