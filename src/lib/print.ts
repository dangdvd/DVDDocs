// In tài liệu qua hộp thoại in của hệ thống (chọn "Lưu thành PDF" để xuất PDF).
// Nội dung cần in được dựng trong một vùng riêng #print-root, phần giao diện còn lại bị ẩn khi in.

/** Dựng nội dung in bằng `build`, mở hộp thoại in rồi dọn dẹp. */
export async function printContent(build: (root: HTMLElement) => Promise<void> | void, pageCss = '') {
  document.getElementById('print-root')?.remove();
  const root = document.createElement('div');
  root.id = 'print-root';
  const style = document.createElement('style');
  style.textContent = pageCss;
  root.appendChild(style);
  document.body.appendChild(root);
  try {
    await build(root);
    // Đợi ảnh tải xong để không in ra trang trắng.
    await Promise.all([...root.querySelectorAll('img')].map((img) => (img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; }))));
  } catch (e) {
    root.remove();
    throw e;
  }
  document.body.classList.add('printing');
  const cleanup = () => {
    document.body.classList.remove('printing');
    root.querySelectorAll('img').forEach((img) => { if (img.src.startsWith('blob:')) URL.revokeObjectURL(img.src); });
    root.remove();
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup);
  window.print();
  // Một số trình duyệt (Safari, WebView) không phát sự kiện afterprint: dọn dẹp sau một lúc.
  setTimeout(() => { if (document.body.contains(root)) cleanup(); }, 60_000);
}

export const escapeHtml = (s: string) => s.replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]!));
