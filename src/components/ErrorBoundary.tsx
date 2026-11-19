import { Component, type ErrorInfo, type ReactNode } from 'react';

/** Bắt lỗi giao diện bất ngờ: hiện màn hình thân thiện thay vì trang trắng, cho phép về trang chủ hoặc tải lại. */
export class ErrorBoundary extends Component<{ children: ReactNode; onHome(): void }, { error: Error | null }> {
  state = { error: null as Error | null };

  static getDerivedStateFromError(error: Error) {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error('[DVDDocs] Lỗi giao diện:', error, info.componentStack);
  }

  render() {
    const { error } = this.state;
    if (!error) return this.props.children;
    const details = `${error.name}: ${error.message}\n${error.stack ?? ''}`.slice(0, 4000);
    return (
      <div className="crash" role="alert">
        <div className="crash-card">
          <h1>Rất tiếc, đã xảy ra lỗi</h1>
          <p>Trình soạn thảo gặp lỗi ngoài ý muốn và phải dừng lại. Các thay đổi chưa lưu được giữ thành <b>bản nháp</b> (nếu kịp lưu); mở lại ở mục <b>Khôi phục tài liệu chưa lưu</b> trên trang chủ.</p>
          <div className="modal-actions">
            <button className="btn primary" onClick={() => { this.setState({ error: null }); this.props.onHome(); }}>Về trang chủ</button>
            <button className="btn" onClick={() => location.reload()}>Tải lại ứng dụng</button>
            <button className="btn" onClick={() => navigator.clipboard?.writeText(details).catch(() => {})}>Sao chép chi tiết lỗi</button>
          </div>
          <details>
            <summary>Chi tiết kỹ thuật</summary>
            <pre>{details}</pre>
          </details>
        </div>
      </div>
    );
  }
}
