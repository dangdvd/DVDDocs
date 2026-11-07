# DVDDocs

DVDDocs: Đọc & Sửa PDF, Office. Một ứng dụng duy nhất để **xem và chỉnh sửa PDF, Word (.docx) và Excel (.xlsx)**, chạy trên:

| Nền tảng | Công nghệ | Cách chạy |
|---|---|---|
| Web (mọi trình duyệt) | React + Vite | `npm run dev` |
| Windows / macOS / Linux | Electron | `npm run desktop`, đóng gói: `npm run dist:desktop` |
| Android | Capacitor | `npm run mobile:android` (cần Android Studio) |
| iPhone / iPad | Capacitor | `npm run mobile:ios` (cần macOS + Xcode) |

Mọi tệp được xử lý ngay trên thiết bị, không tải lên máy chủ nào.

**Tải bản cài đặt dùng thử:** xem mục [Releases](https://github.com/dangdvd/DVDDocs/releases) (Windows, macOS, Linux, Android APK, iPhone IPA chưa ký, Web).
Để phát hành bản mới: tăng `version` trong `package.json` và merge vào `main`; quy trình `.github/workflows/release.yml` tự build và tạo Release `vX.Y.Z`.

## Tính năng (bản 0.2)

**PDF**
- Xem nhiều trang, phóng to/thu nhỏ (chụm hai ngón), vừa khung, đi tới trang; tìm chữ và tô sáng từng kết quả, chuyển qua lại giữa các kết quả.
- **Điền biểu mẫu PDF (AcroForm)**: ô nhập chữ tiếng Việt, ô đánh dấu, nút chọn một, danh sách thả xuống; giá trị được ghi vào tệp khi lưu.
- Thêm chữ (hỗ trợ đầy đủ tiếng Việt, phông được nhúng vào tệp), tô sáng, vẽ khung, che trắng nội dung cũ, bút vẽ tự do.
- Ký tên bằng tay (chuột/ngón tay/bút) và chèn ảnh; kéo để di chuyển, kéo góc để đổi cỡ.
- Quản lý trang: xoay, sắp xếp lại, chèn trang trắng, xoá trang, gộp thêm PDF khác, **trích các trang chọn ra tệp mới** (ví dụ `1-3, 5`).
- Hoàn tác / làm lại; khi lưu, chú thích được ghi thẳng vào nội dung PDF.

**Word (.docx)**
- Soạn thảo: tiêu đề, **phông chữ, cỡ chữ**, đậm/nghiêng/gạch chân/gạch ngang, **chỉ số trên/dưới**, màu chữ, tô nền, căn lề, danh sách nhiều cấp, trích dẫn, bảng (thêm/xoá hàng cột, gộp ô), ảnh, đường kẻ, **liên kết**, **ngắt trang** (Ctrl+Enter).
- **Tìm & thay thế** (Ctrl+F, Ctrl+H), không phân biệt hoa thường.
- Thiết lập trang khi lưu: **A4 dọc/ngang**, **đánh số trang** ở chân trang ("Trang x / y").
- Chế độ **Bố cục gốc** (chỉ đọc) hiển thị trung thực tệp gốc với header/footer, phông và khổ giấy.
- Lưu ra .docx chuẩn mở được bằng mọi phần mềm văn phòng; in / xuất PDF qua hộp thoại in.

**Excel (.xlsx)**
- Lưới ảo hoá (mượt với bảng lớn), nhiều trang tính (thêm, đổi tên, xoá).
- Hàng trăm hàm Excel (SUM, IF, VLOOKUP, AVERAGE, …), tham chiếu liên trang, phát hiện tham chiếu vòng.
- Định dạng ô: đậm/nghiêng/gạch chân, màu chữ, màu nền, căn lề, định dạng số (tiền VNĐ/USD, %, ngày).
- **Gộp ô** (gộp và căn giữa), **kẻ khung**, **xuống dòng trong ô**, **cố định hàng/cột** khi cuộn, kéo đổi độ rộng cột và **chiều cao hàng**.
- **Tìm & thay thế** (Ctrl+F, Ctrl+H).
- Chèn/xoá hàng cột (tự dịch công thức và vùng gộp), sắp xếp, tự động tính tổng, Ctrl+D / Ctrl+R điền nhanh, sao chép/dán với Excel và Google Sheets.
- Khi lưu, giữ nguyên các phần của workbook gốc mà ứng dụng không chỉnh sửa.

**Chung:** mở tệp gần đây (lưu cục bộ), tạo tài liệu trống, kéo thả tệp, giao diện tối/sáng, phím tắt Ctrl/Cmd+S, Z, Y.

## Kiến trúc

```
src/
  App.tsx                 Màn hình chính: mở tệp, tạo mới, tệp gần đây
  lib/fileio.ts           Mở/lưu tệp theo nền tảng (Electron, Capacitor, trình duyệt)
  lib/recent.ts           Danh sách tệp gần đây trong IndexedDB
  components/EditorShell  Khung chung: thanh tiêu đề, thanh công cụ, thanh trạng thái
  editors/pdf/            pdf.js (hiển thị) + pdf-lib (chỉnh sửa, ghi tệp)
  editors/docx/           mammoth (đọc) + TipTap (soạn thảo) + docx (ghi) + docx-preview (bố cục gốc)
  editors/xlsx/           ExcelJS (đọc/ghi) + fast-formula-parser (công thức) + lưới tự viết
electron/                 Vỏ ứng dụng máy tính, hộp thoại mở/lưu của hệ điều hành
tests/                    Kiểm thử tự động (Vitest)
```

Một mã nguồn web duy nhất được đóng gói cho cả ba nhóm nền tảng, nên mọi tính năng mới có mặt ở mọi nơi cùng lúc. Trên điện thoại, nút **Lưu** ghi tệp vào thư mục Tài liệu của ứng dụng rồi mở bảng Chia sẻ để gửi hoặc lưu sang Files/Drive.

## Phát triển

```bash
npm install
npm run dev          # mở http://localhost:5173
npm test             # kiểm thử tự động
npm run build        # bản build web trong dist/
npm run desktop      # chạy bản máy tính
```

Ứng dụng di động: lần đầu chạy `npm run mobile:android` hoặc `npm run mobile:ios` để tạo dự án gốc; các lần sau chỉ cần `npm run mobile:sync` rồi build trong Android Studio / Xcode.

## Giới hạn hiện tại và lộ trình

- Word: chế độ soạn thảo chưa giữ hết định dạng phức tạp (header/footer có nội dung riêng, cột, ô có màu nền, màu chữ của tệp gốc, giãn dòng); các phần này vẫn xem được ở chế độ Bố cục gốc. Dự kiến: theo dõi thay đổi, chú thích, mục lục.
- Excel: chưa có biểu đồ, định dạng có điều kiện, lọc dữ liệu; viền được đơn giản hoá thành nét mảnh khi kẻ mới; ExcelJS có thể bỏ biểu đồ có sẵn khi lưu.
- PDF: chưa sửa trực tiếp chữ gốc trong PDF (dùng Che trắng + Chữ để thay thế), chưa có chữ ký số chứng thư; biểu mẫu XFA (một số mẫu của Adobe) chưa được hỗ trợ.
- Chưa mở tệp .doc/.xls/.ppt đời cũ và PowerPoint.

## Giấy phép bên thứ ba

Phông DejaVu Sans (public/fonts) theo giấy phép trong `public/fonts/DejaVu-LICENSE.txt`. Các thư viện npm theo giấy phép riêng của từng gói (MIT/Apache-2.0/BSD).
