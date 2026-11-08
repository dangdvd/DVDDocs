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

## Tính năng (bản 0.4)

**PDF**
- Xem nhiều trang, phóng to/thu nhỏ (chụm hai ngón), vừa khung, đi tới trang; tìm chữ và tô sáng từng kết quả, chuyển qua lại giữa các kết quả.
- **Điền biểu mẫu PDF (AcroForm)**: ô nhập chữ tiếng Việt, ô đánh dấu, nút chọn một, danh sách thả xuống; giá trị được ghi vào tệp khi lưu.
- **Sửa chữ có sẵn trong PDF**: bấm vào một dòng chữ để sửa ngay tại chỗ; chữ cũ được xoá khỏi nội dung trang (không còn tìm/sao chép được), chữ mới giữ vị trí, cỡ và màu, **dùng phông gần giống phông gốc** (có chân / không chân / đơn cách, đậm, nghiêng; đổi được trên thanh công cụ).
- Thêm chữ (hỗ trợ đầy đủ tiếng Việt, phông được nhúng vào tệp), tô sáng, vẽ khung, che trắng nội dung cũ, bút vẽ tự do.
- Ký tên bằng tay (chuột/ngón tay/bút) và chèn ảnh; kéo để di chuyển, kéo góc để đổi cỡ.
- Quản lý trang: xoay, sắp xếp lại, chèn trang trắng, xoá trang, gộp thêm PDF khác, **trích các trang chọn ra tệp mới** (ví dụ `1-3, 5`).
- Hoàn tác / làm lại; khi lưu, chú thích được ghi thẳng vào nội dung PDF.

**Word (.docx)**
- Soạn thảo: tiêu đề, **phông chữ, cỡ chữ**, đậm/nghiêng/gạch chân/gạch ngang, **chỉ số trên/dưới**, màu chữ, tô nền, căn lề, danh sách nhiều cấp, trích dẫn, bảng (thêm/xoá hàng cột, gộp ô), ảnh, đường kẻ, **liên kết**, **ngắt trang** (Ctrl+Enter).
- **Tìm & thay thế** (Ctrl+F, Ctrl+H), không phân biệt hoa thường.
- **Theo dõi thay đổi**: chữ thêm hiện gạch chân xanh, chữ xoá hiện gạch ngang đỏ kèm tên người sửa; chuyển qua lại, chấp nhận/từ chối từng chỗ hoặc tất cả. Lưu thành thay đổi thật của Word (w:ins/w:del) và đọc lại được từ tệp do Word tạo.
- **Giữ định dạng của tệp gốc** khi mở để sửa: màu chữ, tô nền chữ, căn lề, giãn dòng, màu nền ô bảng.
- **Sửa đầu trang / chân trang** (bấm vào đầu/chân trang hoặc nút trên thanh công cụ), **chèn mục lục** tự cập nhật theo các tiêu đề, **giãn dòng** 1 – 3.
- Thiết lập trang khi lưu: **A4 dọc/ngang**, **đánh số trang** ở chân trang ("Trang x / y").
- Chế độ **Bố cục gốc** (chỉ đọc) hiển thị trung thực tệp gốc với header/footer, phông và khổ giấy.
- Lưu ra .docx chuẩn mở được bằng mọi phần mềm văn phòng; in / xuất PDF qua hộp thoại in.

**Excel (.xlsx)**
- Lưới ảo hoá (mượt với bảng lớn), nhiều trang tính (thêm, đổi tên, xoá).
- Hàng trăm hàm Excel (SUM, IF, VLOOKUP, AVERAGE, …), tham chiếu liên trang, phát hiện tham chiếu vòng.
- Định dạng ô: đậm/nghiêng/gạch chân, màu chữ, màu nền, căn lề, định dạng số (tiền VNĐ/USD, %, ngày).
- **Gộp ô** (gộp và căn giữa), **kẻ khung**, **xuống dòng trong ô**, **cố định hàng/cột** khi cuộn, kéo đổi độ rộng cột và **chiều cao hàng**.
- **Tìm & thay thế** (Ctrl+F, Ctrl+H).
- **Định dạng có điều kiện**: lớn hơn / nhỏ hơn / bằng / khác / trong khoảng / chứa chữ (tô nền và màu chữ), **thang màu** 2–3 màu, **thanh dữ liệu**; đọc và ghi đúng quy tắc của Excel.
- **Lọc dữ liệu (AutoFilter)**: nút lọc trên hàng tiêu đề, chọn giá trị cần hiện, tìm giá trị, sắp xếp A→Z / Z→A; hàng bị lọc được ẩn và tiêu chí lọc được lưu vào tệp như Excel.
- **Mở tệp CSV/TSV** (tự nhận dấu phân cách, bảng mã, số kiểu Việt Nam 1.234,5; giữ số 0 ở đầu mã/số điện thoại), lưu thành .xlsx.
- **Biểu đồ**: cột, thanh ngang, đường, vùng, tròn, tạo từ vùng chọn (tự nhận hàng tiêu đề và cột nhãn), kéo để di chuyển, kéo góc để đổi cỡ, đổi loại và tiêu đề; lưu thành biểu đồ Excel thật và đọc lại biểu đồ có sẵn trong tệp; biểu đồ có sẵn không bị sửa thì **giữ nguyên màu sắc và kiểu riêng** (chỉ cập nhật số liệu).
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

- Word: chế độ soạn thảo chưa giữ bố cục nhiều cột, hộp văn bản, đầu/chân trang khác nhau cho trang đầu hoặc trang chẵn/lẻ (chỉ sửa đầu/chân trang mặc định); các phần này vẫn xem được ở chế độ Bố cục gốc. Theo dõi thay đổi chưa ghi lại thay đổi định dạng và việc nối/tách đoạn. Chưa có chú thích (comment).
- Excel: định dạng có điều kiện dạng công thức, bộ biểu tượng, top 10 được giữ nguyên khi lưu nhưng chưa hiển thị; lọc theo điều kiện số/ngày của Excel được giữ hàng ẩn nhưng chưa sửa được. Viền được đơn giản hoá thành nét mảnh khi kẻ mới. Biểu đồ phân tán, radar được lưu thành biểu đồ đường nếu bị sửa.
- PDF: phông khi sửa chữ là họ DejaVu gần giống phông gốc, không phải chính phông gốc (phông nhúng trong PDF thường chỉ chứa những chữ đã dùng nên không thể gõ chữ mới bằng nó). Chữ nằm nghiêng hoặc nằm trong nhóm đồ hoạ (form XObject) chỉ được che, không xoá được khỏi tệp.
- **Chưa hỗ trợ và lý do:**
  - Tệp .doc/.xls/.ppt đời cũ (định dạng nhị phân trước 2007): thư viện đọc miễn phí còn lỗ hổng bảo mật chưa vá; hãy mở bằng Word/Excel hoặc LibreOffice rồi lưu thành .docx/.xlsx. Tệp CSV thì mở được trực tiếp.
  - PowerPoint: là một trình soạn thảo riêng, chưa nằm trong phạm vi ứng dụng.
  - Chữ ký số bằng chứng thư (USB token, chữ ký số doanh nghiệp): cần chứng thư số và trình điều khiển token của nhà cung cấp; ứng dụng chỉ có chữ ký tay.
  - Biểu mẫu XFA (biểu mẫu động của Adobe LiveCycle): định dạng đã bị Adobe ngừng phát triển, chỉ Adobe Reader mở được.

## Giấy phép bên thứ ba

Phông DejaVu Sans, Serif, Sans Mono (public/fonts) theo giấy phép trong `public/fonts/DejaVu-LICENSE.txt`. Các thư viện npm theo giấy phép riêng của từng gói (MIT/Apache-2.0/BSD).
