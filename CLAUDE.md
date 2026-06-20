# Hướng dẫn cho Claude khi làm việc trong repo này

## Quy tắc làm việc (theo yêu cầu của chủ dự án)
- Với mỗi yêu cầu **phức tạp hoặc chưa rõ ràng**, PHẢI **xác nhận lại với người dùng
  trước khi làm** (ưu tiên hỏi dạng **trắc nghiệm**). Chỉ bắt tay code sau khi đã hiểu rõ.
- Các yêu cầu đơn giản, rõ ràng thì làm luôn, không cần hỏi.
- Trả lời và giao tiếp bằng **tiếng Việt**.

## Tổng quan dự án
P-FMEA Builder — công cụ web một trang (offline) giúp thiết lập P-FMEA tự động từ
Control Plan (CP). Bản phát hành là **một file HTML độc lập**: `P-FMEA-Builder.html`
(được build gộp từ `index.html`, `styles.css`, `js/*.js`, `data/*.js`, `vendor/*`).

### Cấu trúc
- `index.html` — bố cục giao diện.
- `styles.css` — giao diện.
- `js/app.js` — state, dựng bảng P-FMEA, lưu/mở dự án (localStorage), chấm điểm S.
- `js/parser.js` — đọc Control Plan (.xlsx) và rút hạng mục chất lượng.
- `js/export-template.js` — xuất Excel: đổ dữ liệu vào template gốc, giữ định dạng, in A4.
- `data/template.js` — template .xlsx (base64), `data/severity.js`, `data/material.js`.
- `vendor/` — `xlsx.full.min.js`, `fflate.min.js`.

### Build & kiểm thử
- Build file độc lập: `node /tmp/build_standalone.js` (nếu có) → `P-FMEA-Builder.html`.
- Kiểm thử logic bằng Node/JSDOM (đã dùng trong phiên trước).
- Khi xuất Excel: luôn kiểm tra XML hợp lệ (tránh Excel báo "We found a problem").
