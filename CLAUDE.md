# Hướng dẫn cho Claude khi làm việc trong repo này

> **ĐỌC TRƯỚC:** Khi bắt đầu một phiên làm việc mới, hãy đọc cả file
> [`HANDOFF.md`](./HANDOFF.md) — tài liệu bàn giao đầy đủ (kiến trúc, cấu hình
> Supabase/AI, trạng thái hiện tại, và việc cần làm tiếp).

## Quy tắc làm việc (theo yêu cầu của chủ dự án)
- Với mỗi yêu cầu **phức tạp hoặc chưa rõ ràng**, PHẢI **xác nhận lại với người dùng
  trước khi làm** (ưu tiên hỏi dạng **trắc nghiệm**). Chỉ bắt tay code sau khi đã hiểu rõ.
- Các yêu cầu đơn giản, rõ ràng thì làm luôn, không cần hỏi.
- Trả lời và giao tiếp bằng **tiếng Việt**.
- **BẢO MẬT:** Tuyệt đối **không commit** khóa bí mật (service_role / secret key).
  Chỉ được nhúng **anon/publishable key** của Supabase (đây là khóa công khai).
  API key Gemini chỉ lưu ở trình duyệt người dùng, **không ghi vào code**.

## Tổng quan dự án
P-FMEA Builder — công cụ web một trang (offline) giúp thiết lập P-FMEA tự động từ
Control Plan (CP). Bản phát hành là **một file HTML độc lập**: `P-FMEA-Builder.html`
(được build gộp từ `index.html`, `styles.css`, `js/*.js`, `data/*.js`, `vendor/*`).

### Cấu trúc
- `index.html` — bố cục giao diện (2 tab: Thiết lập P-FMEA / Hướng dẫn).
- `styles.css` — giao diện.
- `js/app.js` — state, dựng bảng P-FMEA, lưu/mở dự án (localStorage + Supabase),
  chấm điểm S, gợi ý AI, dựng tab Hướng dẫn.
- `js/parser.js` — đọc Control Plan (.xlsx) và rút hạng mục chất lượng.
- `js/export-template.js` — xuất Excel: đổ dữ liệu vào template gốc, giữ định dạng, in A4.
- `data/template.js` — template .xlsx (base64); `data/severity.js` — bảng điểm S;
  `data/material.js` — cây Bộ phận→Sản phẩm→Dây chuyền.
- `vendor/` — `xlsx.full.min.js`, `fflate.min.js`.
- `tools/build.js` — script build file độc lập.

### Build & kiểm thử
- Build file độc lập: **`node tools/build.js`** → `P-FMEA-Builder.html` (~1.4 MB).
  **Luôn build lại sau khi sửa** `index.html` / `styles.css` / `js/*.js`.
- Kiểm thử logic bằng Node + JSDOM (xem ví dụ trong `HANDOFF.md`).
- Khi xuất Excel: luôn kiểm tra XML hợp lệ (tránh Excel báo "We found a problem").

### Triển khai (deploy)
- GitHub Pages chạy từ branch `claude/upbeat-heisenberg-xpkxr3`.
- Chỉ cần `git push` đúng branch → web tự cập nhật sau vài phút.
- URL: https://hailv0994.github.io/Thiet-lap-P-FMEA/ (phân biệt hoa/thường).
- File `.nojekyll` bắt buộc giữ lại (để Pages phục vụ thư mục `vendor/`, `data/`).
