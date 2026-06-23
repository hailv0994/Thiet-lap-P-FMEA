# 📋 Tài liệu bàn giao — P-FMEA Builder

> File này giúp một phiên Claude Code **mới** (ví dụ chạy trong Antigravity IDE,
> hoặc Claude Code web) nắm toàn bộ dự án và làm tiếp được ngay.
> Đọc cùng với [`CLAUDE.md`](./CLAUDE.md).

---

## 1. Dự án là gì

**P-FMEA Builder** — công cụ web một trang (chạy offline trong trình duyệt) giúp
kỹ sư thiết lập bảng **P-FMEA** (Process Failure Mode and Effects Analysis) tự
động từ file **Control Plan (CP)** Excel, rồi xuất ra file Excel đúng form chuẩn
**GL SQS0811** (in A4 ngang).

- Người dùng: kỹ sư chất lượng tại nhà máy sản xuất **giảm xóc xe máy**.
- Giao tiếp & nội dung: **tiếng Việt**.
- Bản phát hành: **một file HTML độc lập** `P-FMEA-Builder.html` (build gộp hết).
- Đang chạy online qua GitHub Pages cho cả nhóm dùng chung.

---

## 2. Kiến trúc & file

| File | Vai trò |
|------|---------|
| `index.html` | Bố cục. 2 tab: **Thiết lập P-FMEA** (`#tabBuilder`) và **Hướng dẫn** (`#tabGuide`). |
| `styles.css` | Toàn bộ giao diện. Biến màu ở `:root` (tông đen/trắng/xám tối giản). |
| `js/app.js` | Lõi: state, dựng bảng P-FMEA, chấm điểm S, lưu/mở dự án, đồng bộ Supabase, gợi ý AI Gemini, dựng tab Hướng dẫn (data-driven). |
| `js/parser.js` | Đọc CP `.xlsx`, rút tên công đoạn + hạng mục chất lượng. Hàm `vnText()` nối các dòng tiếng Việt liền nhau. |
| `js/export-template.js` | Xuất Excel: đổ dữ liệu vào template gốc, tự cân cột, in vừa khổ A4 ngang. |
| `data/template.js` | Template `.xlsx` gốc dạng base64. |
| `data/severity.js` | `window.SEVERITY_TABLE` — bảng tiêu chuẩn điểm **S** (scope `product`/`process`, rank 1–10). **Là dữ liệu thật GL SQS0811.** |
| `data/material.js` | Cây phân cấp **Bộ phận → Sản phẩm → Dây chuyền** cho dropdown. |
| `vendor/` | `xlsx.full.min.js` (SheetJS), `fflate.min.js`. |
| `tools/build.js` | Script build → `P-FMEA-Builder.html`. |
| `.nojekyll` | Bắt buộc giữ — để GitHub Pages phục vụ `vendor/`, `data/`. |

### Build
```bash
node tools/build.js      # gộp tất cả → P-FMEA-Builder.html (~1.4 MB)
```
Build script dùng đường dẫn động (`__dirname`), chạy được trên mọi máy. Nó giữ lại
các thẻ `<script src="https://...">` (CDN như Supabase), chỉ gỡ thẻ script nội bộ
và nhúng nội dung file vào (base64 → eval theo đúng thứ tự).

**Quy tắc vàng:** mỗi lần sửa `index.html`/`styles.css`/`js/*.js` → **chạy lại
`node tools/build.js`** rồi mới commit (nếu không, file web online không đổi).

### Kiểm thử bằng JSDOM (không cần trình duyệt)
```bash
cd /tmp && npm install jsdom --no-save   # nếu chưa có
# rồi load P-FMEA-Builder.html với runScripts:'dangerously',
# dispatch DOMContentLoaded, kiểm tra DOM. (Đã dùng nhiều lần, hoạt động tốt.)
```

---

## 3. Triển khai (deploy)

- **Branch phát triển & deploy:** `claude/upbeat-heisenberg-xpkxr3`.
- **Cập nhật web:** chỉ cần `git push` đúng branch → cả Netlify lẫn GitHub Pages tự
  build lại sau vài phút (không cần thao tác thêm).
- **URL chính (Netlify):** https://pfmea-tool.netlify.app — không có username trong URL.
- **URL dự phòng (GitHub Pages):** https://hailv0994.github.io/Thiet-lap-P-FMEA/ (phân biệt hoa/thường).
- Repo GitHub: `hailv0994/thiet-lap-p-fmea`.
- Netlify: site name `pfmea-tool`, deploy từ branch `claude/upbeat-heisenberg-xpkxr3`,
  publish directory `.`, không có build command.

---

## 4. Lưu trữ dữ liệu

### localStorage (cục bộ trong trình duyệt mỗi người)
| Key | Nội dung |
|-----|----------|
| `pfmea_autosave_v1` | Tự lưu phiên đang làm (meta + nội dung bảng). |
| `pfmea_projects_v1` | Các dự án/model base đã lưu. |
| `pfmea_gemini_key_v1` | API key Gemini của người dùng (KHÔNG đẩy lên cloud). |
| `pfmea_gemini_model_v1` | Model AI đã chọn. |
| `pfmea_context_v1` | Bối cảnh cho AI theo bộ phận. |
| `pfmea_phrases_v1` | Bộ nhớ mẫu câu theo cột + bộ phận. |

### Supabase (dữ liệu dùng chung cho cả nhóm)
- Project: `iccrgjtkaizosocrxsql`
- URL: `https://iccrgjtkaizosocrxsql.supabase.co`
- **Anon key** (công khai, an toàn để nhúng frontend) đã có sẵn trong `js/app.js`
  (hằng `SB_KEY`). **TUYỆT ĐỐI không thay bằng service_role key.**
- 3 bảng dùng chung: dự án (model base), bối cảnh AI, mẫu câu — RLS mở (không cần
  đăng nhập). Đồng bộ tự động khi Lưu/Mở; nếu offline thì giảm cấp êm (vẫn dùng
  localStorage). Chấm tròn trạng thái: `#cloudDot` xanh = đã kết nối.

### AI gợi ý (Google Gemini, free tier)
- Mỗi người tự dán API key (lấy ở `aistudio.google.com/apikey`), lưu ở trình duyệt.
- `callGemini()` thử lần lượt nhiều model: `gemini-2.5-flash` → `2.0-flash` →
  `2.0-flash-lite` → `1.5-flash`, bỏ qua model hết quota/không tồn tại.
- Nút AI mở popup: mẫu câu đã lưu (theo cột+bộ phận) + gợi ý AI + ô nhập tự do.

---

## 5. Tính năng đã có

- **Dropdown phân cấp** Bộ phận → Sản phẩm → Dây chuyền (chỉ hiện dữ liệu Material).
- **Tải Control Plan** → nạp **tất cả công đoạn** một lần; sắp xếp thứ tự công đoạn (▲▼).
- **Model base + merge:** chọn model base đã lưu → khi nạp CP mới sẽ giữ phần phân
  tích cũ, cập nhật phần thông số (spec) từ CP mới.
- **Chấm điểm S tự động** theo bảng tiêu chuẩn (chọn câu kết luận → ra điểm).
- **RPN tự tính** = S × O × D. Gộp ô (merge) khi 1 dạng hỏng có nhiều nguyên nhân.
- **Số thứ tự dạng hỏng hóc**: ô cột B tự đánh số khớp số yêu cầu ở cột A (vd yêu cầu
  "1. ..." → dạng hỏng hóc "1. ... không đạt"), hiển thị cả trên web lẫn file Excel.
- **Gộp dạng hỏng hóc** (nút 🔗 trên mỗi ô cột B): gộp nhiều dạng hỏng hóc trong cùng
  công đoạn vào 1 ô khi **chữ ký giống hệt** (mọi cột trừ yêu cầu & dạng hỏng hóc —
  tức Ảnh hưởng/Nguyên nhân/Dự phòng/Phát hiện ra, kể cả mục kiểm tra + tần suất ở ②).
  Dữ liệu chung được đồng bộ giữa các thành viên (`syncMergeGroup`), nút 🔓 để tách.
  Cài đặt: `req.mergeId`, hàm `reqSig()`/`reqGroups()` (có cả trong `export-template.js`).
- **Xuất Excel** đúng form gốc, in vừa khổ A4 ngang, tự đánh số trang, đã bỏ sheet "VÍ DỤ".
- **Sao lưu / Nạp** dự án ra file `.json`.
- **Đồng bộ đám mây** (Supabase) cho model base / bối cảnh / mẫu câu.
- **Gợi ý AI** (Gemini) cho từng cột.
- **Tab Hướng dẫn** (mới): trang đầu hiển thị format P-FMEA tổng quan; mỗi trang sau
  hướng dẫn 1 cột (trái = cột + ví dụ, phải = cách hiểu/cách làm), dựng động trong
  `js/app.js` (hàm `gPages()`, `setupGuide()`, `renderGuidePage()`).

---

## 6. ⭐ VIỆC CẦN LÀM TIẾP — TỐI ƯU GIAO DIỆN WEB

Chủ dự án thấy **giao diện hiện tại chưa đẹp**, muốn **làm lại cho hiện đại,
chuyên nghiệp, dễ nhìn hơn** mà **không làm hỏng chức năng đang chạy**.

### Nguyên tắc bắt buộc khi chỉnh giao diện
1. **KHÔNG đổi `id`/`name`** của các phần tử mà `js/app.js` đang truy cập
   (vd: `#mDept`, `#mProduct`, `#mLine`, `#mModel`, `#projSelect`, `#fileCP`,
   `#btnLoad`, `#btnSave`, `#btnExport`, `#fmea`, `#guideRoot`, `#guideJump`,
   `#guidePrev`, `#guideNext`, `.tab-btn[data-tab]`, các id trong `#aiModal`...).
   Đổi id sẽ làm gãy JS. Nếu cần đổi, phải sửa đồng bộ trong `app.js`.
2. **Giữ nguyên logic & cấu trúc bảng P-FMEA** (`buildHeader()`, các class cột
   `.c-a`…`.c-s`, cơ chế merge ô). Bảng phải vẫn hiển thị đủ tới cột cần thiết.
3. **Build lại** (`node tools/build.js`) và **test JSDOM** sau khi sửa.
4. Mọi thay đổi giao diện phức tạp → **hỏi xác nhận trắc nghiệm trước** (theo CLAUDE.md).

### Hiện trạng & gợi ý hướng cải thiện (để bàn với chủ dự án)
- Tông màu hiện tại: đen/trắng/xám tối giản (biến ở `:root` trong `styles.css`).
  Có thể nâng cấp: thêm 1 màu nhấn (accent) nhẹ, bo góc mềm hơn, đổ bóng tinh tế,
  khoảng cách (spacing) thoáng hơn, font dễ đọc.
- Thanh công cụ (toolbar/meta-bar) khá dày đặc — có thể nhóm lại gọn gàng, thêm
  icon rõ nghĩa, trạng thái rõ ràng (đã lưu / đang đồng bộ).
- Bảng P-FMEA: cải thiện màu header, đường kẻ, ô đang sửa, ô tự động điền cho
  tương phản dễ chịu; giữ khả năng cuộn ngang.
- Tab Hướng dẫn: trau chuốt typography, bảng tiêu chuẩn, thẻ 4M.
- Cân nhắc **responsive** tốt hơn (màn nhỏ).
- **Quan trọng:** hỏi chủ dự án thích phong cách nào (vd: tối giản kiểu Notion /
  hiện đại kiểu phần mềm SaaS / đậm chất công nghiệp) trước khi làm lớn.

### Việc còn treo khác (chủ dự án sẽ cấp file sau)
- **Bảng tiêu chuẩn Tần suất O** và **bảng Phát hiện D** trong tab Hướng dẫn hiện
  là **bảng tham khảo** (có banner cam đánh dấu). Thay bằng bảng chính thức khi có.
- **Bảng "Tiêu chuẩn thực hiện biện pháp đề xuất"** — đã chừa chỗ ở trang cột M.
- (Tùy chọn) Edge Function để dùng chung 1 API key Gemini; lớp đăng nhập giới hạn nội bộ.

---

## 7. Bài học / lỗi đã từng gặp (tránh lặp lại)

- **Mất dữ liệu khi mở:** đừng xóa autosave hay "dọn dữ liệu cũ" khi load. Mở trang
  mới = bảng trống; chỉ khôi phục khi chọn đúng model base.
- **Excel "We found a problem":** template đã có sẵn thẻ `<headerFooter/>` — phải
  **thay thế**, không được thêm thẻ thứ hai. Luôn kiểm tra XML hợp lệ.
- **Excel mất cột bên phải / chữ bị cắt:** dùng `fitToWidth="1"` + `fitToPage="1"`
  (Excel tự co vừa khổ), và phân bổ phần thiếu chiều cao dòng cho mọi dòng trong ô
  gộp (đừng dồn hết vào dòng 1 — vượt giới hạn 409pt).
- **Parser cắt tên công đoạn:** đã sửa bằng `vnText()` nối các dòng tiếng Việt.
- **GitHub Pages 404:** URL phân biệt hoa/thường; cần `.nojekyll`.
- **Build mất thẻ Supabase CDN:** regex build chỉ gỡ script `src` nội bộ (đã sửa).
- **Gemini 429 limit:0:** model hết quota free → cơ chế thử nhiều model.

---

## 8. Nhịp làm việc khuyến nghị mỗi phiên

```bash
git pull                 # lấy bản mới nhất trước khi sửa
# ... sửa code ...
node tools/build.js      # build lại file độc lập
# ... test JSDOM nếu cần ...
git add -A && git commit -m "Mô tả thay đổi"
git push                 # web tự cập nhật qua GitHub Pages
```
