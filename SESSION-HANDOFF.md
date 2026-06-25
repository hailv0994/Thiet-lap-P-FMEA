# BÀN GIAO SESSION — P-FMEA Builder (cập nhật mới nhất)

> File này tổng hợp TOÀN BỘ thay đổi trong phiên làm việc gần nhất, để chuyển sang
> session khác. Đọc kèm: [`CLAUDE.md`](./CLAUDE.md) (quy tắc) và [`HANDOFF.md`](./HANDOFF.md)
> (kiến trúc nền, cấu hình Supabase/AI). File này ưu tiên mô tả các tính năng MỚI.

---

## 0. Thông tin nhanh

- **Dự án:** P-FMEA Builder — web 1 trang (offline), thiết lập P-FMEA tự động từ Control Plan (CP).
- **Branch phát triển + deploy:** `claude/upbeat-heisenberg-xpkxr3` (GitHub Pages chạy từ branch này).
- **URL:** https://hailv0994.github.io/Thiet-lap-P-FMEA/ (phân biệt hoa/thường). Push đúng branch → web tự cập nhật sau vài phút.
- **Build file độc lập:** `node tools/build.js` → `P-FMEA-Builder.html` (~1.49 MB). **LUÔN build lại** sau khi sửa `index.html`/`styles.css`/`js/*.js`/`data/*.js`, rồi commit cả `P-FMEA-Builder.html`.
- **Giao tiếp:** tiếng Việt. Yêu cầu phức tạp/chưa rõ → HỎI TRẮC NGHIỆM trước khi code.
- **BẢO MẬT:** Không commit service_role/secret key. Chỉ nhúng anon/publishable key Supabase. API key Gemini chỉ lưu ở trình duyệt, không ghi vào code.
- **Cache-busting:** `index.html` gắn `?v=YYYYMMDDx` vào link css/js. Khi đổi version nhớ đổi đồng loạt. (Đã từng gây bug build — xem mục 9.)

### Cấu trúc file
- `index.html` — bố cục (2 tab: Thiết lập P-FMEA / Hướng dẫn).
- `styles.css` — giao diện.
- `js/app.js` — state, dựng bảng, chấm S/O/D, lưu/mở, AI, tab Hướng dẫn, autocomplete, gộp.
- `js/parser.js` — đọc Control Plan (.xlsx), rút hạng mục chất lượng.
- `js/export-template.js` — xuất Excel (đổ vào template gốc, giữ định dạng, in A4).
- `data/template.js` (template base64), `data/severity.js` (bảng S), `data/material.js` (cây Bộ phận→Sản phẩm→Dây chuyền).
- `vendor/` — `xlsx.full.min.js`, `fflate.min.js`.
- `tools/build.js` — gộp tất cả thành 1 file HTML độc lập.

### Kiểm thử bằng Node (đã dùng nhiều trong session)
- Parser: `global.window={}; global.XLSX=require('./vendor/xlsx.full.min.js'); require('./js/parser.js')` → `window.CPParser.parseSheet(wb, sheet)`.
- Export: load `fflate`, `data/template.js`, eval `export-template.js` → `global.TemplateExport.buildFromTemplate(state, templateBytes, fflate)`; mở lại bằng `XLSX.read` để kiểm tra XML hợp lệ (tránh Excel "We found a problem").
- Logic JS: `node -e "new Function(require('fs').readFileSync('js/app.js','utf8'))"` để check cú pháp.

---

## 1. Tab Hướng dẫn — viết lại hoàn toàn (2 chế độ)

**File:** `js/app.js` (vùng `gPagesThink`, `gPagesWeb`, `gFmeaBlankTable`, `setupGuide`, `renderGuidePage`), `index.html` (nút chế độ), `styles.css`.

- **Nút chuyển chế độ** trong tab Hướng dẫn: `🧠 Hiểu & tư duy đúng` / `🖱️ Thao tác trên web` (class `.guide-mode-btn`, `data-mode="think|web"`). `setupGuide()` có `loadMode(mode)`.
- **Chế độ "Hiểu & tư duy đúng"** (`gPagesThink`): 14 slide.
  - Slide 1: **form P-FMEA trống** dựng bằng HTML (`gFmeaBlankTable(hi)`) — song ngữ Việt–Nhật, đủ cột A→S, nhóm "Kết quả xử lý".
  - 13 slide cột (A→M): mỗi slide **tô sáng đúng 1 cột** trong form (`hiCol` → `renderGuidePage` render full bảng + cột vàng `.ff-hi`) + nội dung tư duy (🎯 Tư duy đúng / ✍️ Cách làm / ví dụ G823-00 / bảng tiêu chuẩn).
- **Chế độ "Thao tác trên web"** (`gPagesWeb`): 5 slide — B1 chọn Bộ phận/SP/Dây chuyền/Model/Model base; B2 tải CP & nạp; B3 thao tác từng cột (1–7, 8–13) kèm icon thật (`▲▼` đổi thứ tự, `🔗`/`🔓` gộp/tách, `✎ đổi` chọn câu kết luận, `＋ NN`/`✕ NN`/`✨` nguyên nhân & AI); B4 AI/Lưu/Xuất.
- Bảng tiêu chuẩn O/D/Action lấy từ PDF chính thức (GL SQS0811 Bảng-2): `gOccurrenceTable`, `gDetectionTable`, `gActionTable`, banner `G_OFFICIAL`.
- **Slide C (Ảnh hưởng) & D (S) đã sửa:** KHÔNG đánh số ①② song song. Cách đúng: phân tích ảnh hưởng thực tế XONG → mới **chọn câu kết luận** từ bảng tiêu chuẩn → điểm S tự điền từ câu kết luận đó.

---

## 2. Tự điền & gợi ý theo nội dung đã nhập (autocomplete + chooser)

**File:** `js/app.js`.

### 2a. Tự điền khi NGUYÊN NHÂN giống y hệt — `autofillFromMatchingCause(pid,rid,cid)`
- Khi rời ô nguyên nhân (hoặc chọn từ gợi ý), nếu nguyên nhân **giống y hệt** (so khớp `normKey`: bỏ hoa/thường + khoảng trắng) một nguyên nhân đã có → tự điền các ô đang **trống**.
- **Phạm vi: TOÀN BỘ P-FMEA** (đã đổi từ "cùng công đoạn"). 
- **Tự điền (single value):** điểm O (`occurrence`), điểm D (`detection`), biện pháp (`action`). (Đây là fix bug "ô điểm không tự nhảy".)
- Chỉ điền ô trống, không ghi đè.

### 2b. Tự điền khi ẢNH HƯỞNG giống — `autofillFromMatchingEffect(pid,rid)`
- Ô Ảnh hưởng giống y hệt một ô đã chọn câu kết luận (phạm vi **toàn bộ P-FMEA**) → tự điền **câu kết luận** (`effectStdText`, `effectScope`) + **điểm S** (`severity`), chỉ khi hiện chưa chọn.

### 2c. Gợi ý gõ-lọc (autocomplete) — `showAutocomplete(el, field)`
- Gõ ≥ 2 ký tự đầu trùng nội dung đã nhập → hiện dropdown (`.ac-pop`) để bấm chọn. Esc/click ngoài để đóng.
- Áp dụng cho **5 ô**: `effectAnalysis`, `cause`, `prevention`, `detectCause`, `action`.
- Nguồn: nội dung trong bảng hiện tại (`collectFieldValues`) + bộ nhớ câu đã lưu (`getSavedPhrases`, localStorage `LS_PHRASES`, theo cột+bộ phận, đồng bộ Supabase qua `cloudPushPhrase`).

### 2d. Chooser theo nguyên nhân — `showCauseOptions(el, field)` (focusin)
- Khi **bấm vào ô Dự phòng / Phát hiện ra nguyên nhân ① còn trống** mà nguyên nhân của dòng đó đã có dữ liệu ở nơi khác → hiện dropdown **các phương án đã dùng cho cùng nguyên nhân** (`collectCauseLinkedValues`, toàn bộ P-FMEA, gồm nhiều phương án khác nhau) để chọn nhanh — KHÔNG cần gõ.
- Lý do: prevention & detectCause có thể có nhiều phương án → để người dùng chọn, không tự điền âm thầm.
- Hàm dựng dropdown chung: `openACList(el, items, head, pickFn)`.

---

## 3. Tách dạng hỏng hóc & hiển thị spec (`failureModesFor`, `nominalSpec`)

**File:** `js/app.js`.

- **Dung sai 2 phía** → tách 2 dạng hỏng "lớn hơn / nhỏ hơn tiêu chuẩn". Nhận diện từ: cột `tol` riêng, spec dạng khoảng `~`, **HOẶC dung sai nằm lẫn trong spec** `±` / `+x/-y` (VD "216.5 ±1", "203.3(+0.3/-1.1)").
- **Chỉ hiển thị kích thước danh nghĩa, BỎ dung sai** trong dạng hỏng (`nominalSpec()` loại `±1`, `(+0.3/-1.1)`, `+0.3/-1.1`). VD "Kích thước chiều cao tổng **216.5** lớn hơn tiêu chuẩn" (không kèm ±1). Ô **Yêu cầu** vẫn giữ đầy đủ dung sai.
- **Yêu cầu phủ định "Không X" / "Không được X"** → dạng hỏng là khẳng định "**có X**". VD: yêu cầu "Không dò rỉ khí... dưới 490kPa" → dạng hỏng "Kiểm tra dò rỉ khí: **có** dò rỉ khí... dưới 490kPa". (Trước đó bị thêm "không đạt" vì spec có chứa số.)
- **Migration khi load** (`migrateState`): base cũ tự dựng lại dạng hỏng từ ô Yêu cầu (`parseReqText` → `failureModesFor`) — nên các fix trên tự áp dụng cho base đã lưu khi mở lại; nhóm đã tách cũng được chuẩn hóa lại (bỏ dung sai).

---

## 4. Gộp dạng hỏng hóc — viết lại (`buildGroupDetect`, `normalizeMergeGroups`)

**File:** `js/app.js`.

- **Bỏ giới hạn**: gộp **bất kỳ dạng hỏng hóc nào** trong **cùng công đoạn** (trước chỉ cho gộp khi cùng dụng cụ+tần suất). `openMergePop` candidates = mọi req khác trong proc.
- **Trường mới `detectOwn`**: lưu câu phát hiện RIÊNG (phương pháp+tần suất gốc) của từng dạng hỏng, để gộp không làm mất thông tin (cần khi tách ý). Khởi tạo = `detectAuto(item)`. Migration backfill cho data cũ.
- **`buildGroupDetect(members)`** dựng nội dung "Phát hiện ra dạng hỏng hóc" (lưu SẠCH, mỗi ý 1 dòng):
  - Nhóm con CÙNG phương pháp+tần suất (cùng `suffix` của `parseDetect`) → 1 ý "Kiểm tra các <chung> bằng… theo tần suất…" (`summarizeNames`).
  - Nhóm con KHÁC nhau → nhiều ý, mỗi ý 1 dòng.
- **`normalizeMergeGroups(p)`**: dựng lại câu phát hiện cho mọi nhóm; nhóm còn 1 thành viên thì tự giải tán (khôi phục `detectOwn`). Gọi sau merge/unmerge và trong migration.
- **Giữ cột chung khi gộp**: nếu đại diện (ô gộp-vào) đang TRỐNG mà thành viên có → lấy của thành viên cho: **đặc tính đặc thù (classification)**, ảnh hưởng + câu kết luận + điểm S. (Fix bug "gộp làm mất đặc tính đặc thù".)

---

## 5. Cột "Phát hiện ra" — định dạng cuối cùng (web + Excel)

**File:** `js/app.js` (`detectCellHTML`, `fmtDetectFailure`), `js/export-template.js`, `styles.css`.

Định dạng chốt (theo ảnh mẫu người dùng gửi):
```
Phát hiện ra nguyên nhân:        <- NHÃN IN ĐẬM, dòng riêng
<nội dung người dùng nhập>
Phát hiện ra dạng hỏng hóc:      <- NHÃN IN ĐẬM, dòng riêng
-<ý 1>                            <- mỗi ý 1 dòng, prefix "-"
-<ý 2>
```
- Nhãn **in đậm** (web: `.detect-label{font-weight:700}`; Excel: **rich text run** có `<b/>`, font Arial 10 — xem mục 8).
- "Phát hiện ra dạng hỏng hóc": mỗi ý xuống dòng, prefix `-` (`fmtDetectFailure()` — bỏ prefix cũ, prefix lại "-").
- **Lưu trữ:** `detectFailureAuto` lưu dạng SẠCH (mỗi ý 1 dòng, không prefix). Prefix "-" thêm khi hiển thị/xuất. Ô ② trên web giờ **read-only** (nội dung tự động).
- **Nhiều nguyên nhân trong 1 ô** (1 hay nhiều dạng hỏng gộp): nguyên nhân ĐẦU hiện đầy đủ; các nguyên nhân SAU phần "Phát hiện ra dạng hỏng hóc" = **"Giống với nội dung trên"** (viết hoa). Dựa vào index `ci` của cause (`detectCellHTML(p,r,c,ci)`).
- ① "Phát hiện ra nguyên nhân" (`detectCause`) vẫn nhập tay riêng từng nguyên nhân.

---

## 6. Parser — hỗ trợ template GL SQS0831 (khối xếp dọc) — `parseStacked`

**File:** `js/parser.js`.

- **Bug đã sửa:** CP theo **GL SQS0831** đặt khối "Quản lý đặc tính chất lượng" **xếp DỌC** (dưới "Quản lý điều kiện chế tạo", **cùng cột**). Logic cũ dò theo cột-biên (`colMin`=cột đặc tính, `colMax`=cột điều kiện) → khoảng cột rỗng → không thấy "Hạng mục quản lý" → báo "Không tìm thấy công đoạn nào có hạng mục chất lượng".
- **`parseStacked(ws, merges, sheetName, processName)`** = FALLBACK, chỉ chạy khi `nameCol < 0` (cách cũ thất bại → KHÔNG ảnh hưởng template cũ):
  - Dò header theo **nhãn khối "Quản lý đặc tính" + span cột (từ merge)**, tiêu đề nằm DƯỚI nhãn (`row > qRow`).
  - Cột giá trị tên header là **"Giá trị quản lý"** (khác "Giá trị tiêu chuẩn" của 0811). Tách spec/dung sai bằng cách quét các ô con trong vùng [specCol, methodCol), bỏ "R/L", "Max/Min", phân loại dung sai (có `±+-`) vs trị số.
  - Nhận hạng mục theo **dòng tên tiếng Việt** (`VN_DIACRITIC`), bỏ dòng dịch tiếng Anh riêng (tránh nhân đôi).
  - Bound dưới: dòng trước "Quản lý điều kiện"/"Hạng mục cấm"/"Following rules".
- **Kiểm thử file thật** (`HGJR Line 8 (SXHL)…CHECK.xlsx`): 17/17 sheet, 153 hạng mục — đọc đúng tên/spec/tol/method/freq.
- **Lưu ý còn lại:** spec/tol của template này tách thành nhiều ô con khá phức tạp; có thể vài hạng mục spec chưa thật chuẩn → người dùng chỉnh tay hoặc báo để tinh chỉnh thêm.

---

## 7. Ô Sản phẩm & Dây chuyền — cho tự nhập (input + datalist)

**File:** `index.html`, `js/app.js`.

- Đổi `<select id="mProduct">` / `<select id="mLine">` thành `<input list="dlProduct">` / `<input list="dlLine">` + `<datalist>` → chọn từ danh sách Material có sẵn HOẶC tự gõ giá trị mới.
- `fillProduct`/`fillLine` nay dùng `fillInputList(inputId, listId, items, keep)` (đổ datalist + set value).
- **Bộ phận (`mDept`) vẫn là `<select>`** (chỉ Sản phẩm & Dây chuyền tự nhập).
- Handler `change` cũ vẫn dùng `.value` nên hoạt động với input. **KHÔNG thêm listener `input`** cho product/line vì sẽ phá logic `confirmContext` (product/line là "khóa" context; cập nhật meta mỗi keystroke làm hỏng revert khi người dùng hủy).

---

## 8. Xuất Excel — in đậm nhãn (rich text)

**File:** `js/export-template.js`.

- Ô trước đây ghi `t="inlineStr"` thường → không đậm được từng phần.
- Thêm hỗ trợ **rich text**: `genCell` render `cell.rich` (mảng run `{t, b}`) thành `<is><r><rPr>…</rPr><t>…</t></r>…</is>`. Run đậm: `<rPr><rFont val="Arial"/><family val="2"/><b/><sz val="10"/></rPr>` (đúng thứ tự schema CT_RPrElt).
- `putRich(r,c,plain,runs)`: lưu `cell.v=plain` (để tính chiều cao/độ rộng) + `cell.rich=runs`.
- Ô cột "Phát hiện ra" (col 10): nhãn "Phát hiện ra nguyên nhân/dạng hỏng hóc/Bổ sung…" là run **đậm**, nội dung run thường.
- **Đã kiểm thử:** tạo file thật → mở lại bằng `XLSX.read` hợp lệ (không lỗi "We found a problem"); XML chứa đúng các run đậm.

---

## 9. Bug build.js đã sửa (quan trọng)

- `tools/build.js` inline CSS bằng regex khớp `<link rel="stylesheet" href="styles.css" />`. Khi `index.html` thêm `?v=…` vào href thì regex **không khớp** → file `P-FMEA-Builder.html` (bản offline) bị **mất CSS** (còn link ngoài, không có `<style>`).
- Đã sửa regex thành `/<link rel="stylesheet" href="styles\.css[^"]*"\s*\/>/`. Nếu sau này đổi cách gắn version, kiểm tra lại build có inline `<style>` không.

---

## 10. Triển khai (deploy) & lưu ý môi trường

- **GitHub Pages** từ branch `claude/upbeat-heisenberg-xpkxr3`. `git push -u origin claude/upbeat-heisenberg-xpkxr3` → web cập nhật sau 2–5 phút. Người dùng cần **Ctrl+Shift+R** (hard refresh) để bỏ cache.
- File `.nojekyll` BẮT BUỘC giữ (để Pages phục vụ thư mục `vendor/`, `data/`).
- (Lịch sử) Trước dùng Netlify (pfmea-tool.netlify.app) nhưng **hết credit** → đã chuyển sang GitHub Pages.
- Commit message kết thúc bằng:
  ```
  Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>
  Claude-Session: https://claude.ai/code/session_01VLQoYMFTz9x9FBrvnBUTJW
  ```

---

## 11. Việc có thể làm tiếp / điểm cần để ý

- Spec/tol của template **GL SQS0831** (mục 6) là heuristic — nếu người dùng gặp hạng mục spec/dung sai sai, xin file/ảnh để tinh chỉnh `parseStacked`.
- Các kiểu yêu cầu phủ định khác ngoài "Không X"/"Không được X" (VD "chưa…", "tránh…") chưa xử lý — bổ sung khi có yêu cầu.
- `reqSig` trong `parser.js`/`app.js` có thể còn sót chỗ không dùng (đã bỏ giới hạn gộp) — vô hại.
- Khi sửa bất kỳ file nguồn: **build lại `P-FMEA-Builder.html`** và commit kèm.
- Dữ liệu ví dụ minh họa trong tài liệu/code dùng **G823-00** (PRO2 — giảm xóc, công đoạn HÀN CAP) và file CP thật **HGJR Line 8 (SXHL)** (template 0831).
