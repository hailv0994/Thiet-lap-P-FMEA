# BÀN GIAO SESSION — P-FMEA Builder (cập nhật 2026-06-26)

> Đọc file này đầu tiên khi bắt đầu session mới.
> Đọc kèm: `CLAUDE.md` (quy tắc), `HANDOFF.md` (kiến trúc nền, Supabase/AI).

---

## 0. Thông tin nhanh

- **Dự án:** P-FMEA Builder — web 1 trang (offline), thiết lập P-FMEA tự động từ Control Plan (CP).
- **Branch phát triển + deploy:** `claude/upbeat-heisenberg-xpkxr3` (GitHub Pages chạy từ branch này).
- **URL:** https://hailv0994.github.io/Thiet-lap-P-FMEA/ (phân biệt hoa/thường). Push đúng branch → web tự cập nhật sau 2–5 phút. Người dùng cần **Ctrl+Shift+R** để bỏ cache.
- **Build file độc lập:** `node tools/build.js` → `P-FMEA-Builder.html` (~1.54 MB). **LUÔN build lại** sau khi sửa `index.html`/`styles.css`/`js/*.js`/`data/*.js`, rồi commit cả `P-FMEA-Builder.html`.
- **Giao tiếp:** tiếng Việt. Yêu cầu phức tạp/chưa rõ → **hỏi trắc nghiệm** trước khi code.
- **BẢO MẬT:** Không commit service_role/secret key. Chỉ nhúng anon/publishable key Supabase. API key Gemini chỉ lưu ở trình duyệt, không ghi vào code.

### Cấu trúc file
- `index.html` — bố cục (2 tab: Thiết lập P-FMEA / Hướng dẫn).
- `styles.css` — giao diện.
- `js/app.js` — state, dựng bảng, chấm S/O/D, lưu/mở (localStorage + Supabase), AI (Gemini), tab Hướng dẫn, autocomplete, gộp dạng hỏng.
- `js/parser.js` — đọc CP .xlsx (SheetJS), rút hạng mục chất lượng.
- `js/export-template.js` — xuất Excel (đổ vào template gốc, giữ định dạng, in A4).
- `data/template.js` (base64 template), `data/severity.js` (bảng S), `data/material.js` (cây Bộ phận→Sản phẩm→Dây chuyền).
- `vendor/` — `xlsx.full.min.js`, `fflate.min.js`.
- `tools/build.js` — gộp thành 1 file HTML độc lập.

### Kiểm thử bằng Node
SheetJS trong repo (`vendor/xlsx.full.min.js`) **không chạy được dưới Node thuần**. Khi test parser bằng Node:
```bash
npm install xlsx --prefix C:\temp\xlsx-tool   # Windows
# hoặc /tmp/xlsx-tool trên Linux/Mac
```
Rồi `require('C:/temp/xlsx-tool/node_modules/xlsx')` và copy logic từ `parser.js` ra script test.
Đọc file CP: `XLSX.readFile(path, { cellStyles: true })` — bắt buộc có `cellStyles: true` để phát hiện font Symbol.

---

## 1. Commit gần nhất (trạng thái hiện tại)

| Commit | Nội dung |
|--------|----------|
| `ab8a2f8` | fix(parser): "Max F"→"Max Ø"; ghép giá trị kề spec không có ±/-/+ vào spec |
| `38f30fa` | fix(parser): Symbol-font 'F'→Ø (cellStyles:true + applySymbolFont); spec nhiều ô (Bug D) |
| `660459f` | feat: mở rộng phủ định dạng hỏng + tự sửa dạng hỏng khi đổi ô Yêu cầu |
| `47f44cc` | fix: SC luôn theo CP — không kế thừa từ base/cũ qua merge/sync |

---

## 2. Logic quan trọng trong parser.js

### Ký hiệu đường kính Ø
Trong CP Excel, Ø thường gõ là **'F' với font Symbol**. 3 lớp xử lý:
1. `applySymbolFont(cell, val)` — nếu `cell.s.font.name` chứa "symbol" → đổi 'F' đứng độc lập thành 'Ø'. Gọi trong `cellRC`. Cần `cellStyles: true` khi `XLSX.read`.
2. `fixDiameter(s)` — fallback regex: đổi 'F' trước số/±, 'F' cuối chuỗi ("Max F"→"Max Ø"), 'F' đơn lẻ → 'Ø'. Chỉ bắt 'F' **không nằm trong từ chữ cái** → "FMEA", "F-type" không bị đổi.
3. Một số CP gõ sẵn Φ (phi Hy Lạp) hoặc Ø Unicode — giữ nguyên.

### Cột spec + cột kề
- **specCol** (cột "Giá trị tiêu chuẩn"): đọc gộp tất cả ô r0..r1 (nhiều dòng), dedup, nối bằng dấu cách.
- **Cột kề** (specCol+1 đến specCol+5): nếu có dấu ±/-/+ → dung sai (vào ngoặc); không có → ghép vào spec. Ví dụ: "Max"+"Rz12.5" → "Max Rz12.5"; "Max Ø"+"0.1" → "Max Ø 0.1".

### SC (đặc tính đặc thù)
- **Luôn đọc từ CP**, không kế thừa từ base/dữ liệu cũ.
- Đọc ĐÚNG 1 ô tại hàng r0, KHÔNG theo merge, KHÔNG quét xuống r1.
- Mọi chỗ cập nhật SC trong `app.js` (syncFromCP, mergeWithBase, applyBaseSuggest) đều dùng `nr.classification` (từ CP), không từ old/base data.

### Template CP hỗ trợ
- **GL SQS0811** (chuẩn): `parseSheet` — dò header theo tên cột.
- **GL SQS0831** (khối xếp dọc): `parseStacked` — fallback khi `parseSheet` không thấy cột tên.

---

## 3. Logic quan trọng trong app.js

### Dạng hỏng hóc (`failureModesFor`, `negateSpecText`)
- **Yêu cầu số giới hạn 1 phía** → "không đạt tiêu chuẩn" / "lớn/nhỏ hơn tiêu chuẩn".
- **Yêu cầu dung sai 2 phía** (±, +x/-y, khoảng ~) → tách 2 dạng hỏng; hiển thị giá trị danh nghĩa (bỏ dung sai) trong dạng hỏng.
- **Yêu cầu toàn chữ** → phủ định:
  - "phải được X" → "không được X"
  - "phải có X" → "không có X"
  - "phải X" → "không X"
  - "không được X" → "được X"
  - "không bị X" → "bị X"
  - "không có X" → "có X"
  - "không X" → "X"
  - fallback → "không " + câu gốc
- **Khi sửa ô Yêu cầu** → `syncFailureModeFromReq` tự cập nhật dạng hỏng, giữ nguyên các cột khác.

### SC (đặc tính đặc thù) trong app.js
- `reqFromBaseAnalysis`: `classification: nr.classification || ''` (không lấy từ `a.classification`).
- `syncFromCP` (ô khớp/đổi/mới): `matchOld.classification = nr.classification || ''`.
- `applyBaseSuggest`: không copy SC từ base.

### Autocomplete & autofill
- **Gõ ≥2 ký tự** → dropdown gợi ý (`showAutocomplete`), áp dụng cho 5 ô: `effectAnalysis`, `cause`, `prevention`, `detectCause`, `action`.
- **Nguyên nhân giống y hệt** → tự điền O, D, biện pháp đang trống (`autofillFromMatchingCause`, phạm vi toàn P-FMEA).
- **Ảnh hưởng giống** → tự điền câu kết luận + điểm S (`autofillFromMatchingEffect`).

### Gộp dạng hỏng hóc
- Gộp bất kỳ req nào trong cùng công đoạn.
- `detectOwn`: lưu câu phát hiện riêng của từng req để tách ý khi unmerge.
- `buildGroupDetect`: gom nhóm con cùng phương pháp+tần suất thành 1 ý.

---

## 4. Tính năng đã hoàn thiện (tham khảo SESSION-HANDOFF cũ nếu cần chi tiết)

- Tab Hướng dẫn: 2 chế độ (Hiểu & tư duy / Thao tác web), form P-FMEA trống song ngữ Việt–Nhật, bảng O/D/Action theo GL SQS0811 Bảng-2.
- Xuất Excel: rich text in đậm nhãn, ngắt trang A4 đúng, giãn đều hàng.
- Ô Sản phẩm & Dây chuyền: `<input list>` cho tự nhập hoặc chọn từ danh sách.
- Migration (`migrateState`): mở base cũ tự dựng lại dạng hỏng từ Yêu cầu.
- Build.js: regex CSS đã sửa để không bị mất `<style>` khi href có `?v=…`.

---

## 5. Việc đang cần làm

**Folder CP trên máy bạn:** `C:\Users\vieth\OneDrive\Desktop\PQC P2\PQC P2`

**Kiểm tra đã làm trong phiên cloud** (không cần làm lại):
- Đã parse toàn bộ `CP_HGJR_FP_GC.xlsx` bằng Node, so sánh với Excel, sửa khớp tất cả 9 sheet:
  Nhập phôi (2), GC1 (22), GC2 (17), GC3 (6), Khoan FP56 (3), Mài×3 (6 mỗi), Kiểm tra/Xuất hàng (6).
- Sheet "HGJR" là sơ đồ quy trình, không có hạng mục chất lượng → bỏ qua (đúng).

**Việc cần làm tiếp trên Desktop:**
1. Mở web, tải `CP_HGJR_FP_GC.xlsx`, nạp vào P-FMEA, xem kết quả trực tiếp trong UI — đối chiếu mắt với Excel để chắc không có gì lọt qua Node test.
2. **Test hồi quy** với các CP sản phẩm khác đã lưu trước đây — đảm bảo fix không phá dữ liệu cũ.
3. Nếu phát hiện thêm lỗi → sửa code, build lại, push.

---

## 6. Lệnh hay dùng

```bash
# Build file độc lập
node tools/build.js

# Push lên deploy
git push -u origin claude/upbeat-heisenberg-xpkxr3

# Kiểm tra cú pháp app.js
node -e "new Function(require('fs').readFileSync('js/app.js','utf8'))"
```
