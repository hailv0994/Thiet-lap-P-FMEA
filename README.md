# P-FMEA Builder

Công cụ web **chạy hoàn toàn trên máy tính cá nhân** (offline) để thiết lập
bảng **P-FMEA** từ dữ liệu **Control Plan (CP)**. Web tự rút dữ liệu từ CP,
điền sẵn các cột suy ra được, để bạn bổ sung phần tự phân tích rồi xuất ra
file Excel đúng format P-FMEA.

> Màu sắc tối giản đen / trắng / xám. Không cần internet, không cần cài đặt.

---

## Cách chạy

Nhấp đúp vào **`index.html`** để mở bằng trình duyệt (Chrome / Edge / Firefox).
Không cần server, không cần internet — mọi thư viện đã được nhúng sẵn trong
thư mục `vendor/`.

## Cách dùng

1. **Tải Control Plan** → chọn file `.xlsx` của CP.
2. Chọn **sheet công đoạn** (mỗi sheet thường là 1 công đoạn), nhập **STT công
   đoạn**, rồi bấm **＋ Nạp vào P-FMEA**.
3. Bảng P-FMEA hiện ra với các cột **tự động** điền từ CP. Bổ sung các phần
   **tự phân tích** và chọn ô **ảnh hưởng ②** để tự chấm điểm **S**.
4. Bấm **⬇ Xuất Excel** để lưu file `P-FMEA_<ngày>.xlsx`.

Có thể nạp nhiều công đoạn (mỗi lần 1 sheet) vào cùng một bảng, hoặc bấm
**＋ Công đoạn trống** để tự nhập tay.

### Lưu & mở lại

- Điền **Bộ phận / Sản phẩm / Dây chuyền / Model** ở thanh trên cùng, rồi bấm
  **💾 Lưu** — dữ liệu được lưu trong trình duyệt theo Model.
- Web **tự lưu** mỗi khi nhập; mở lại trang sẽ khôi phục phiên gần nhất.
- Chọn ở **Mở dự án đã lưu** để tải lại một Model đã lập trước đó.
- **⬇ Sao lưu / ⬆ Nạp sao lưu**: xuất/nhập toàn bộ dự án ra file `.json` để
  chuyển sang máy khác hoặc lưu trữ.

### Nguyên nhân (4M)

Mỗi nguyên nhân chọn nhóm **Man / Machine / Method / Material** trước khi phân
tích; khi xuất Excel sẽ in kèm (vd: `Man: cài đặt điện áp sai`).

---

## Quy tắc ánh xạ dữ liệu (CP → P-FMEA)

| Cột P-FMEA | Nguồn / cách điền |
|---|---|
| **Quy trình / Bước / Chức năng** (A) | Tên công đoạn + yêu cầu (hạng mục quản lý + giá trị tiêu chuẩn) lấy **tự động** từ CP. Chức năng nhập tay. |
| **Dạng hỏng hóc** (B) | **Tự động** = phủ định của yêu cầu (`<tên hạng mục> không đạt`). Sửa được. |
| **Ảnh hưởng** (C) | ① **tự phân tích** (nhập tay) + ② **chọn** từ tiêu chuẩn đánh giá S (sản phẩm / công đoạn). |
| **Mức độ nghiêm trọng S** (D) | **Tự chấm** theo Rank ứng với nội dung ② đã chọn. |
| **Phân loại** (E) | Lấy ký hiệu đặc tính đặc thù (S.C) từ CP nếu có; sửa được. |
| **Nguyên nhân** (F) | Tự phân tích (nhập tay). Mỗi yêu cầu có thể có **nhiều nguyên nhân**. |
| **Tần suất phát sinh O** (H) | Nhập tay (1–10). |
| **Quản lý dự phòng** (I) | Tự phân tích (nhập tay). |
| **Quản lý phát hiện ra** (J) | ① phát hiện nguyên nhân (nhập tay) + ② phát hiện dạng hỏng hóc **tự động**: `Kiểm tra <yêu cầu> bằng <phương pháp> theo tần suất <tần suất>`. |
| **Phát hiện D** (K) | Nhập tay (1–10). |
| **RPN** (L) | **Tự tính** = S × O × D. |

Khi một dạng hỏng hóc có **nhiều nguyên nhân**, các cột **Dạng hỏng hóc /
Ảnh hưởng / S / Phân loại** sẽ tự **gộp ô (merge)** theo số nguyên nhân —
cả trên web lẫn khi xuất Excel. Mỗi nguyên nhân + dự phòng + phát hiện (2 ý)
nằm trong **một ô** Excel.

---

## Ghi chú về file Excel xuất ra

File Excel xuất ra **giữ nguyên 100% định dạng** của form P-FMEA gốc: viền,
màu nền, font, vùng gộp ô (merge), độ rộng cột, tiêu đề song ngữ… Cách làm:
dữ liệu được **đổ thẳng vào file template gốc** (`FORM_NOI_DUNG_PFMEA.xlsx`
nhúng sẵn trong app), chỉ thay giá trị các ô và thêm vùng merge cho dữ liệu,
**không dựng lại** bảng — nên không mất bất kỳ định dạng nào.

- Dữ liệu được ghi vào sheet **FORMAT**; sheet **VÍ DỤ** giữ nguyên để tham khảo.
- Khi một dạng hỏng hóc có nhiều nguyên nhân, các cột Dạng hỏng / Ảnh hưởng /
  S / Phân loại được gộp ô tương ứng.

## Cấu trúc thư mục

```
index.html          Giao diện chính
styles.css          Định dạng (tối giản đen/trắng/xám)
data/severity.js    Bảng tiêu chuẩn đánh giá mức độ nghiêm trọng (S, Rank 1–10)
js/parser.js        Đọc & rút dữ liệu từ Control Plan
js/app.js           State, dựng bảng P-FMEA, chấm điểm S, xuất Excel
vendor/             Thư viện SheetJS (đọc/ghi Excel) — offline
samples/            File mẫu: CP, FORM P-FMEA, tiêu chuẩn đánh giá S
```
