# 📘 HƯỚNG DẪN SỬ DỤNG VÀ CHẠY DỰ ÁN — ADS TRADING SYSTEM

Tài liệu này giải thích nhiệm vụ từng chi tiết thư mục và mô tả các thành phần script hiện hữu trong phiên bản mới nhất, cùng pipeline thứ tự chạy chuẩn.

---

## 🗂️ CẤU TRÚC TỔNG QUAN

```
TEST/
├── RULE&TEMPLATE/         ← Dữ liệu tham chiếu, Master File gốc.
├── A_create_campaign/     ← (Cũ) Tạo campaign nhỏ lẻ từ báo cáo Search Term.
├── B_update_campaign/     ← CLI tương tác thay đổi State/Bid campaign thủ công.
├── C_mass_sop_factory/    ← 🏭 Factory tạo mới hàng loạt Campaign Amazon.
├── D_auto_bulk_updater/   ← 🤖 Hệ thống Tracker theo dõi Index & update bulk theo quy luật.
├── E_debug_tools/         ← Bộ đo lường, xuất Logs, Validator Cấu Trúc (Tools độc lập).
└── F_core_engine/         ← JSON Rule Engine kết hợp JSON Schema xử lý Bulk linh hoạt.
```

---

## 📁 C_mass_sop_factory/
*Giai đoạn vận hành: Tạo Sản phẩm mới - Campaigns Generation*

### Nhiệm vụ
Hệ thống tự động hóa Pipeline 2 chặng: Lọc yêu cầu khởi tạo tự động từ Master file, nhân bản cấu trúc theo 3 dạng Match Type, và kết xuất file Excel đúng Template Bulk upload lên Amazon (hỗ trợ phân trang mỗi \~1.400 campaigns/file).

### Thứ tự chạy (Pipeline Khởi Tạo Mới)

1. **`excel_new_camp.py`**
   - Đọc: `data/input/PPC_NGUYÊN.xlsx` (Sheet: Listing)
   - Lọc các SKU có Trạng thái `Công việc mới` hoặc `Tiếp nhận`.
   - Kết quả: Áp dụng Cross Join (x3 Match Types) hoặc giữ nguyên nếu Ghi chú lớn hơn 3. Đóng gói ra `data/input/PPC_PROCESSED_OUTPUT.xlsx`.

2. **`mass_sop_factory.py`**
   - Đọc: `data/input/PPC_PROCESSED_OUTPUT.xlsx`
   - Nhập tham số cấu hình: Base Prefix Campaign, Budget hàng ngày, Date-Suffix, Store Target.
   - Kết xuất file cho Amazon: Các file trả ra có định dạng `Bulk_*.xlsx` với đầy đủ Record: `Campaign`, `Bidding Adjustment` (Top/Product), `Ad Group`, `Keyword`.

3. **`sync_master_file.py`**
   - Đọc ngược file Output và ghi đè lại file Master hiện hữu.
   - Tự động thay đổi Trạng thái SKU sang dòng `Chờ Upload`.
   - Lưu sang tên file `PPC_NGUYÊN_UPDATED.xlsx` đảm bảo không bị lỗi Lock file.

4. **`amz_schema_validator.py`**
   - Rà soát `Bulk_*.xlsx` vừa được sinh ra đem so với Template mẫu chuẩn, báo cáo Log nếu thừa thiếu các cột.

---

## 📁 D_auto_bulk_updater/
*Giai đoạn vận hành: Tối ưu và quản lý vòng đời*

### Nhiệm vụ
Đây là cốt lõi của hoạt động Tracking — tự động kéo Performance định kỳ, gắn vào sát bên cạnh chỉ số gốc làm dữ liệu so sánh cột (Horizontal Appending) đồng thời quyết định tắt/mở (Pause/Archive) khi chiến dịch hoạt động kém hiệu quả.

### Thứ tự chạy (Pipeline Vận Hành)

1. **Chuẩn bị Dữ liệu:**
   - Đưa Master file có sẵn vào thẻ `data/input 1/`.
   - Đưa các Bulk files (tải từ Amazon) của các Stores về lưu trữ trong cùng `data/input 1/`.
 
2. **Chạy `end_to_end_ppc_tracker.py`**
   - Đọc cấu hình Rule Engine tại `Rule_Engine.json` định nghĩa hạn chế lãng phí click/hiển thị.
   - Lọc điều kiện "Đã tạo campaign" từ Sheet Listing của Master File.
   - Gom nhóm Performance (Imp, Click, Order) theo `Campaign Name` từ file Amazon Bulk tải về.
   - Map và nối thêm các Metric theo thời gian vào sheet chỉ tiết của SKU.
   - Nếu từ khóa vi phạm chỉ số tối thiểu: Label action = "Archive/Pause".
   - **Kết Quả Output:** 
      - Master File được update `PPC_NGUYÊN_UPDATED.xlsx` (đưa vào input 2 hoặc output).
      - Output Upload: `Amazon_Bulk_Update_MMDD_TênStore.xlsx` chỉ chứa các lệnh chuyển State Keyword sang Pause.

3. **Chạy `auto_bulk_updater.py` (Tuỳ chọn bổ sung Update Bid)**
   - Sử dụng để tăng/giảm Bid (vd: +0.1 Bid cho Exact Keyword hoặc giảm % Bid Product Page xuống 0 cho nhóm Broad).

---

## 📁 F_core_engine/
*Công cụ Rule-Engine Độc lập*

### Nhiệm vụ 
Hệ điều hành chạy dựa hoàn toàn trên file `rules.json` – không cần Hard Code. Rất tốt nếu bạn cần Update 1 tham số như % ngân sách mà không quan tâm đến Code Python.

1. Chạy `core_engine.py` và chỉ định thư mục JSON Rule (ví dụ Tăng Bid keyword <1000 lượt Click).
2. Code quét qua dòng file Amazon và lưu ra `Final_Upload_Ready.xlsx`.

---

## 📁 Nhóm Thư Mục Khác
*Sử dụng trong kịch bản đặc thù*

- **`A_create_campaign/`**: Tiện ích lọc lấy Top Keywords xịn nhất từ File báo cáo để tự động khởi tạo Campaign mới mẻ chỉ chứa những Keyword chất lượng nhất (Sử dụng `run_pipeline.bat`).
- **`B_update_campaign/`**: Tool giao diện chọn số theo Menu (Command Line Input). Nhập tên Campaign tìm trên Tool trực tiếp -> Nhấn Phím số -> Sửa Ngân sách / Tên. Dùng để Update cho các Key mà Rule Engine không với tới.
- **`E_debug_tools/`**: Một số công cụ kiểm tra Data NaN, xem cấu trúc, không nên tham gia trực tiếp vào việc sản xuất Output Bulk.

---
🚀 *Chúc bạn quản lý hiệu quả! (Updated: Apr 2026)*
