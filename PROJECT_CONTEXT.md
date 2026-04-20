# 📦 ADS TRADING SYSTEM — PROJECT CONTEXT

> **Tác giả:** nnh16  
> **Thư mục gốc:** `c:\Users\nnh16\ads-trading-system\TEST\`  
> **Mục tiêu tổng thể:** Tự động hóa toàn diện quy trình Amazon PPC Ads — từ việc chuẩn bị từ khóa, tạo hàng loạt chiến dịch (Campaign Generation), đồng bộ dữ liệu vào Master File, đến việc theo dõi hiệu suất và cập nhật giá thầu, trạng thái hoàn toàn tự động.

---

## 📁 CẤU TRÚC THƯ MỤC CHÍNH

```
TEST/
├── RULE&TEMPLATE/         ← Chứa các file tham chiếu, template chuẩn và file Master (Dữ liệu nguồn)
├── A_create_campaign/     ← (Cũ) Tạo campaign SP mới từ report.
├── B_update_campaign/     ← Cập nhật campaign thủ công qua giao diện CLI tương tác.
├── C_mass_sop_factory/    ← 🚀 Pipeline 1: Tạo hàng loạt Campaign mới (2-stage file processing).
├── D_auto_bulk_updater/   ← 🚀 Pipeline 2: Tracking hiệu suất End-to-End & Cập nhật Bulk hàng loạt.
├── E_debug_tools/         ← Các script rà soát dữ liệu (Data Explorer, Schema Validator).
├── F_core_engine/         ← Rule engine lõi được điều khiển bằng file rules JSON (Quy tắc logic).
├── PROJECT_CONTEXT.md     ← Ý nghĩa dự án, tài liệu tổng quan.
└── PROJECT_GUIDE.md       ← Hướng dẫn chạy và giải thích từng folder/script.
```

---

## 🔄 CÁC PIPELINE PHỤC VỤ KINH DOANH CHÍNH

Hệ thống hiện tại xoay quanh 2 quy trình tự động hóa lớn.

### 1. Vòng lặp Tạo Chiến Dịch Mới (Campaign Generation - Thư mục C)

**Mục tiêu:**  
Tạo hàng loạt (Mass creation) hàng nghìn chiến dịch Amazon chuẩn SOP từ file Master Planning.

**Quy trình (Nhiều chặng):**
1. **Tiền xử lý (excel_new_camp.py):** Quét file Master (`PPC_NGUYÊN.xlsx`) trong sheet `Listing`, tìm các SKU có trạng thái "Tiếp nhận" hoặc "Công việc mới". Trích xuất từ khóa, nhân bản qua các Match Type (Exact, Phrase, Broad) hoặc giữ nguyên nếu vượt quá số lượng ghi chú tự định dạng. Xuất kết quả xử lý ra `PPC_PROCESSED_OUTPUT.xlsx`.
2. **Tạo File Tham số Amazon (mass_sop_factory.py):** Từ dữ liệu đã làm sạch ở bước 1, mapping cấu trúc chuẩn 1 chiều cho file Amazon BulkSheet (Campaign, Ad Group, Keyword, Bidding Adjustment). Bổ sung biến số Start Date, tạo file `Bulk_*.xlsx` để tải trực tiếp lên Amazon Seller Central.
3. **Đồng bộ hóa Master File (sync_master_file.py):** Đọc ngược từ `PPC_PROCESSED_OUTPUT.xlsx`, đổi status của các SKU từ "Công việc mới" thành "Chờ Upload" và ghi đè nội dung quay lại cập nhật cho Master sheet thành file `PPC_NGUYÊN_UPDATED.xlsx`.
4. **Kiểm tra Schema (amz_schema_validator.py):** Validator đối chiếu số lượng cột, tên cột file đã build so với template chuẩn tránh lỗi khi upload lên Amazon.

### 2. Vòng lặp Theo dõi & Cập nhật Hiệu Suất (End-To-End Tracker - Thư mục D)

**Mục tiêu:**  
Tracking trực tiếp hiệu suất Impressions, Clicks, Orders... từ các file Bulk tải từ Amazon xuống và đối chiếu với Ruleset kinh doanh để đưa ra các Action phù hợp (Pause, Archive, tăng giảm bid).

**Quy trình:**
1. **End_to_End_PPC_Tracker (end_to_end_ppc_tracker.py):**
   - Đọc các sku có trạng thái "Đã tạo campaign" từ Master File.
   - Bóc tách file thực tế Amazon `bulk-*.xlsx`, tự nhận diện store (`LHHKMT`, `Musemory`) và khoảng thời gian (DateRange).
   - Horizontal Appending: Nối ngang dữ liệu vào các sheet SKU chi tiết (Các cột theo ngày `_Impression`, `_Click`, `_Order`).
   - Rule Engine Xử Lý (sử dụng `Rule_Engine.json`): Nếu `Click >= 11` & `Order = 0` → Gắn nhãn "Archive/Pause". Nếu `Impressions >= 1000` & `Click = 0` → Cần sửa từ khóa.
   - Ghi đè vào `PPC_NGUYÊN_UPDATED.xlsx` (bản Master File cập nhật phiên bản mới).
   - Xuất file giải quyết xử lý: `Amazon_Bulk_Update_*.xlsx` chỉ chứa các dòng Entity (Keyword) cần "Update" thành Paused/Archived để đẩy lên Amazon.

---

## ⚙️ MÔI TRƯỜNG CÀI ĐẶT

### Python & Virtual Environment
```powershell
# Virtual environment đã tạo sẵn tại:
c:\Users\nnh16\ads-trading-system\TEST\venv\

# Cách activate:
venv\Scripts\activate

# Các thư viện chính:
pip install pandas openpyxl xlsxwriter tabulate numpy
```

### Chạy script ở bất kỳ đâu (gọi thẳng python từ venv):
```powershell
venv\Scripts\python.exe D_auto_bulk_updater\end_to_end_ppc_tracker.py
```

### Lỗi hiển thị Tiếng Việt trên Windows (Terminal):
```powershell
$env:PYTHONIOENCODING="utf-8"
venv\Scripts\python.exe C_mass_sop_factory\excel_new_camp.py
```

---

## 📊 CẤU TRÚC FILE BULK AMAZON (Quy tắc chuẩn)

File Bulk Amazon chuẩn cần đảm bảo:
- Các cột Editable (Có thể xử lý): `Operation` (update/create/archive), `Entity` (Campaign/Keyword/Bidding Adjustment...), `State` (enabled/paused), `Bid`, `Percentage`.
- Các cột Read Only / Informational Only: `Impressions`, `Clicks`, `Spend`... Những cột này bị tự làm mờ đi bởi Seller Central và sửa chúng trong Excel cũng không gây lỗi trong quá trình Upload.

**Nền tảng kiến trúc mới:** Đã cô lập Master Data File riêng. Mọi chỉnh sửa được thực hiện trên Copy Versions (như `PPC_NGUYÊN_UPDATED.xlsx`) và sử dụng hệ thống Checksum Validation (trong `amz_schema_validator.py`).

---

*Cập nhật lần cuối thay đổi luồng xử lý: Ngày 17 Tháng 04 Năm 2026*
