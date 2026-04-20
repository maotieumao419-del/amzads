# 🚀 ADS TRADING SYSTEM — TỔNG HỢP HỆ THỐNG (SOP & SYNOPSIS)
> **Cập nhật:** 17/04/2026 (Tổng hợp từ dữ liệu 14/04/2026 đến nay)
> **Mục tiêu:** Tự động hóa vận hành Amazon PPC (Sponsored Products, Brands, Display).

---

## 1. NGỮ CẢNH VÀ LUỒNG XỬ LÝ (FOLDER & CODE)

Hệ thống được chia thành 6 Module chính (A-F) để quản lý vòng đời của một chiến dịch Amazon Ads.

### 📁 C_mass_sop_factory (Pipeline 1: Tạo mới hàng loạt)
Giải quyết vấn đề tạo hàng nghìn chiến dịch từ file kế hoạch thô.
- **`excel_new_camp.py`**:
  - **Vấn đề**: Lọc SKU cần chạy Ads từ Master File.
  - **Luồng**: Đọc `PPC_NGUYÊN.xlsx` -> Lọc SKU "Tiếp nhận/Công việc mới" -> Nhân bản Match Type (Exact, Phrase, Broad).
  - **Input**: `data/input/PPC_NGUYÊN.xlsx`
  - **Output**: `data/input/PPC_PROCESSED_OUTPUT.xlsx`
- **`mass_sop_factory.py`**:
  - **Vấn đề**: Build file chuẩn cấu hình Amazon (Campaign, Ad Group, Keyword, Bid Adjustment).
  - **Luồng**: Đọc Processed Output -> Map vào Mapping Sheet -> Xuất file Bulk.
  - **Input**: `data/input/PPC_PROCESSED_OUTPUT.xlsx`
  - **Output**: `Bulk_*.xlsx` (Ready to upload).
- **`sync_master_file.py`**:
  - **Vấn đề**: Cập nhật trạng thái "Chờ Upload" ngược lại Master File sau khi đã tạo xong file Bulk.
  - **Input**: `PPC_PROCESSED_OUTPUT.xlsx`
  - **Output**: `PPC_NGUYÊN_UPDATED.xlsx`.

### 📁 D_auto_bulk_updater (Pipeline 2: Tối ưu & Tracking)
Giải quyết vấn đề theo dõi hiệu suất và tự động tắt/mở/điều chỉnh bid.
- **`end_to_end_ppc_tracker.py`**:
  - **Vấn đề**: Nối ngang dữ liệu hiệu suất (Horizontal Appending) và tự động gắn nhãn "Pause/Archive" cho từ khóa kém.
  - **Luồng**: Quét Master File (SKU "Đã tạo campaign") + Quét file Bulk tải từ Amazon -> Nhóm data theo Campaign -> Nối cột `_Impression`, `_Click`, `_Order` vào sheet SKU -> Áp dụng Rule Engine.
  - **Input 1**: `PPC_NGUYÊN.xlsx` và các file `bulk-*.xlsx` từ Amazon.
  - **Input 2**: `PPC_NGUYÊN_UPDATED.xlsx` (File trung gian để code khác hoặc người dùng kiểm tra).
  - **Output**: `Amazon_Bulk_Update_*.xlsx` (File chỉ chứa lệnh Update để upload lên Amazon).
- **`auto_bulk_updater.py`**:
  - **Vấn đề**: Điều chỉnh Bid hàng loạt theo Portfolio hoặc từ khóa.
  - **Input 1**: File Bulk tải từ Amazon.
  - **Output**: File Update Bid/State.

### 📁 F_core_engine
- **`core_engine.py`**: Logic lõi xử lý file Excel dựa trên các quy tắc JSON (`rules.json`), giúp hệ thống linh hoạt không cần sửa code khi đổi logic kinh doanh.

---

## 2. MÔI TRƯỜNG VÀ THƯ VIỆN CHẠY CODE

Để chạy được toàn bộ hệ thống, máy tính cần cài đặt **Python 3.10+** và các thư viện sau:

- **Thư viện chính**:
  - `pandas`: Xử lý bảng biểu dữ liệu tập trung.
  - `openpyxl`: Đọc/ghi file Excel `.xlsx` (Bảo toàn định dạng).
  - `xlsxwriter`: Xuất file Excel với định dạng chuyên sâu (Text-stict).
  - `numpy`: Hỗ trợ tính toán mảng dữ liệu.
  - `re`, `os`, `glob`, `shutil`: Xử lý file, đường dẫn và Regex.

- **Kích hoạt môi trường (Windows)**:
  ```powershell
  # Đường dẫn venv: TEST\venv\
  .\venv\Scripts\activate
  ```

---

## 3. CẤU TRÚC THƯ MỤC BẮT BUỘC & Ý NGHĨA

Mỗi module (C, D) đều có cấu trúc `data/` riêng để cô lập dữ liệu:

### 📍 Thư mục `data/` trong Pipeline D (Cập nhật mới nhất):
1. **`data/input 1/`**: 
   - **Ý nghĩa**: Chứa dữ liệu "Đầu vào gốc". 
   - **Thành phần**: File `PPC_NGUYÊN.xlsx` và các file Bulk kết quả kinh doanh tải về từ Amazon.
2. **`data/input 2/`**: 
   - **Ý nghĩa**: Chứa dữ liệu "Đầu vào cho code khác". 
   - **Thành phần**: File `PPC_NGUYÊN_UPDATED.xlsx` (Kết quả sau khi tracker chạy xong, để dùng cho các bước tiếp theo hoặc lưu trữ).
3. **`data/output/`**: 
   - **Ý nghĩa**: Chứa kết quả cuối cùng để thực thi.
   - **Thành phần**: File `Amazon_Bulk_Update_*.xlsx` dùng để upload trực tiếp lên Seller Central.

### 📍 Thư mục `RULE&TEMPLATE/`:
- Chứa các file Template mẫu của Amazon để Validation và các file Mapping chuẩn SOP.

---

> [!IMPORTANT]
> **Lưu ý quan trọng về luồng dữ liệu:**
> Từ ngày **17/04/2026**, hệ thống ưu tiên sử dụng `Horizontal Appending` (Nối ngang). Thay vì ghi đè dữ liệu cũ, hệ thống sẽ tạo các cột mới theo tiền tố ngày tháng (ví dụ: `0415-0416_Click`) để giữ lại lịch sử hiệu suất trong Master File.
