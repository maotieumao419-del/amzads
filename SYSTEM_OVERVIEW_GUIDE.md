# 📦 ADS TRADING SYSTEM - Phân loại & Hướng dẫn cho người mới

Chào mừng bạn đến với hệ thống **ADS TRADING SYSTEM**! Hệ thống này được sinh ra để tự động hóa toàn diện quy trình Amazon PPC Ads — từ việc chuẩn bị từ khóa, tạo hàng loạt chiến dịch, đồng bộ dữ liệu vào Master File, đến việc theo dõi hiệu suất và cập nhật giá thầu, trạng thái một cách hoàn toàn tự động.

Dưới đây là bảng phân loại, thống kê dữ liệu hiện tại trong hệ thống, cùng với hướng dẫn sử dụng từng bước (step-by-step) dành cho người mới.

---

## 1. 🗂️ Phân Loại và Thống Kê Thư Mục

Dựa trên chức năng, các thư mục con trong dự án được chia làm 4 nhóm chính. *(Thống kê kích thước và số lượng file được cập nhật mới nhất)*:

### Nhóm 1: Core Pipelines (Vòng lặp Vận Hành Chính)
> Nơi diễn ra các hoạt động cốt lõi của hệ thống (Tạo mới và Quản lý vòng đời Campaign)

| Thư mục | Số lượng file | Dung lượng | Chức năng chính |
|---------|---------------|------------|-----------------|
| `C_mass_sop_factory` | 13 files | 0.44 MB | **Pipeline 1**: Chuyên tạo mới hàng loạt Campaign Amazon từ file Excel. Xử lý nhiều chặng và xuất Bulk upload. |
| `D_auto_bulk_updater`| 173 files| 13.64 MB| **Pipeline 2**: Theo dõi (Tracker) hiệu suất, nối ngang dữ liệu metrics và áp dụng tự động các điều kiện (pause/archive). |
| `F_core_engine` | 2 files | 0.01 MB | Chứa Rule Engine lõi (`rules.json`). Xử lý Bulk linh hoạt không cần sửa code Python. |
| `Store_Musemory` | 127 files | 7.45 MB | Dữ liệu/Tracker cụ thể chạy riêng cho store **Musemory**. |
| `Store_LHHMKT` | 79 files | 6.90 MB | Dữ liệu/Tracker cụ thể chạy riêng cho store **LHHMKT**. |

### Nhóm 2: Tools & Utilities (Công Cụ và Cập Nhật Thủ Công)
> Các script nhỏ để xử lý những trường hợp đặc thù ngoài Rule Engine.

| Thư mục | Số lượng file | Dung lượng | Chức năng chính |
|---------|---------------|------------|-----------------|
| `A_create_campaign` | 4 files | 0.02 MB | (Bản cũ) Tiện ích lọc từ khóa từ file báo cáo Search Term và khởi tạo campaign thủ công. |
| `B_update_campaign` | 1 files | 0.02 MB | Chạy CLI chọn số, dùng để đổi ngân sách, state thủ công cho các key lọt ngoài phạm vi auto. |
| `E_debug_tools` | 7 files | 0.01 MB | Kiểm tra dữ liệu bị lỗi (NaN), xuất logs, Validator rà soát cấu trúc. |

### Nhóm 3: Data & Template (Dữ liệu Tham Chiếu và Lưu Trữ)
> Các file gốc và nơi chứa kết quả đầu vào/đầu ra.

| Thư mục | Số lượng file | Dung lượng | Chức năng chính |
|---------|---------------|------------|-----------------|
| `RULE&TEMPLATE` | 4 files | 1.40 MB | Nơi chứa các file Excel Template chuẩn, dữ liệu Master tham chiếu gốc. |
| `data` / `output` / `separated_campaign_files` | \~ 5 files | 0.03 MB | Các folder chứa file input gốc, file trung gian (separated) và file kết xuất đầu ra. |

### Nhóm 4: Môi trường & Hệ thống khác
> Các tệp tin cấu hình không trực tiếp liên quan đến quy trình Ads.

| Thư mục | Số lượng file | Dung lượng | Chức năng chính |
|---------|---------------|------------|-----------------|
| `UI` | 13343 files| 130.22 MB| Mã nguồn giao diện quản trị (React/JS/TS). |
| `venv` | 8464 files | 137.50 MB| Môi trường ảo Python của dự án. |
| `Test5`, `scratch` | 69 files | 3.70 MB | Các script, notebook chạy nháp, test riêng lẻ. |

---

## 2. 🚀 Hướng Dẫn Sử Dụng (Dành Cho Người Mới)

> [!TIP]
> **Yêu cầu trước khi chạy**:
> - Đảm bảo mở terminal trong đúng thư mục dự án (`c:\Users\nnh16\ads-trading-system\TEST`).
> - Luôn kích hoạt môi trường ảo (venv) trước khi chạy bất kỳ lệnh nào.
> 
> ```powershell
> venv\Scripts\activate
> $env:PYTHONIOENCODING="utf-8"
> ```

Hệ thống xoay quanh 2 quy trình làm việc (Pipeline) chính tương ứng với vòng đời của 1 sản phẩm.

### Quy Trình 1: Tạo Hàng Loạt Campaign Mới (Dùng `C_mass_sop_factory`)
*Mục đích: Lấy danh sách sản phẩm mới từ Master File -> nhân bản 3 loại match-types -> Xuất file cho Amazon.*

1. **Chuẩn bị file:** Cập nhật sheet `Listing` trong file `PPC_NGUYÊN.xlsx` với các SKU mang trạng thái `Công việc mới` hoặc `Tiếp nhận`. Đặt file này vào thư mục `data/input/`.
2. **Tiền xử lý dữ liệu:** Chạy lệnh sau để nhân bản cấu trúc và lọc từ khoá.
   ```powershell
   python C_mass_sop_factory\excel_new_camp.py
   ```
   *(Hệ thống sẽ tạo ra file trung gian `PPC_PROCESSED_OUTPUT.xlsx`)*
3. **Tạo Bulk File Upload:** Tiếp tục chạy lệnh chuyển đổi sang định dạng Bulk Amazon.
   ```powershell
   python C_mass_sop_factory\mass_sop_factory.py
   ```
   *(Kết quả: Thu được file `Bulk_*.xlsx` sẵn sàng để tải lên Amazon)*
4. **Cập nhật ngược Master File:** Đánh dấu là đã tạo xong.
   ```powershell
   python C_mass_sop_factory\sync_master_file.py
   ```

### Quy Trình 2: Theo Dõi Và Tối Ưu Hiệu Suất (Dùng `D_auto_bulk_updater`)
*Mục đích: Đọc file Bulk tải từ Amazon về, kết nối với Master file và xử lý tắt (Pause/Archive) các key/campaign chạy kém.*

1. **Chuẩn bị file:**
   - Để `PPC_NGUYÊN.xlsx` (bản đã update) vào thẻ `data/input 1/`.
   - Lên Amazon Seller Central tải các file Bulk Report theo ngày, đưa luôn vào thẻ `data/input 1/`.
2. **Chạy Trình Cập Nhật:**
   ```powershell
   python D_auto_bulk_updater\end_to_end_ppc_tracker.py
   ```
3. **Kết Quả Nhận Được:**
   - Hệ thống sẽ nối thêm cột thống kê (Impressions, Clicks, Orders) cho từng SKU vào Master File (`PPC_NGUYÊN_UPDATED.xlsx`).
   - Tự sinh ra file Excel tên `Amazon_Bulk_Update_*.xlsx` chỉ chứa các từ khóa cần "Pause" / "Archive" (theo logic trong thẻ Rules). Bạn đem file này Upload lên Amazon.

---

> [!NOTE]
> - Nếu bạn muốn sửa logic (vd: bao nhiêu click mới Pause, bao nhiêu hiển thị mới cần sửa Key) -> Hãy chỉnh sửa trong file JSON `Rule_Engine.json` ở folder `F_core_engine`, không cần phải vào code Python để chỉnh sửa!
> - Thỉnh thoảng chạy các script ở `E_debug_tools` để đảm bảo dữ liệu Master không bị lỗi cấu trúc cột trước khi upload Amazon.
