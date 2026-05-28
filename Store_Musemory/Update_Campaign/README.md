# D_setup_name — Hướng Dẫn Sử Dụng Chi Tiết

> **Mục đích:** Tự động đổi tên chiến dịch Amazon Sponsored Products theo SOP chuẩn, tạo file Bulksheet upload-ready để import thẳng lên Amazon Ads.

---

## 📁 Cấu Trúc Thư Mục

```
D_setup_name/
│
├── data/
│   ├── input/                          ← ĐẶT FILE AMAZON VÀO ĐÂY
│   │   └── BulkSheetExport_*.xlsx      ← File export từ Amazon Ads Console
│   │
│   ├── sku_files/                      ← [Tự động tạo] Output Bước 1
│   │   ├── SKU_<sku>.xlsx              ← Toàn bộ row của từng SKU
│   │   └── Complex_Campaigns_Review.xlsx ← Chiến dịch phức tạp cần review tay
│   │
│   ├── upload_files/                   ← [Tự động tạo] Output Bước 3
│   │   └── Update_CampaignName_<sku>.xlsx ← File update per-SKU
│   │
│   └── master/                         ← [Tự động tạo] Output Bước 4
│       ├── Master_Amazon_Bulksheet_Update_Upload.xlsx ← File upload cuối cùng
│       └── Master_Amazon_Bulksheet_Update_Upload.csv  ← Bản CSV (utf-8-sig)
│
├── output_files/                       ← Thư mục output phụ (kết quả merged)
│   ├── merged_amazon_bulksheet_upload.csv
│   └── merged_amazon_bulksheet_upload.xlsx
│
├── run_pipeline.py                     ← ĐIỂM CHẠY CHÍNH (all-in-one)
├── module1_isolation.py                ← Bước 1: Phân tách SKU
├── module2_naming.py                   ← Bước 2: Đặt tên theo SOP
├── module3_upload_generator.py         ← Bước 3: Tạo file update
├── merge_upload_files.py               ← Bước 4: Gộp master file
└── utils.py                            ← Hàm dùng chung
```

---

## 🚀 Cách Sử Dụng Nhanh (Khuyến Nghị)

### Bước chuẩn bị
1. Export BulkSheet từ Amazon Ads Console (chọn **Sponsored Products Campaigns**)
2. Đặt file `.xlsx` vào thư mục `data/input/`
3. Mở terminal tại thư mục `D_setup_name/`

### Chạy toàn bộ pipeline một lệnh
```powershell
python run_pipeline.py
```
Pipeline sẽ **tự động tìm** file `BulkSheetExport_*.xlsx` mới nhất trong `data/input/`.

---

## ⚙️ Các Tham Số CLI

```powershell
# Chỉ định file BulkSheet cụ thể
python run_pipeline.py --bulk data/input/BulkSheetExport_1805-2105.xlsx

# Chỉ định Amazon Template (để giữ nguyên format chuẩn)
python run_pipeline.py --template path/to/AmazonAdvertisingBulksheetSellerTemplate.xlsx

# Bỏ qua việc tạo file CSV
python run_pipeline.py --no-csv

# Chỉ chạy Bước 1-3, không gộp master file
python run_pipeline.py --skip-merge

# Kết hợp nhiều tham số
python run_pipeline.py --bulk data/input/file.xlsx --no-csv --skip-merge
```

| Tham số | Mặc định | Mô tả |
|---|---|---|
| `--bulk` | Tự động tìm | Đường dẫn file BulkSheetExport_*.xlsx |
| `--template` | Tự động tìm | Đường dẫn Amazon template .xlsx |
| `--no-csv` | Tắt | Không tạo file .csv ở Bước 4 |
| `--skip-merge` | Tắt | Bỏ qua Bước 4 (chỉ tạo file per-SKU) |

---

## 🔄 Chi Tiết 4 Bước Pipeline

### Bước 0 — Đọc & Làm Sạch BulkSheet
- Mở file Excel, ưu tiên sheet **"Sponsored Products Campaigns"**
- Ép tất cả cột về `dtype=str` để tránh lỗi scientific notation trên Campaign ID lớn
- Làm sạch các ID column (loại bỏ `.0`, `e+17`, v.v.)
- Xóa dòng trống hoàn toàn

---

### Bước 1 — Phân Tách SKU & Phát Hiện Chiến Dịch Phức Tạp

**File:** `module1_isolation.py`

Phân loại từng chiến dịch thành 2 nhóm:

| Loại | Điều kiện | Xử lý |
|---|---|---|
| **Standard** | 1 SKU duy nhất + chỉ 1 loại targeting | Ghi ra `SKU_<sku>.xlsx` |
| **Complex** | Nhiều SKU HOẶC có cả Keyword + Product Targeting | Ghi vào `Complex_Campaigns_Review.xlsx` để review tay |

**Output:**
- `data/sku_files/SKU_<sku>.xlsx` — 1 file per SKU, chứa toàn bộ row của SKU đó
- `data/sku_files/Complex_Campaigns_Review.xlsx` — chiến dịch cần review thủ công

**Chạy độc lập:**
```powershell
python module1_isolation.py
```

---

### Bước 2 — Đặt Tên Theo SOP

**File:** `module2_naming.py`

Áp dụng quy tắc đặt tên chuẩn cho từng chiến dịch.

#### 📐 Quy Tắc Đặt Tên SOP

| Loại chiến dịch | Pattern tên mới |
|---|---|
| AUTO Campaign | `{ASIN}_SP02_AUTO_Nguyen` |
| Keyword (có KT suffix trong tên cũ) | `{ASIN}_SP00_KT{suffix}_Nguyen` |
| Keyword (fallback từ data) | `{ASIN}_SP00_{keyword}_{MATCH_TYPE}_Nguyen` |
| Product Targeting (có PT suffix) | `{ASIN}_SP03_PT{suffix}_Nguyen` |
| Product Targeting (ASIN đối thủ) | `{ASIN}_SP03_PT_{CompetitorASIN}_Nguyen` |
| Product Targeting (fallback) | `{ASIN}_SP03_PT_{expression}_Nguyen` |

#### 🔍 Thứ Tự Resolve ASIN
1. Cột `ASIN (Informational only)` hoặc `ASIN` trên dòng **Product Ad**
2. Regex trích xuất từ chuỗi **SKU**
3. Regex trích xuất từ **Campaign Name** cũ
4. SKIP nếu không tìm được ASIN → cảnh báo

#### 🔁 De-Duplication (Tên Bị Trùng)
Nếu 2 chiến dịch sinh ra cùng tên → tự động thêm version:
```
{ASIN}_SP00_keyword_EXACT_Nguyen      ← Campaign đầu tiên
{ASIN}_SP00_keyword_EXACT_v2_Nguyen   ← Campaign thứ hai (tên bị trùng)
{ASIN}_SP00_keyword_EXACT_v3_Nguyen   ← Campaign thứ ba...
```

**Chạy độc lập (test 1 file SKU):**
```powershell
python module2_naming.py data/sku_files/SKU_B0XXXXXXXXX.xlsx
```

---

### Bước 3 — Tạo File Update Per-SKU

**File:** `module3_upload_generator.py`

Tạo file Bulksheet update theo format chuẩn Amazon (30 cột chính xác).

**Các quy tắc quan trọng:**
- Chỉ ghi dòng có `Entity = Campaign` (không ghi Ad Group, Keyword, v.v.)
- **Campaign ID** và **Portfolio ID** phải giữ nguyên ID gốc (không để mất liên kết portfolio)
- Các cột khác để trống → Amazon sẽ không thay đổi những trường đó

**Output:** `data/upload_files/Update_CampaignName_<sku>.xlsx`

**Template Strategy:**
- Nếu có Amazon template `.xlsx` → ghi vào bản copy của template (giữ nguyên hidden sheets, format)
- Nếu không có template → tạo plain XLSX (vẫn upload được)

**Chạy độc lập:**
```powershell
python module3_upload_generator.py data/sku_files/SKU_B0XXXXXXXXX.xlsx
# Với template:
python module3_upload_generator.py data/sku_files/SKU_B0XXXXXXXXX.xlsx path/to/template.xlsx
```

---

### Bước 4 — Gộp Master Upload File

**File:** `merge_upload_files.py`

Gộp tất cả file `Update_CampaignName_*.xlsx` trong `data/upload_files/` thành 1 file master.

**Các bảo đảm về data integrity:**
- Tất cả ID column được clean lần 2 (loại bỏ `.0`, `e+` notation)
- Chỉ giữ dòng `Entity = Campaign`
- Sắp xếp cột theo đúng thứ tự Amazon template
- Không có giá trị `nan` trong file output

**Output:**
- `data/master/Master_Amazon_Bulksheet_Update_Upload.xlsx`
- `data/master/Master_Amazon_Bulksheet_Update_Upload.csv` (encoding: utf-8-sig, mở được bằng Excel Windows)

**Chạy độc lập:**
```powershell
python merge_upload_files.py
python merge_upload_files.py --input data/upload_files/ --output data/final/
python merge_upload_files.py --no-csv
```

---

## 📋 30 Cột Amazon Template (Thứ Tự Chính Xác)

```
Product | Entity | Operation | Campaign ID | Ad Group ID | Portfolio ID |
Ad ID | Keyword ID | Product Targeting ID | Campaign Name | Ad Group Name |
Start Date | End Date | Targeting Type | State | Daily Budget | SKU |
Ad Group Default Bid | Bid | Keyword Text | Native Language Keyword |
Native Language Locale | Match Type | Bidding Strategy | Placement |
Percentage | Product Targeting Expression | Audience ID |
Shopper Cohort Percentage | Shopper Cohort Type
```

**Các giá trị tĩnh được điền tự động:**
| Cột | Giá trị |
|---|---|
| Product | `Sponsored Products` |
| Entity | `Campaign` |
| Operation | `Update` |
| State | `enabled` |

---

## 🛠️ Hàm Tiện Ích (utils.py)

| Hàm | Mô tả |
|---|---|
| `load_bulksheet(path)` | Đọc file Excel Amazon, ép dtype=str, làm sạch ID |
| `find_latest_bulksheet(base_dir)` | Tự động tìm file BulkSheetExport mới nhất trong `data/input/` |
| `clean_id(val)` | Loại bỏ `.0` và scientific notation khỏi Campaign ID |
| `safe_str(val)` | Chuyển bất kỳ giá trị về string, trả về `""` nếu None/NaN |
| `extract_asin(text)` | Trích xuất ASIN (pattern `B[A-Z0-9]{9}`) từ chuỗi bất kỳ |
| `clean_keyword(kw)` | Làm sạch keyword: chỉ giữ chữ, số, space, dấu gạch ngang |
| `no_double_underscore(name)` | Loại bỏ ký tự `__` thừa trong tên chiến dịch |
| `safe_filename(sku)` | Chuyển SKU thành tên file an toàn (thay ký tự đặc biệt bằng `_`) |

---

## ⚠️ Lưu Ý Quan Trọng

### Về File Input
- File phải có sheet tên **"Sponsored Products Campaigns"** (hoặc chứa cụm từ "sponsored product")
- Nếu không có sheet đúng tên → pipeline sẽ dùng sheet đầu tiên và cảnh báo
- File phải được export trực tiếp từ Amazon Ads → **không chỉnh sửa thủ công trước khi chạy**

### Về Chiến Dịch Phức Tạp
- Chiến dịch có **nhiều hơn 1 SKU** → được cách ly, **KHÔNG đổi tên tự động**
- Chiến dịch có **cả Keyword lẫn Product Targeting** → được cách ly
- Kiểm tra file `Complex_Campaigns_Review.xlsx` và xử lý thủ công

### Về Portfolio ID
- **Không được để trống Portfolio ID** khi upload → pipeline tự động giữ nguyên giá trị gốc
- Nếu để trống → chiến dịch sẽ bị **tách khỏi portfolio** trên Amazon

### Về ASIN Không Tìm Được
- Nếu pipeline không resolve được ASIN cho SKU → toàn bộ chiến dịch của SKU đó bị **bỏ qua**
- Kiểm tra: SKU có chứa ASIN không? Cột `ASIN` trong BulkSheet có dữ liệu không?

---

## 🔧 Yêu Cầu Python

```powershell
pip install pandas openpyxl
```

| Thư viện | Phiên bản tối thiểu | Mục đích |
|---|---|---|
| `pandas` | ≥ 1.5 | Đọc/ghi Excel, xử lý DataFrame |
| `openpyxl` | ≥ 3.0 | Engine đọc/ghi file .xlsx |
| `Python` | ≥ 3.10 | Cú pháp `str \| None` type hints |

---

## 📊 Ví Dụ Output Console

```
╔════════════════════════════════════════════════════════════════════════╗
║     AMAZON ADS – MASTER PIPELINE RUNNER (ALL-IN-ONE)                  ║
╠════════════════════════════════════════════════════════════════════════╣
║  Bulk file  : BulkSheetExport_1805-2105.xlsx                          ║
║  Template   : AmazonAdvertisingBulksheetSellerTemplate (3).xlsx       ║
║  Bắt đầu   : 2026-05-23 15:00:00                                      ║
╚════════════════════════════════════════════════════════════════════════╝

┌──────────────────────────────────────────────────────────────────────┐
│  BƯỚC 1 – Phân Tách SKU & Phát Hiện Chiến Dịch Phức Tạp             │
└──────────────────────────────────────────────────────────────────────┘
[MODULE 1] Total unique Campaign IDs found: 42
[MODULE 1] 38 standard SKU file(s) written, 4 complex campaign(s) quarantined.

╔════════════════════════════════════════════════════════════════════════╗
║  PIPELINE HOÀN TẤT – TỔNG KẾT                                         ║
╠════════════════════════════════════════════════════════════════════════╣
║  Tổng SKU tiêu chuẩn            : 12                                  ║
║  Chiến dịch phức tạp (cần review): 4                                  ║
║  Chiến dịch đã đặt tên (Bước 2) : 38                                  ║
║  File per-SKU đã tạo (Bước 3)   : 12                                  ║
║  Master Upload file (Bước 4)    : Master_Amazon_Bulksheet_Update_...  ║
╚════════════════════════════════════════════════════════════════════════╝
```

---

## 📤 Quy Trình Upload Lên Amazon

1. Chạy pipeline → lấy file **`Master_Amazon_Bulksheet_Update_Upload.xlsx`**
2. Vào **Amazon Ads Console → Bulk Operations → Upload**
3. Chọn file và upload
4. Kiểm tra kết quả sau 5-15 phút trong tab **History**

> **Lưu ý:** Nếu Amazon báo lỗi validation → kiểm tra file `Complex_Campaigns_Review.xlsx` xem các chiến dịch đó có vấn đề gì không.
