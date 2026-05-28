# D_setup_name (SKU-Prefix Variant) — Hướng Dẫn Sử Dụng Chi Tiết

> **Mục đích:** Tự động đổi tên chiến dịch Amazon Sponsored Products theo SOP chuẩn, **dùng SKU trực tiếp làm prefix tên Campaign** (thay vì ASIN như ở folder G gốc). Tạo file Bulksheet upload-ready để import thẳng lên Amazon Ads.

> ⚠️ **Điểm khác biệt duy nhất so với `G_reset_system/D_setup_name`:**
> Tên campaign dùng **SKU** làm prefix thay vì **ASIN**.
> Ví dụ: `MY-SKU-123_SP02_AUTO_Nguyen` thay vì `B0XXXXXXXXX_SP02_AUTO_Nguyen`

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
├── module2_naming.py                   ← Bước 2: Đặt tên theo SOP (dùng SKU prefix)
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

### Bước 2 — Đặt Tên Theo SOP (SKU Prefix)

**File:** `module2_naming.py`

Áp dụng quy tắc đặt tên chuẩn cho từng chiến dịch, **dùng SKU làm prefix** thay vì ASIN.

#### 📐 Quy Tắc Đặt Tên SOP

| Loại chiến dịch | Pattern tên mới |
|---|---|
| AUTO Campaign | `{SKU}_SP02_AUTO_Nguyen` |
| Keyword (có KT suffix trong tên cũ) | `{SKU}_SP00_KT{suffix}_Nguyen` |
| Keyword (fallback từ data) | `{SKU}_SP00_{keyword}_{MATCH_TYPE}_Nguyen` |
| Product Targeting (có PT suffix) | `{SKU}_SP03_PT{suffix}_Nguyen` |
| Product Targeting (ASIN đối thủ) | `{SKU}_SP03_PT_{CompetitorASIN}_Nguyen` |
| Product Targeting (fallback) | `{SKU}_SP03_PT_{expression}_Nguyen` |

> **Lưu ý:** `{SKU}` là SKU gốc được sanitise (ký tự đặc biệt → `_`, viết HOA).
> Ví dụ: SKU `my-sku/123` → prefix `MY_SKU_123`

#### 🔁 De-Duplication (Tên Bị Trùng)
Nếu 2 chiến dịch sinh ra cùng tên → tự động thêm version:
```
{SKU}_SP00_keyword_EXACT_Nguyen      ← Campaign đầu tiên
{SKU}_SP00_keyword_EXACT_v2_Nguyen   ← Campaign thứ hai (tên bị trùng)
{SKU}_SP00_keyword_EXACT_v3_Nguyen   ← Campaign thứ ba...
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

**Output:**
- `data/master/Master_Amazon_Bulksheet_Update_Upload.xlsx`
- `data/master/Master_Amazon_Bulksheet_Update_Upload.csv` (encoding: utf-8-sig)

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

## ⚠️ Lưu Ý Quan Trọng

### Về SKU Prefix
- SKU được **viết HOA** và các ký tự đặc biệt (dấu `/`, `.`, khoảng trắng...) được thay bằng `_`
- Nhiều `_` liên tiếp sẽ được gộp thành 1 `_`
- Ví dụ: `my-product/v2.0` → `MY_PRODUCT_V2_0`

### Về File Input
- File phải có sheet tên **"Sponsored Products Campaigns"**
- File phải được export trực tiếp từ Amazon Ads → **không chỉnh sửa thủ công trước khi chạy**

### Về Chiến Dịch Phức Tạp
- Chiến dịch có **nhiều hơn 1 SKU** → được cách ly, **KHÔNG đổi tên tự động**
- Kiểm tra file `Complex_Campaigns_Review.xlsx` và xử lý thủ công

### Về Portfolio ID
- **Không được để trống Portfolio ID** khi upload → pipeline tự động giữ nguyên giá trị gốc

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

## 📤 Quy Trình Upload Lên Amazon

1. Chạy pipeline → lấy file **`Master_Amazon_Bulksheet_Update_Upload.xlsx`**
2. Vào **Amazon Ads Console → Bulk Operations → Upload**
3. Chọn file và upload
4. Kiểm tra kết quả sau 5-15 phút trong tab **History**

> **Lưu ý:** Nếu Amazon báo lỗi validation → kiểm tra file `Complex_Campaigns_Review.xlsx` xem các chiến dịch đó có vấn đề gì không.
