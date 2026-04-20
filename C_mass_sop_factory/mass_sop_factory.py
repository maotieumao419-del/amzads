# =============================================================================
# FILE: mass_sop_factory.py  (V3 — Self-contained, no A/B folder dependency)
#
# NHIỆM VỤ: Map dữ liệu từ PPC_PROCESSED_OUTPUT.xlsx → Amazon Bulk Operations format
#           Xuất 1 file Excel riêng mỗi SKU (n sheets → n files)
#
# ── QUY TRÌNH SỬ DỤNG ───────────────────────────────────────────────────────────
#  B1: Đặt PPC_NGUYEN.xlsx vào:   C_mass_sop_factory/data/input/
#  B2: Chạy: python excel_new_camp.py
#             → tạo PPC_PROCESSED_OUTPUT.xlsx vào cùng data/input/
#  B3: Chạy: python mass_sop_factory.py
#             → xuất Bulk_Create_[SKU].xlsx vào data/output/
# =============================================================================

import os
import re
import pandas as pd
from datetime import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR   = os.path.join(SCRIPT_DIR, 'data', 'input')

# File 1: Output từ excel_new_camp.py (mỗi sheet = 1 SKU đã xử lý)
# Được tạo tự động bởi excel_new_camp.py trong cùng thư mục data/input/
PROCESSED_FILE = os.path.join(INPUT_DIR, 'PPC_PROCESSED_OUTPUT.xlsx')

# File 2: File kế hoạch gốc — chỉ dùng sheet "Listing" để tra Portfolio ID
# Đặt cùng thư mục: C_mass_sop_factory/data/input/PPC_NGUYEN.xlsx
MASTER_FILE    = os.path.join(INPUT_DIR, 'PPC_NGUYEN.xlsx')

# ── [OUTPUT MỚI] Thư mục chứa các file bulk riêng theo Portfolio ID ─────────
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, 'data', 'output')

# ---------------------------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------------------------
DEFAULT_BUDGET   = 5.0          # Ngân sách hàng ngày mặc định nếu không có trong dữ liệu
DATE_SUFFIX      = datetime.now().strftime("%Y%m%d")   # 20260416
LISTING_SHEET    = "Listing"
COL_SKU_LISTING  = "SKU"
COL_PID_LISTING  = "Portfolio Id"   # tên cột Portfolio ID trong sheet Listing

# ---------------------------------------------------------------------------
# AMAZON BULK TEMPLATE — GIỮ NGUYÊN TỪ V2
# ---------------------------------------------------------------------------
AMAZON_TEMPLATE_COLUMNS = [
    "Product", "Entity", "Operation", "Campaign Id", "Ad Group Id", "Portfolio Id",
    "Ad Id", "Keyword Id", "Product Targeting Id", "Campaign Name", "Ad Group Name",
    "Start Date", "End Date", "Targeting Type", "State", "Daily Budget", "SKU", "ASIN",
    "Eligibility Status", "Reasons for Ineligibility", "Ad Group Default Bid",
    "Bid", "Keyword Text", "Match Type", "Bidding Strategy", "Placement", "Percentage",
    "Product Targeting Expression"
]


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def find_col(df: pd.DataFrame, col_name: str) -> str | None:
    """Case-insensitive column lookup, returns None nếu không tìm thấy."""
    for c in df.columns:
        if c.strip().lower() == col_name.strip().lower():
            return c
    return None


# =============================================================================
# [INPUT MỚI] BƯỚC 1: Xây dựng dictionary {SKU → Portfolio_Id}
#   Nguồn: sheet "Listing" trong PPC_NGUYÊN.xlsx
# =============================================================================

def build_sku_portfolio_map(master_file: str) -> dict[str, str]:
    """
    Đọc sheet Listing từ file kế hoạch gốc.
    Trả về dict: { 'UY-AO7W-2FWB': '128613535409943', ... }
    """
    print(f"\n[INPUT] Đọc Portfolio mapping từ: {master_file}")

    try:
        df = pd.read_excel(master_file, sheet_name=LISTING_SHEET, dtype=str)
    except Exception as e:
        print(f"  [ERROR] Không đọc được sheet '{LISTING_SHEET}': {e}")
        return {}

    df.columns = df.columns.str.strip()
    df.fillna("", inplace=True)

    col_sku = find_col(df, COL_SKU_LISTING)
    col_pid = find_col(df, COL_PID_LISTING)

    if not col_sku or not col_pid:
        print(f"  [ERROR] Không tìm thấy cột '{COL_SKU_LISTING}' hoặc '{COL_PID_LISTING}'.")
        print(f"  Các cột hiện có: {list(df.columns)}")
        return {}

    mapping = {}
    for _, row in df.iterrows():
        sku = str(row[col_sku]).replace("\n", "").strip()
        pid = str(row[col_pid]).strip()
        if sku and sku.lower() != "nan":
            mapping[sku] = pid

    print(f"  → Đã load {len(mapping)} SKU → Portfolio ID mappings.")
    return mapping


# =============================================================================
# [INPUT MỚI] BƯỚC 2: Đọc toàn bộ sheet từ PPC_PROCESSED_OUTPUT.xlsx
#   Mỗi sheet name = tên SKU; mỗi dòng = 1 campaign keyword đã xử lý
# =============================================================================

def load_processed_sheets(processed_file: str) -> dict[str, pd.DataFrame]:
    """
    Đọc tất cả sheet trong PPC_PROCESSED_OUTPUT.xlsx.
    Trả về dict: { 'UY-AO7W-2FWB': DataFrame, ... }
    """
    print(f"\n[INPUT] Đọc dữ liệu processed từ: {processed_file}")

    try:
        xls = pd.ExcelFile(processed_file)
    except Exception as e:
        print(f"  [ERROR] Không mở được file: {e}")
        return {}

    sheets_data = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
        df.columns = df.columns.str.strip()
        df.fillna("", inplace=True)
        sheets_data[sheet_name] = df
        print(f"  [OK] Sheet '{sheet_name}': {len(df)} dòng")

    print(f"  → Tổng {len(sheets_data)} SKU sheet được nạp.")
    return sheets_data


# =============================================================================
# REGEX PARSERS — Bóc tách từ cột "Ghi chú"
# =============================================================================

def parse_match_type(note: str) -> str | None:
    """
    Tìm match type (exact / phrase / broad) trong chuỗi Ghi chú.

    REGEX GIẢI THÍCH:
      \\b(exact|phrase|broad)\\b
        \\b   → word boundary — không bắt "exactly" hay "broadly"
        (...)  → capture group chứa 1 trong 3 giá trị
    Trả về lowercase, hoặc None nếu không tìm thấy.
    """
    m = re.search(r'\b(exact|phrase|broad)\b', note, re.IGNORECASE)
    return m.group(1).lower() if m else None


def parse_base_bid(note: str) -> float:
    """
    Tìm giá thầu (bid) trong chuỗi Ghi chú.

    Chiến lược ưu tiên:
      1. Float đầu tiên tìm được (Ví dụ: "Phrase 0.55, 30TRP" → 0.55)
      2. Tìm pattern 'bid <số>' (Ví dụ: "bid 0.6")
      3. Số nguyên/thực nhỏ hơn 10 đầu tiên (phòng khi chỉ có "55 0.55")
      4. Mặc định 0.50

    REGEX GIẢI THÍCH (bước 1):
      \\d+\\.\\d+  → Số thực (ví dụ: 0.55, 1.20)
    REGEX GIẢI THÍCH (bước 2):
      (?i)bid\\s*(\\d+(?:\\.\\d+)?)
        (?i)         → case-insensitive
        bid\\s*      → từ "bid" + 0 hoặc nhiều khoảng trắng
        (\\d+(...))  → số nguyên hoặc số thực
    """
    # Bước 1: float đầu tiên
    floats = [float(x) for x in re.findall(r'\d+\.\d+', note)]
    if floats:
        return floats[0]

    # Bước 2: pattern "bid 0.6"
    b_match = re.search(r'(?i)bid\s*(\d+(?:\.\d+)?)', note)
    if b_match:
        return float(b_match.group(1))

    # Bước 3: số thực nhỏ hơn 10
    raw_nums = re.findall(r'\d+(?:\.\d+)?', note)
    for n in raw_nums:
        v = float(n)
        if v < 10:
            return v

    return 0.50  # Mặc định


def parse_placement_pct(note: str) -> int | None:
    """
    Tìm Placement Percentage trong chuỗi Ghi chú.

    REGEX GIẢI THÍCH:
      (\\d+)\\s*(?:TRP|TPR|T(?![A-Z])|%)
        (\\d+)          → con số cần bắt (Ví dụ: 30 trong "30TRP")
        \\s*            → cho phép có khoảng trắng giữa số và chữ
        (?:TRP|TPR|...) → non-capture group: các hậu tố chấp nhận được
        T(?![A-Z])      → chữ T đơn (không theo sau bởi chữ hoa khác)
    """
    m = re.search(r'(\d+)\s*(?:TRP|TPR|T(?![A-Z])|%)', note, re.IGNORECASE)
    return int(m.group(1)) if m else None


# =============================================================================
# CORE LOGIC: Tạo 7 dòng Entity cho mỗi Campaign — GIỮ NGUYÊN CẤU TRÚC V2
# =============================================================================

def build_7_row_block(
    campaign_name: str,
    target_sku: str,
    keyword_text: str,
    match_type: str,
    base_bid: float,
    placement_pct: int,
    portfolio_id: str,
    default_budget: float = DEFAULT_BUDGET,
    date_suffix: str = DATE_SUFFIX,
) -> list[dict]:
    """
    Tạo block 7 dòng Entity chuẩn Amazon Bulk Operations.
    Cấu trúc:
      r1: Campaign
      r2: Bidding Adjustment (placement top / TOS)
      r3: Bidding Adjustment (placementProductPage / PP)
      r4: Bidding Adjustment (placementRestOfSearch / ROS)
      r5: Ad Group
      r6: Product Ad
      r7: Keyword
    """
    ad_group_id_str   = campaign_name   # dùng Campaign Name làm Ad Group ID (nhất quán với V2)
    ad_group_name_str = keyword_text

    def base_row() -> dict:
        """Khởi tạo dòng rỗng với các trường cố định."""
        r = {col: "" for col in AMAZON_TEMPLATE_COLUMNS}
        r["Product"]     = "Sponsored Products"
        r["Operation"]   = "Create"
        r["Campaign Id"] = campaign_name
        r["State"]       = "enabled"
        return r

    # Row 1 — Campaign
    r1 = base_row()
    r1["Entity"]           = "Campaign"
    r1["Campaign Name"]    = campaign_name
    r1["Start Date"]       = date_suffix
    r1["Portfolio Id"]     = portfolio_id          # ← Portfolio ID được inject từ dict mapping
    r1["Targeting Type"]   = "MANUAL"
    r1["Daily Budget"]     = default_budget
    r1["Bidding Strategy"] = "Dynamic bids - down only"

    # Row 2 — Bidding Adjustment: Top of Search
    r2 = base_row()
    r2["Entity"]     = "Bidding Adjustment"
    r2["Placement"]  = "placement top"
    r2["Percentage"] = placement_pct

    # Row 3 — Bidding Adjustment: Product Page
    r3 = base_row()
    r3["Entity"]     = "Bidding Adjustment"
    r3["Placement"]  = "placementProductPage"
    r3["Percentage"] = placement_pct

    # Row 4 — Bidding Adjustment: Rest of Search
    r4 = base_row()
    r4["Entity"]     = "Bidding Adjustment"
    r4["Placement"]  = "placementRestOfSearch"
    r4["Percentage"] = placement_pct

    # Row 5 — Ad Group
    r5 = base_row()
    r5["Entity"]              = "Ad Group"
    r5["Ad Group Id"]         = ad_group_id_str
    r5["Ad Group Name"]       = ad_group_name_str
    r5["Ad Group Default Bid"] = base_bid

    # Row 6 — Product Ad
    r6 = base_row()
    r6["Entity"]     = "Product Ad"
    r6["Ad Group Id"] = ad_group_id_str
    r6["SKU"]        = target_sku

    # Row 7 — Keyword
    r7 = base_row()
    r7["Entity"]       = "Keyword"
    r7["Ad Group Id"]  = ad_group_id_str
    r7["Keyword Text"] = keyword_text
    r7["Match Type"]   = match_type
    r7["Bid"]          = base_bid

    return [r1, r2, r3, r4, r5, r6, r7]


# =============================================================================
# [OUTPUT] Xuất 1 file vật lý riêng biệt cho mỗi SKU sheet
#   Quy tắc: n sheet input → n file output
#   Tên file: Bulk_Create_[SKU].xlsx
#   Sheet name trong file: 'Sponsored Products Campaigns'
# =============================================================================

def export_sku_file(target_sku: str, rows: list[dict], output_dir: str) -> None:
    """
    Nhận toàn bộ các dòng của 1 SKU và xuất ra 1 file Excel riêng.
    1 sheet input  →  1 file output  (n sheets → n files).
    Áp dụng text-format (@) cho tất cả cột để tránh Excel tự convert kiểu dữ liệu.
    """
    # ── Sanitize tên file (bỏ ký tự đặc biệt không hợp lệ trên Windows) ────
    safe_sku  = re.sub(r'[^\w\-]', '_', str(target_sku).strip()) or "UNKNOWN_SKU"
    filename  = f"Bulk_Create_{safe_sku}.xlsx"
    out_path  = os.path.join(output_dir, filename)

    df_out = pd.DataFrame(rows, columns=AMAZON_TEMPLATE_COLUMNS).fillna("")

    while True:
        try:
            with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
                sheet_name = 'Sponsored Products Campaigns'
                df_out.to_excel(writer, index=False, sheet_name=sheet_name)

                workbook   = writer.book
                worksheet  = writer.sheets[sheet_name]
                text_fmt   = workbook.add_format({'num_format': '@'})
                for col_idx in range(len(AMAZON_TEMPLATE_COLUMNS)):
                    worksheet.set_column(col_idx, col_idx, 20, text_fmt)

            n_campaigns = len(rows) // 7
            print(f"  [EXPORTED] SKU '{target_sku}' → {filename}")
            print(f"             {n_campaigns} campaigns | {len(rows)} rows")
            break
        except PermissionError:
            input(
                f"\n  [!] File '{out_path}' đang mở.\n"
                f"      Đóng Excel rồi nhấn Enter để thử lại..."
            )


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def main():
    print("\n" + "=" * 65)
    print("  MASS SOP FACTORY V3 — Per-SKU Bulk Exporter")
    print("  Bước 2/2: Đọc processed data → xuất Bulk Amazon")
    print("=" * 65)

    # ── Kiểm tra file input tồn tại ────────────────────────────────────────
    for f, label in [(PROCESSED_FILE, 'PPC_PROCESSED_OUTPUT.xlsx'),
                      (MASTER_FILE,    'PPC_NGUYEN.xlsx')]:
        if not os.path.exists(f):
            print(f"\n[ERROR] Không tìm thấy: {f}")
            if label == 'PPC_PROCESSED_OUTPUT.xlsx':
                print("  → Chạy excel_new_camp.py trước để tạo file này.")
            else:
                print(f"  → Đặt {label} vào: {INPUT_DIR}")
            return

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"\n  Processed Input : {PROCESSED_FILE}")
    print(f"  Master Input    : {MASTER_FILE}")
    print(f"  Output Dir      : {OUTPUT_DIR}")
    print(f"  Date Suffix     : {DATE_SUFFIX}")

    # ─────────────────────────────────────────────────────────────────────────
    # [INPUT MỚI] BƯỚC 1: Xây dựng SKU → Portfolio ID mapping
    # ─────────────────────────────────────────────────────────────────────────
    sku_portfolio_map = build_sku_portfolio_map(MASTER_FILE)
    if not sku_portfolio_map:
        print("\n[ERROR] Không có SKU → Portfolio mapping. Dừng chương trình.")
        return

    # ─────────────────────────────────────────────────────────────────────────
    # [INPUT MỚI] BƯỚC 2: Đọc tất cả sheet từ PPC_PROCESSED_OUTPUT.xlsx
    # ─────────────────────────────────────────────────────────────────────────
    sheets_data = load_processed_sheets(PROCESSED_FILE)
    if not sheets_data:
        print("\n[ERROR] Không đọc được dữ liệu từ PPC_PROCESSED_OUTPUT.xlsx.")
        return

    # ─────────────────────────────────────────────────────────────────────────
    # BƯỚC 3: Xử lý từng SKU sheet → tạo 7-row blocks → bucket theo SKU
    #
    # [OUTPUT] Mỗi SKU = 1 bucket riêng = 1 file Excel riêng
    #          n sheet input  →  n file output (1-to-1 mapping)
    # ─────────────────────────────────────────────────────────────────────────

    # sku_buckets: { target_sku: [list of 7-row dicts] }
    sku_buckets: dict[str, list[dict]] = {}

    total_campaigns  = 0
    total_skipped    = 0

    print("\n" + "=" * 65)
    print("  XỬ LÝ TỪNG SKU SHEET")
    print("=" * 65)

    for target_sku, df_sku in sheets_data.items():
        # Tra Portfolio ID từ mapping — gán rỗng nếu SKU không có trong mapping
        portfolio_id = sku_portfolio_map.get(target_sku, "")
        if not portfolio_id or portfolio_id.lower() == "nan":
            print(f"\n  [WARN] SKU '{target_sku}' không có Portfolio ID → nhóm vào 'NO_PORTFOLIO'")
            portfolio_id = "NO_PORTFOLIO"

        print(f"\n  SKU: {target_sku}  |  Portfolio: {portfolio_id}  |  {len(df_sku)} dòng")

        # Tìm các cột cần thiết trong sheet processed
        col_camp = find_col(df_sku, "Campaign Name")
        col_kw   = find_col(df_sku, "Target")
        col_note = find_col(df_sku, "Ghi chú")

        if not col_camp or not col_kw or not col_note:
            print(f"  [ERROR] Thiếu cột bắt buộc trong sheet '{target_sku}'. Bỏ qua.")
            print(f"          Cột hiện có: {list(df_sku.columns)}")
            continue

        sku_campaign_count = 0
        sku_skip_count     = 0

        for idx, row in df_sku.iterrows():
            campaign_name = str(row[col_camp]).strip()
            keyword_text  = str(row[col_kw]).strip()
            note_str      = str(row[col_note]).strip()

            # ── Skip dòng không có keyword ──────────────────────────────────
            if not keyword_text or keyword_text.lower() == "nan":
                sku_skip_count += 1
                continue

            # ── Bóc tách từ cột Ghi chú bằng Regex ─────────────────────────
            match_type = parse_match_type(note_str)
            if not match_type:
                print(f"    [SKIP] Dòng {idx+2}: Không xác định được Match Type — '{note_str}'")
                sku_skip_count += 1
                continue

            base_bid = parse_base_bid(note_str)

            placement_pct = parse_placement_pct(note_str)
            if placement_pct is None:
                print(f"    [SKIP] Dòng {idx+2}: Không trích xuất được Placement % — '{note_str}'")
                sku_skip_count += 1
                continue

            # ── Tạo block 7 dòng (Campaign → Bidding Adj × 3 → Ad Group → Product Ad → Keyword) ──
            block = build_7_row_block(
                campaign_name = campaign_name,
                target_sku    = target_sku,
                keyword_text  = keyword_text,
                match_type    = match_type,
                base_bid      = base_bid,
                placement_pct = placement_pct,
                portfolio_id  = portfolio_id,
                date_suffix   = DATE_SUFFIX,
            )

            # ── [OUTPUT] Đẩy vào bucket theo SKU (1 SKU = 1 file) ─────────
            if target_sku not in sku_buckets:
                sku_buckets[target_sku] = []
            sku_buckets[target_sku].extend(block)

            sku_campaign_count += 1

        total_campaigns += sku_campaign_count
        total_skipped   += sku_skip_count
        print(f"    → {sku_campaign_count} campaigns OK | {sku_skip_count} dòng bỏ qua")

    # ─────────────────────────────────────────────────────────────────────────
    # [OUTPUT] BƯỚC 4: Xuất 1 file riêng mỗi SKU (n sheets → n files)
    # ─────────────────────────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  XUẤT FILE BULK THEO TỪNG SKU")
    print("=" * 65)

    if not sku_buckets:
        print("  [ERROR] Không có dữ liệu hợp lệ để xuất.")
        return

    for sku, rows in sku_buckets.items():
        export_sku_file(sku, rows, OUTPUT_DIR)

    # ─────────────────────────────────────────────────────────────────────────
    # SUMMARY
    # ─────────────────────────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  FINAL SUMMARY")
    print("=" * 65)
    print(f"  SKU sheets đầu vào     : {len(sheets_data)}")
    print(f"  Campaigns thành công   : {total_campaigns}")
    print(f"  Dòng bỏ qua (invalid)  : {total_skipped}")
    print(f"  Tổng rows được tạo     : {total_campaigns * 7}")
    print(f"  Files Excel đã xuất    : {len(sku_buckets)}  (= số SKU sheet)")
    print(f"  Output directory       : {OUTPUT_DIR}")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    main()
