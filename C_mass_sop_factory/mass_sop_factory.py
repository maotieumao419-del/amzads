# =============================================================================
# FILE: mass_sop_factory.py  (V5 — Self-contained, C_mass_sop_factory)
#
# NHIỆM VỤ: Đọc PPC_PROCESSED_OUTPUT.xlsx từ "data/input 2/"
#           → xuất Amazon Bulk Operations format vào "data/output/"
#           Xuất 1 file Excel riêng mỗi SKU
#
# CẤU TRÚC INPUT (data/input 2/PPC_PROCESSED_OUTPUT.xlsx):
#   Mỗi sheet = 1 SKU, mỗi dòng = 1 keyword MỚI (đã lọc Active từ step 1)
#   Cột: STT | Campaign Name | Loại Campaign | Target | Ghi chú | ...
#
# QUY TRÌNH:
#  B1: Chạy excel_new_camp.py → tạo data/input 2/PPC_PROCESSED_OUTPUT.xlsx
#  B2: Chạy mass_sop_factory.py → xuất data/output/Bulk_Create_[SKU].xlsx
# =============================================================================

import os
import re
import glob
import pandas as pd
from datetime import datetime
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUT_1_DIR = os.path.join(SCRIPT_DIR, 'data', 'input 1')   # file PPC gốc (lấy Portfolio)
INPUT_2_DIR = os.path.join(SCRIPT_DIR, 'data', 'input 2')   # processed output
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, 'data', 'output')

PROCESSED_FILE = os.path.join(INPUT_2_DIR, 'PPC_PROCESSED_OUTPUT.xlsx')

# ---------------------------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------------------------
DEFAULT_BUDGET  = 5.0
DEFAULT_BID     = 0.50
DATE_SUFFIX     = datetime.now().strftime("%Y%m%d")
LISTING_SHEET   = "Listing"
COL_SKU_LISTING = "SKU"
COL_PID_LISTING = "Portfolio Id"

# ---------------------------------------------------------------------------
# AMAZON BULK TEMPLATE COLUMNS
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
# UTILITY
# =============================================================================

def find_col(df: pd.DataFrame, col_name: str) -> str | None:
    for c in df.columns:
        if c.strip().lower() == col_name.strip().lower():
            return c
    return None


# =============================================================================
# NOTE PARSER — nhất quán với excel_new_camp.py
# =============================================================================

def parse_note(note: str) -> dict:
    """
    Phân tích Ghi chú → dict với match_type, bid, placements.

    Hỗ trợ đầy đủ:
      "exact"                  → match=exact,  bid=0.50, T=R=P=0
      "exact 0.3_35TPR"        → match=exact,  bid=0.30, T=R=P=35
      "Phrase 0.45, 30TRP"     → match=phrase, bid=0.45, T=R=P=30
      "exact, 50T, 0.7Đ"       → match=exact,  bid=0.70, T=50, R=0, P=0
      "Phrase, 20TRP, 0.5Đ"    → match=phrase, bid=0.50, T=20, R=20, P=20
      "exact, 50RP, 0.68Đ"     → match=exact,  bid=0.68, T=0,  R=50, P=50
      "0.6 20T"                → match=None,   bid=0.60, T=20, R=0,  P=0
      "0.45"                   → match=None,   bid=0.45, T=R=P=0

    LOGIC PLACEMENT:
      Tìm pattern (số)([trp]+) trong chuỗi (sau khi normalize _ → space).
      Tách số thành int. Tách ký tự: T→Top, R→Rest, P→Product Page.
    """
    note_lower = note.lower()

    # ── 1. Match type ─────────────────────────────────────────────────────────
    # Hỗ trợ: "exact_35TPR", "exact, 50T", "Phrase 0.45"
    m_type = re.search(r'(?<![a-z])(exact|phrase|broad)(?![a-z])', note_lower)
    match_type = m_type.group(1) if m_type else None

    # ── 2. Placements ─────────────────────────────────────────────────────────
    placements = {"T": 0, "R": 0, "P": 0}
    p_matches = re.findall(r'(\d+)\s*[_,]?\s*([trp]{1,3})\b', note_lower)
    for num_str, keys_str in p_matches:
        val = int(num_str)
        if 't' in keys_str: placements["T"] = val
        if 'r' in keys_str: placements["R"] = val
        if 'p' in keys_str: placements["P"] = val

    # ── 3. Bid ────────────────────────────────────────────────────────────────
    bid = DEFAULT_BID
    bid_new = re.search(r'(\d+(?:\.\d+)?)\s*[đĐ]', note)
    if bid_new:
        bid = float(bid_new.group(1))
    else:
        note_norm = note.replace('_', ' ')
        floats = [float(x) for x in re.findall(r'\d+\.\d+', note_norm)]
        for f in floats:
            if f < 10:
                bid = f
                break
        else:
            b = re.search(r'(?i)bid\s*(\d+(?:\.\d+)?)', note)
            if b:
                bid = float(b.group(1))

    return {
        "match_type": match_type,
        "bid":        round(float(bid), 4),
        "placements": placements,
    }


# =============================================================================
# BƯỚC 1: {SKU → Portfolio_Id} từ PPC_*.xlsx gốc trong input 1
# =============================================================================

def build_sku_portfolio_map(master_file: str) -> dict[str, str]:
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
        print(f"  [ERROR] Thiếu cột SKU/Portfolio Id. Cột hiện có: {list(df.columns)}")
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
# BƯỚC 2: Đọc tất cả sheet từ PPC_PROCESSED_OUTPUT.xlsx (input 2)
# =============================================================================

def load_processed_sheets(processed_file: str) -> dict[str, pd.DataFrame]:
    print(f"\n[INPUT] Đọc processed data từ: {processed_file}")
    try:
        xls = pd.ExcelFile(processed_file)
    except Exception as e:
        print(f"  [ERROR] Không mở được file: {e}")
        return {}

    sheets_data = {}
    for sn in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sn, dtype=str)
        df.columns = df.columns.str.strip()
        df.fillna("", inplace=True)
        sheets_data[sn] = df
        print(f"  [OK] Sheet '{sn}': {len(df)} dòng")

    print(f"  → {len(sheets_data)} SKU sheet được nạp.")
    return sheets_data


# =============================================================================
# CORE: build_7_row_block — 7 dòng Entity chuẩn Amazon Bulk
# =============================================================================

def build_7_row_block(
    campaign_name : str,
    target_sku    : str,
    keyword_text  : str,
    match_type    : str,
    base_bid      : float,
    placements    : dict,
    portfolio_id  : str,
    default_budget: float = DEFAULT_BUDGET,
    date_suffix   : str   = DATE_SUFFIX,
    campaign_type : str   = "KT",
) -> list[dict]:
    """
    Tạo block 7 dòng Entity chuẩn Amazon Bulk Operations:
      r1: Campaign
      r2: Bidding Adjustment — Placement Top          ← placements["T"]
      r3: Bidding Adjustment — Placement Product Page ← placements["P"]
      r4: Bidding Adjustment — Placement Rest          ← placements["R"]
      r5: Ad Group
      r6: Product Ad
      r7: Keyword / Product Targeting
    """
    def base_row() -> dict:
        r = {col: "" for col in AMAZON_TEMPLATE_COLUMNS}
        r["Product"]     = "Sponsored Products"
        r["Operation"]   = "Create"
        r["Campaign Id"] = campaign_name
        r["State"]       = "Enabled"
        return r

    r1 = base_row()
    r1["Entity"]           = "Campaign"
    r1["Campaign Name"]    = campaign_name
    r1["Start Date"]       = date_suffix
    r1["Portfolio Id"]     = portfolio_id
    r1["Targeting Type"]   = "Manual"
    r1["Daily Budget"]     = default_budget
    r1["Bidding Strategy"] = "Dynamic bids - down only"

    r2 = base_row()
    r2["Entity"]     = "Bidding Adjustment"
    r2["Placement"]  = "Placement Top"
    r2["Percentage"] = placements.get("T", 0)

    r3 = base_row()
    r3["Entity"]     = "Bidding Adjustment"
    r3["Placement"]  = "Placement Product Page"
    r3["Percentage"] = placements.get("P", 0)

    r4 = base_row()
    r4["Entity"]     = "Bidding Adjustment"
    r4["Placement"]  = "Placement Rest Of Search"
    r4["Percentage"] = placements.get("R", 0)

    r5 = base_row()
    r5["Entity"]               = "Ad Group"
    r5["Ad Group Id"]          = campaign_name
    r5["Ad Group Name"]        = keyword_text
    r5["Ad Group Default Bid"] = base_bid

    r6 = base_row()
    r6["Entity"]      = "Product Ad"
    r6["Ad Group Id"] = campaign_name
    r6["SKU"]         = target_sku

    r7 = base_row()
    r7["Ad Group Id"] = campaign_name
    r7["Bid"]         = base_bid

    is_pt = ("product targeting" in str(campaign_type).lower()
             or str(campaign_type).strip().upper() == "PT")
    if is_pt:
        r7["Entity"] = "Product Targeting"
        r7["Product Targeting Expression"] = keyword_text
    else:
        r7["Entity"]       = "Keyword"
        r7["Keyword Text"] = keyword_text
        r7["Match Type"]   = match_type

    return [r1, r2, r3, r4, r5, r6, r7]


# =============================================================================
# OUTPUT: Xuất 1 file Excel riêng mỗi SKU
# =============================================================================

def export_sku_file(target_sku: str, rows: list[dict], output_dir: str) -> None:
    safe_sku = re.sub(r'[^\w\-]', '_', str(target_sku).strip()) or "UNKNOWN_SKU"
    filename = f"Bulk_Create_{safe_sku}.xlsx"
    out_path = os.path.join(output_dir, filename)

    df_out = pd.DataFrame(rows, columns=AMAZON_TEMPLATE_COLUMNS).fillna("")

    while True:
        try:
            with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
                sn = 'Sponsored Products Campaigns'
                df_out.to_excel(writer, index=False, sheet_name=sn)
                workbook  = writer.book
                worksheet = writer.sheets[sn]
                text_fmt  = workbook.add_format({'num_format': '@'})
                for col_idx in range(len(AMAZON_TEMPLATE_COLUMNS)):
                    worksheet.set_column(col_idx, col_idx, 20, text_fmt)
            n_camp = len(rows) // 7
            print(f"  [EXPORTED] {filename}  ({n_camp} campaigns | {len(rows)} rows)")
            break
        except PermissionError:
            input(f"\n  [!] File '{out_path}' đang mở. Đóng rồi nhấn Enter...")


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def main():
    print("\n" + "=" * 65)
    print("  MASS SOP FACTORY V5 — Per-SKU Bulk Exporter")
    print("  Bước 2/2: Đọc PPC_PROCESSED_OUTPUT → xuất Bulk Amazon")
    print("=" * 65)

    # ── Tìm file PPC_*.xlsx gốc trong input 1 (để tra Portfolio ID) ──────────
    master_files = glob.glob(os.path.join(INPUT_1_DIR, "PPC_*.xlsx"))
    master_files = [f for f in master_files if not os.path.basename(f).startswith("~$")]
    if not master_files:
        print(f"\n[ERROR] Không tìm thấy PPC_*.xlsx trong: {INPUT_1_DIR}")
        return
    if len(master_files) > 1:
        master_files.sort(key=os.path.getmtime, reverse=True)
        print(f"  [WARNING] Nhiều file → chọn mới nhất: {os.path.basename(master_files[0])}")
    MASTER_FILE = master_files[0]

    # ── Kiểm tra processed file tồn tại ─────────────────────────────────────
    if not os.path.exists(PROCESSED_FILE):
        print(f"\n[ERROR] Không tìm thấy: {PROCESSED_FILE}")
        print("  → Chạy excel_new_camp.py trước để tạo file này.")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"\n  Master (input 1) : {MASTER_FILE}")
    print(f"  Processed (input 2): {PROCESSED_FILE}")
    print(f"  Output Dir       : {OUTPUT_DIR}")
    print(f"  Date Suffix      : {DATE_SUFFIX}")

    # ── BƯỚC 1: SKU → Portfolio ID ───────────────────────────────────────────
    sku_portfolio_map = build_sku_portfolio_map(MASTER_FILE)
    if not sku_portfolio_map:
        print("\n[ERROR] Không có Portfolio mapping. Dừng.")
        return

    # ── BƯỚC 2: Đọc processed sheets ─────────────────────────────────────────
    sheets_data = load_processed_sheets(PROCESSED_FILE)
    if not sheets_data:
        print("\n[ERROR] Không đọc được processed data.")
        return

    # ── BƯỚC 3: Xử lý từng sheet → 7-row blocks ─────────────────────────────
    sku_buckets:     dict[str, list[dict]] = {}
    total_campaigns  = 0
    total_skipped    = 0

    print("\n" + "=" * 65)
    print("  XỬ LÝ TỪNG SKU SHEET")
    print("=" * 65)

    for target_sku, df_sku in sheets_data.items():
        portfolio_id = sku_portfolio_map.get(target_sku, "")
        if not portfolio_id or portfolio_id.lower() == "nan":
            print(f"\n  [WARN] '{target_sku}' không có Portfolio ID → để trống")
            portfolio_id = ""

        print(f"\n  SKU: {target_sku}  |  Portfolio: {portfolio_id}  |  {len(df_sku)} dòng")

        col_camp = find_col(df_sku, "Campaign Name")
        col_kw   = find_col(df_sku, "Target")
        col_note = find_col(df_sku, "Ghi chú")
        col_type = find_col(df_sku, "Loại Campaign")
        col_ts   = find_col(df_sku, "Trạng thái")

        if not col_kw:
            print(f"  [ERROR] Thiếu cột 'Target'. Cột hiện có: {list(df_sku.columns)}")
            continue

        sku_ok   = 0
        sku_skip = 0

        for idx, row in df_sku.iterrows():
            keyword_text  = str(row[col_kw]).strip()
            campaign_name = str(row[col_camp]).strip() if col_camp else ""
            note_str      = str(row[col_note]).strip() if col_note else ""
            status_val    = str(row[col_ts]).strip().lower() if col_ts else ""

            # Chỉ xử lý các dòng có trạng thái "Chưa tạo"
            if status_val != "chưa tạo":
                sku_skip += 1
                continue

            # Skip dòng không có keyword
            if not keyword_text or keyword_text.lower() == "nan":
                sku_skip += 1
                continue

            # Skip dòng không có campaign name (lẽ ra đã được generate ở bước 1)
            if not campaign_name or campaign_name.lower() == "nan":
                print(f"    [SKIP] Dòng {idx+2}: Campaign Name rỗng — '{keyword_text[:30]}'")
                sku_skip += 1
                continue

            camp_type_str = str(row[col_type]).strip() if col_type else ""
            is_pt = ("product targeting" in camp_type_str.lower()
                     or camp_type_str.strip().upper() == "PT")

            # ── Parse Ghi chú ────────────────────────────────────────────────
            try:
                parsed = parse_note(note_str)
            except Exception as e:
                print(f"    [SKIP] Dòng {idx+2}: Lỗi parse note '{note_str}': {e}")
                sku_skip += 1
                continue

            match_type = parsed["match_type"]
            base_bid   = parsed["bid"]
            placements = parsed["placements"]

            # Fallback match type từ Campaign Name nếu note không có
            if not match_type and not is_pt:
                camp_lower = campaign_name.lower()
                if   '_ex_' in camp_lower or '_exact_' in camp_lower: match_type = 'exact'
                elif '_pr_' in camp_lower or '_phrase_' in camp_lower: match_type = 'phrase'
                elif '_br_' in camp_lower or '_broad_' in camp_lower:  match_type = 'broad'

            if not match_type and not is_pt:
                print(f"    [SKIP] Dòng {idx+2}: Không xác định Match Type — '{note_str}'")
                sku_skip += 1
                continue

            block = build_7_row_block(
                campaign_name  = campaign_name,
                target_sku     = target_sku,
                keyword_text   = keyword_text,
                match_type     = match_type or "",
                base_bid       = base_bid,
                placements     = placements,
                portfolio_id   = portfolio_id,
                date_suffix    = DATE_SUFFIX,
                campaign_type  = camp_type_str,
            )

            sku_buckets.setdefault(target_sku, []).extend(block)
            sku_ok += 1

        total_campaigns += sku_ok
        total_skipped   += sku_skip
        print(f"    → {sku_ok} campaigns OK | {sku_skip} dòng bỏ qua")

    # ── BƯỚC 4: Xuất file Bulk ────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  XUẤT FILE BULK")
    print("=" * 65)

    if not sku_buckets:
        print("  [ERROR] Không có dữ liệu hợp lệ để xuất.")
        return

    for sku, rows in sku_buckets.items():
        export_sku_file(sku, rows, OUTPUT_DIR)

    # ── SUMMARY ──────────────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  FINAL SUMMARY")
    print("=" * 65)
    print(f"  SKU sheets đầu vào     : {len(sheets_data)}")
    print(f"  Campaigns thành công   : {total_campaigns}")
    print(f"  Dòng bỏ qua (invalid)  : {total_skipped}")
    print(f"  Tổng rows được tạo     : {total_campaigns * 7}")
    print(f"  Files Excel đã xuất    : {len(sku_buckets)}")
    print(f"  Output directory       : {OUTPUT_DIR}")
    print("=" * 65 + "\n")


if __name__ == "__main__":
    main()
