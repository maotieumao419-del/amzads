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
import sys
import pandas as pd
from datetime import datetime
import warnings

# UTF-8 Console Reconfiguration for Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

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

def parse_note(note: str, campaign_name: str = "") -> dict:
    """
    Phân tích Ghi chú → dict với match_type, bid, placements.
    Hỗ trợ cấu trúc linh hoạt: exact_0.5_2T, exact 0.5, 2 T, Phrase 0.45, 30TRP...
    """
    note_lower = note.lower()

    # 1. PLACEMENTS
    placements = {"T": 0, "R": 0, "P": 0}
    placement_tags = []

    # Tìm các mẫu như "2T", "2 T", "35 TPR", "50_R_P"
    p_matches = list(re.finditer(r'(\d+(?:\.\d+)?)\s*[_,]?\s*([tpr][tpr\s_,]*)(\b|$)', note_lower))
    for match in p_matches:
        num_str = match.group(1)
        keys_str = match.group(2).replace(' ', '').replace('_', '').replace(',', '')
        
        # Đảm bảo phần chữ chỉ chứa t, r, p
        if set(keys_str).issubset(set('trp')):
            val = int(float(num_str))
            tag_keys = ""
            if 't' in keys_str: placements["T"] = val; tag_keys += "T"
            if 'r' in keys_str: placements["R"] = val; tag_keys += "R"
            if 'p' in keys_str: placements["P"] = val; tag_keys += "P"
            
            if tag_keys:
                placement_tags.append(f"{val}{tag_keys}")
            
            # Xóa phần placement khỏi note_lower để tránh nhận diện nhầm thành bid
            note_lower = note_lower.replace(match.group(0), ' ')

    # Format dạng [T: 0 / P: 0 / R: 0]
    p_new = list(re.finditer(r'([tpr])\s*:\s*(\d+)', note_lower))
    for match in p_new:
        k = match.group(1)
        v = int(match.group(2))
        if k == 't': placements["T"] = v
        if k == 'r': placements["R"] = v
        if k == 'p': placements["P"] = v
        note_lower = note_lower.replace(match.group(0), ' ')

    placement_tag = "_".join(dict.fromkeys(placement_tags)) if placement_tags else "00T"

    # 2. MATCH TYPE
    if "exact" in note_lower or "ex" in note_lower:
        match_type = "exact"
    elif "phrase" in note_lower or "ph" in note_lower:
        match_type = "phrase"
    elif "broad" in note_lower or "br" in note_lower:
        match_type = "broad"
    elif "targeting" in note_lower or "pt" in note_lower or "asin" in note_lower:
        match_type = "targeting"
    elif "auto" in note_lower or "au" in note_lower:
        match_type = "auto"
    else:
        match_type = "unknown"

    # 3. BID
    bid = DEFAULT_BID
    bid_found = False
    
    # Tìm tất cả các số còn lại trong chuỗi (sau khi đã xóa placement)
    floats = re.findall(r'(\d+(?:\.\d+)?)', note_lower)
    for f_str in floats:
        f = float(f_str)
        # Thông thường bid < 30 và > 0
        if 0 < f < 30:
            bid = f
            bid_found = True
            break

    # Fallback to campaign_name if no explicit bid found in note
    if not bid_found and campaign_name:
        # Tìm số thập phân trong Campaign Name, ví dụ: _0.33 hoặc _0.3_070825
        camp_floats = [float(x) for x in re.findall(r'_(\d+\.\d+)(?:_|$)', campaign_name)]
        if camp_floats:
            for f in reversed(camp_floats): # Lấy số cuối cùng thỏa mãn < 10
                if f < 10:
                    bid = f
                    break

    return {
        "match_type": match_type,
        "bid":        round(float(bid), 4),
        "placements": placements,
        "placement_tag": placement_tag
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
        # Read without header to find the true header row dynamically
        df = pd.read_excel(xls, sheet_name=sn, dtype=str, header=None)
        
        header_idx = 0
        for idx, row in df.iterrows():
            row_vals = [str(x).strip().lower() for x in row.values if pd.notna(x)]
            if "target" in row_vals or "campaign name" in row_vals or "ghi chú" in row_vals or "note" in row_vals:
                header_idx = idx
                break
                
        # Set the columns to the header row found
        df.columns = df.iloc[header_idx].astype(str).str.strip()
        # Keep only the rows after the header
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
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
    raw_keyword   : str,
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

    is_pt = ("product targeting" in str(campaign_type).lower()
             or str(campaign_type).strip().upper() == "PT")

    rows = [r1, r2, r3, r4]

    import re
    clean_keyword = re.sub(r'[^\w\s-]', '', str(raw_keyword)).strip()

    if str(match_type).strip().lower() == 'exact':
        num_clones = 3
    else:
        num_clones = 1

    for i in range(1, num_clones + 1):
        if num_clones > 1:
            ad_group_name = f"AG_{i:02d}_Exact_{clean_keyword}"
            ad_group_id = ad_group_name
        else:
            ad_group_name = raw_keyword
            ad_group_id = campaign_name  # Preserve original logic for non-clones

        r5 = base_row()
        r5["Entity"]               = "Ad Group"
        r5["Ad Group Id"]          = ad_group_id
        r5["Ad Group Name"]        = ad_group_name
        r5["Ad Group Default Bid"] = base_bid

        r6 = base_row()
        r6["Entity"]      = "Product Ad"
        r6["Ad Group Id"] = ad_group_id
        r6["SKU"]         = target_sku

        r7 = base_row()
        r7["Ad Group Id"] = ad_group_id
        r7["Bid"]         = base_bid

        if is_pt:
            r7["Entity"] = "Product Targeting"
            r7["Product Targeting Expression"] = raw_keyword
        else:
            r7["Entity"]       = "Keyword"
            r7["Keyword Text"] = raw_keyword
            r7["Match Type"]   = match_type

        rows.extend([r5, r6, r7])

    if str(match_type).strip().lower() == 'exact':
        assert len(rows) == 13, f"Lỗi: Không tạo đủ 13 dòng cho Exact Matrix (hiện tại: {len(rows)})"

    return rows


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


def export_merged_file(store_name: str, task_name: str, all_rows: list[dict], output_dir: str) -> str:
    from datetime import datetime
    date_str = datetime.now().strftime("%d%m%y")
    filename_xlsx = f"{store_name}_{task_name}_{date_str}.xlsx"
    filename_csv = f"{store_name}_{task_name}_{date_str}.csv"
    out_path_xlsx = os.path.join(output_dir, filename_xlsx)
    out_path_csv = os.path.join(output_dir, filename_csv)

    df_out = pd.DataFrame(all_rows, columns=AMAZON_TEMPLATE_COLUMNS).fillna("")

    while True:
        try:
            # Save Excel
            with pd.ExcelWriter(out_path_xlsx, engine='xlsxwriter') as writer:
                sn = 'Sponsored Products Campaigns'
                df_out.to_excel(writer, index=False, sheet_name=sn)
                workbook  = writer.book
                worksheet = writer.sheets[sn]
                text_fmt  = workbook.add_format({'num_format': '@'})
                for col_idx in range(len(AMAZON_TEMPLATE_COLUMNS)):
                    worksheet.set_column(col_idx, col_idx, 20, text_fmt)
            
            # Save CSV
            df_out.to_csv(out_path_csv, index=False, encoding="utf-8-sig")
            
            print(f"\n  [MERGED EXPORT] Excel saved → {out_path_xlsx}")
            print(f"  [MERGED EXPORT] CSV  saved → {out_path_csv}")
            break
        except PermissionError:
            input(f"\n  [!] File '{out_path_xlsx}' or '{out_path_csv}' đang mở. Đóng rồi nhấn Enter...")
    return out_path_xlsx


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
        original_sku = target_sku
        portfolio_id = sku_portfolio_map.get(target_sku, "")
        
        # Recover full SKU if truncated to 31 chars
        if not portfolio_id:
            for map_sku, pid in sku_portfolio_map.items():
                if map_sku[:31] == target_sku:
                    target_sku = map_sku
                    portfolio_id = pid
                    break
                    
        if not portfolio_id or portfolio_id.lower() == "nan":
            print(f"\n  [WARN] '{target_sku}' không có Portfolio ID → để trống")
            portfolio_id = ""

        print(f"\n  SKU: {target_sku} (Sheet: {original_sku})  |  Portfolio: {portfolio_id}  |  {len(df_sku)} dòng")

        col_camp = find_col(df_sku, "Campaign Name")
        col_kw   = find_col(df_sku, "Target")
        col_note = find_col(df_sku, "Ghi chú") or find_col(df_sku, "Note")
        col_type = find_col(df_sku, "Loại Campaign")
        col_ts   = find_col(df_sku, "Trạng thái") or find_col(df_sku, "Trạng Thái") or find_col(df_sku, "Status")

        if not col_kw:
            print(f"  [ERROR] Thiếu cột 'Target'. Cột hiện có: {list(df_sku.columns)}")
            continue

        sku_ok   = 0
        sku_skip = 0
        
        seen_targets = set()

        for idx, row in df_sku.iterrows():
            raw_keyword   = str(row[col_kw]).strip()
            campaign_name = str(row[col_camp]).strip() if col_camp else ""
            note_str      = str(row[col_note]).strip() if col_note else ""
            status_val    = str(row[col_ts]).strip().lower() if col_ts else ""

            # Chỉ xử lý các dòng có trạng thái "Chưa tạo" hoặc để trống (do bước 1 có thể xuất rỗng)
            if status_val not in ("chưa tạo", ""):
                sku_skip += 1
                continue

            # Skip dòng không có keyword
            if not raw_keyword or raw_keyword.lower() == "nan":
                sku_skip += 1
                continue

            # Skip dòng không có campaign name (lẽ ra đã được generate ở bước 1)
            if not campaign_name or campaign_name.lower() == "nan":
                print(f"    [SKIP] Dòng {idx+2}: Campaign Name rỗng — '{raw_keyword[:30]}'")
                sku_skip += 1
                continue

            camp_type_str = str(row[col_type]).strip() if col_type else ""
            is_pt = ("product targeting" in camp_type_str.lower()
                     or camp_type_str.strip().upper() == "PT")

            # ── Parse Ghi chú ────────────────────────────────────────────────
            try:
                parsed = parse_note(note_str, campaign_name)
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

            # --- LỌC TRÙNG (DEDUPLICATION) ---
            target_key = (raw_keyword.lower(), (match_type or "unknown").lower())
            if target_key in seen_targets:
                print(f"    [SKIP] Dòng {idx+2}: Trùng lặp Keyword '{raw_keyword}' (Match Type: {match_type})")
                sku_skip += 1
                continue
            seen_targets.add(target_key)
            # ---------------------------------

            block = build_7_row_block(
                campaign_name  = campaign_name,
                target_sku     = target_sku,
                raw_keyword    = raw_keyword,
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

    # ── Gộp tất cả campaign rows và xuất file Master Create ──────────────────
    all_rows = []
    for sku, rows in sku_buckets.items():
        all_rows.extend(rows)
    export_merged_file("Musemory", "CreateCampaigns", all_rows, OUTPUT_DIR)

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
