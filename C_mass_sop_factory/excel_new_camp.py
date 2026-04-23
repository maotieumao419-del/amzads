# =============================================================================
# FILE: excel_new_camp.py  (Self-contained — thuộc C_mass_sop_factory)
# MỤC ĐÍCH: Bước 1/2 pipeline — đọc PPC_*.xlsx từ "data/input 1/",
#           xử lý, xuất PPC_PROCESSED_OUTPUT.xlsx vào "data/input 2/"
#           để mass_sop_factory.py đọc tiếp.
#
# PIPELINE:
#   Bước 1 → Quét sheet Listing: lấy danh sách SKU cần xử lý
#   Bước 2 → Xác thực Portfolio (Log Only)
#   Bước 3 → Đọc từng sheet SKU, SKIP dòng Active, tách Target multi-line,
#             parse Ghi chú, generate Campaign Name mới
#
# CẤU TRÚC "GHI CHÚ" — 2 format:
#   Format cũ: "Phrase 0.45, 30TRP"  /  "exact 0.3_35TPR"  /  "exact"
#   Format mới: "exact, 50T, 0.7Đ"  /  "Phrase, 20TRP, 0.5Đ"  /  "exact, 50RP, 0.68Đ"
#
# QUY TẮC TÊN CAMPAIGN: [SKU]_[LoaiCamp]_[match_type]_[Target]_[Placement]
#   Ví dụ: DEVOM_KT_exact_250 anniversary_50T
#
# INPUT:  data/input 1/PPC_*.xlsx
# OUTPUT: data/input 2/PPC_PROCESSED_OUTPUT.xlsx
# =============================================================================

import os
import re
import glob
import pandas as pd

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUT_1_DIR = os.path.join(SCRIPT_DIR, "data", "input 1")
INPUT_2_DIR = os.path.join(SCRIPT_DIR, "data", "input 2")
OUTPUT_FILE = os.path.join(INPUT_2_DIR, "PPC_PROCESSED_OUTPUT.xlsx")

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------
LISTING_SHEET   = "Listing"
PORTFOLIO_SHEET = "Portfolio ID"

# Status hợp lệ trong Listing để chọn SKU cần xử lý
VALID_LISTING_STATUSES = [
    "công việc mới",
    "tiếp nhận",
    "đã tạo campaign",
]

# Trạng thái của dòng trong SKU sheet → BỎ QUA nếu là Active
SKIP_ROW_STATUS = "active"

COL_SKU          = "SKU"
COL_PORTFOLIO_ID = "Portfolio Id"
COL_STATUS       = "Status"       # trong Listing
COL_TRANG_THAI   = "Trạng thái"  # trong SKU sheet
COL_TARGET       = "Target"
COL_NOTE         = "Ghi chú"
COL_CAMP_NAME    = "Campaign Name"
COL_LOOP_TYPE    = "Loại Campaign"
COL_STT          = "STT"

DEFAULT_BID      = 0.50


# =============================================================================
# UTILITIES
# =============================================================================

def normalize_col(df: pd.DataFrame, col_name: str) -> str | None:
    for c in df.columns:
        if c.strip().lower() == col_name.strip().lower():
            return c
    return None


def find_col(df: pd.DataFrame, col_name: str) -> str:
    found = normalize_col(df, col_name)
    if found is None:
        raise ValueError(
            f"[ERROR] Không tìm thấy cột '{col_name}'. "
            f"Các cột hiện có: {list(df.columns)}"
        )
    return found


def is_valid_listing_status(status_str: str) -> bool:
    s = status_str.strip().lower()
    return any(s == v for v in VALID_LISTING_STATUSES)


def is_active_row(trang_thai_str: str) -> bool:
    """
    Kiểm tra xem dòng đã được tạo campaign chưa.
    - Trả về True (Bỏ qua) nếu: 'Active', 'Done', 'Đã tạo'
    - Trả về False (Xử lý) nếu: rỗng, 'Chưa tạo', hoặc bất kỳ giá trị nào khác.
    """
    s = trang_thai_str.strip().lower()
    if s in ("active", "done", "đã tạo"):
        return True
    return False


# =============================================================================
# NOTE PARSER
# =============================================================================

def parse_note(note: str) -> dict:
    """
    Phân tích ô Ghi chú.
    """
    note_lower = note.lower()

    # 1. Match type
    m_type = re.search(r'(?<![a-z])(exact|phrase|broad)(?![a-z])', note_lower)
    match_type = m_type.group(1) if m_type else None

    # 2. Placements
    placements = {"T": 0, "R": 0, "P": 0}
    placement_tags = []

    p_matches = re.findall(r'(\d+)\s*[_,]?\s*([trp]{1,3})\b', note_lower)
    for num_str, keys_str in p_matches:
        val = int(num_str)
        tag_keys = ""
        if 't' in keys_str:
            placements["T"] = val
            tag_keys += "T"
        if 'r' in keys_str:
            placements["R"] = val
            tag_keys += "R"
        if 'p' in keys_str:
            placements["P"] = val
            tag_keys += "P"
        if tag_keys:
            placement_tags.append(f"{val}{tag_keys}")

    seen_tags = set()
    unique_tags = []
    for tag in placement_tags:
        if tag not in seen_tags:
            seen_tags.add(tag)
            unique_tags.append(tag)
    placement_tag = "_".join(unique_tags) if unique_tags else "00T"

    # 3. Bid
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
        "match_type":    match_type,
        "bid":           round(float(bid), 4),
        "placements":    placements,
        "placement_tag": placement_tag,
    }


# =============================================================================
# CAMPAIGN NAME GENERATOR
# =============================================================================

def build_campaign_name(sku: str, type_code: str, match_type: str,
                        target: str, placement_tag: str) -> str:
    """
    Tạo tên Campaign độc nhất.
    Format: [SKU]_[LoaiCamp]_[match_type]_[Target]_[Placement]
    """
    mt = match_type if match_type else "unknown"
    # Clean target: bỏ ký tự đặc biệt, lấy tối đa 3 từ hoặc 30 ký tự để tên không quá dài
    clean_target = re.sub(r'[^\w\s-]', '', target).strip()
    target_slug = "_".join(clean_target.split()[:5]) # Lấy tối đa 5 từ đầu
    
    return f"{sku}_{type_code}_{mt}_{target_slug}_{placement_tag}"


# =============================================================================
# BƯỚC 1 – QUÉT LISTING
# =============================================================================

def step1_load_listing(xls: pd.ExcelFile) -> list[dict]:
    print("\n" + "=" * 60)
    print("[BƯỚC 1] Quét sheet Listing")
    print("=" * 60)

    df = pd.read_excel(xls, sheet_name=LISTING_SHEET, dtype=str)
    df.columns = df.columns.str.strip()
    df.fillna("", inplace=True)

    col_sku    = find_col(df, COL_SKU)
    col_pid    = find_col(df, COL_PORTFOLIO_ID)
    col_status = find_col(df, COL_STATUS)
    col_store  = normalize_col(df, "Store")

    print(f"  Columns: {list(df.columns)}")
    print(f"  Status values: {df[col_status].unique().tolist()}")

    results = []
    for _, row in df.iterrows():
        sku    = str(row[col_sku]).replace("\n", "").strip()
        status = str(row[col_status]).strip()

        if not sku or sku.lower() == "nan":
            continue
        if not is_valid_listing_status(status):
            print(f"  [SKIP] '{sku}': Status='{status}' → bỏ qua")
            continue

        pid   = str(row[col_pid]).strip() if col_pid else ""
        store = str(row[col_store]).strip() if col_store else ""
        results.append({"sku": sku, "portfolio_id": pid, "store": store})
        print(f"  [OK] {sku!r:42s}  Portfolio={pid}  Status='{status}'")

    print(f"\n  → Tổng SKU cần xử lý: {len(results)}")
    return results


# =============================================================================
# BƯỚC 2 – XÁC THỰC PORTFOLIO (LOG ONLY)
# =============================================================================

def step2_validate_portfolio(xls: pd.ExcelFile, sku_rows: list[dict]) -> None:
    print("\n" + "=" * 60)
    print("[BƯỚC 2] Xác thực Portfolio ID (Log Only)")
    print("=" * 60)
    try:
        df_pid = pd.read_excel(xls, sheet_name=PORTFOLIO_SHEET, dtype=str)
        df_pid.columns = df_pid.columns.str.strip()
        df_pid.fillna("", inplace=True)
    except Exception as e:
        print(f"  [BYPASS] Không đọc được sheet '{PORTFOLIO_SHEET}': {e}")
        return

    all_pid_values: set[str] = set()
    for col in df_pid.columns:
        all_pid_values.update(df_pid[col].astype(str).str.strip().tolist())
    all_pid_values.discard("")

    for item in sku_rows:
        pid = item["portfolio_id"]
        if not pid or pid.lower() in ("", "nan"):
            print(f"  [WARN] '{item['sku']}' — Portfolio ID rỗng.")
        elif pid in all_pid_values:
            print(f"  [OK]   '{item['sku']}' — Portfolio '{pid}' tìm thấy.")
        else:
            print(f"  [WARN] '{item['sku']}' — Portfolio '{pid}' KHÔNG tìm thấy!")


# =============================================================================
# BƯỚC 3 – ĐỌC SKU SHEET: SKIP ACTIVE, PARSE NOTE, GENERATE CAMPAIGN NAME
#
# Logic xử lý dòng:
#   - Dòng có Trạng thái = Active/Done → SKIP toàn bộ group (tất cả dòng
#     liên tiếp có cùng "Ghi chú" group đó)
#   - Dòng mới (Trạng thái rỗng, Campaign Name rỗng) → TẠO MỚI
#   - Target multi-line → tách thành nhiều dòng riêng
# =============================================================================

def step3_load_sku_sheet(xls: pd.ExcelFile, sku: str) -> pd.DataFrame:
    """
    Đọc sheet SKU:
    1. Forward-fill Ghi chú / Campaign Name / Loại Campaign theo nhóm
       *** Trạng thái KHÔNG ffill — rỗng = dòng mới, Active/Done = đã tạo ***
    2. SKIP dòng có Trạng thái = Active/Done
    3. Tách Target multi-line
    4. Parse Ghi chú → Campaign Name
    """
    df_raw = pd.read_excel(xls, sheet_name=sku, dtype=str)
    df_raw.columns = df_raw.columns.str.strip()
    df_raw.fillna("", inplace=True)

    col_target = find_col(df_raw, COL_TARGET)
    col_note   = find_col(df_raw, COL_NOTE)
    col_camp   = normalize_col(df_raw, COL_CAMP_NAME)
    col_type   = normalize_col(df_raw, COL_LOOP_TYPE)
    col_ts     = normalize_col(df_raw, COL_TRANG_THAI)
    col_stt    = normalize_col(df_raw, COL_STT)

    # ── 3a. Xác định type_code ───────────────────────────────────────────────
    TYPE_CODE_MAP = {
        "keyword":           "KT",
        "keyword (từ khóa)": "KT",
        "product targeting": "PT",
        "auto":              "AU",
    }
    type_code = "KT"
    if col_type:
        first_type = next(
            (str(v).strip() for v in df_raw[col_type] if str(v).strip() and str(v).strip().lower() != "nan"),
            ""
        )
        type_code = TYPE_CODE_MAP.get(first_type.strip().lower(), "KT")
    print(f"  [INFO] Loại Campaign = '{first_type if col_type else '?'}'  →  type_code='{type_code}'")

    # ── 3b. Forward-fill Ghi chú / Loại Campaign ───────────
    # Ghi chú và Loại Campaign được share trong 1 nhóm dòng → cần ffill.
    # !! Campaign Name và Trạng thái KHÔNG ffill để đảm bảo tính độc nhất !!
    if col_note:
        df_raw[col_note] = df_raw[col_note].replace("", None).ffill().fillna("")

    if col_type:
        df_raw[col_type] = df_raw[col_type].replace("", None).ffill().fillna("")

    # ── 3c. Đếm dòng SKIP vs dòng MỚI ──────────────────────────────────────
    skip_active = 0
    new_rows    = 0

    expanded_rows = []

    for _, row in df_raw.iterrows():
        # Lấy Trạng thái GỐC (không ffill) — rỗng/Chưa tạo = dòng mới, Active = đã tạo
        ts_val = str(row[col_ts]).strip() if col_ts else ""
        if is_active_row(ts_val):
            skip_active += 1
            continue

        # Lấy Target
        raw_target = str(row[col_target])
        keywords = [kw.strip() for kw in raw_target.split("\n") if kw.strip()]
        if not keywords:
            continue

        note_val = str(row[col_note]).strip() if col_note else ""
        
        # Parse note để lấy placement_tag và match_type
        parsed = parse_note(note_val)

        for kw in keywords:
            new_row = row.copy()
            new_row[col_target] = kw
            
            # Đánh dấu trạng thái là "Chưa tạo" cho các dòng mới
            if col_ts:
                new_row[col_ts] = "Chưa tạo"

            # LUÔN tạo Campaign Name mới cho dòng chưa Active
            new_row[col_camp] = build_campaign_name(
                sku           = sku,
                type_code     = type_code,
                match_type    = parsed["match_type"] or "unknown",
                target        = kw,
                placement_tag = parsed["placement_tag"],
            )

            expanded_rows.append(new_row)
            new_rows += 1

    if not expanded_rows:
        print(f"  [BƯỚC 3] SKU '{sku}': Không có dòng mới nào (skip_active={skip_active})")
        return pd.DataFrame(columns=df_raw.columns)

    df_out = pd.DataFrame(expanded_rows, columns=df_raw.columns)
    df_out = df_out.reset_index(drop=True)
    # Đánh lại STT
    if col_stt:
        df_out[col_stt] = range(1, len(df_out) + 1)
    else:
        df_out.insert(0, COL_STT, range(1, len(df_out) + 1))

    print(f"  [BƯỚC 3] SKU '{sku}': skip_active={skip_active} | new_keywords={new_rows} dòng")
    return df_out


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def main():
    print("\n" + "=" * 60)
    print("  EXCEL NEW CAMP — Bước 1/2 Pipeline C_mass_sop_factory")
    print(f"  Input 1 : {INPUT_1_DIR}")
    print(f"  Input 2 : {INPUT_2_DIR}")
    print("=" * 60)

    os.makedirs(INPUT_1_DIR, exist_ok=True)
    os.makedirs(INPUT_2_DIR, exist_ok=True)

    # ── Tìm file PPC_*.xlsx trong input 1 ────────────────────────────────────
    files = glob.glob(os.path.join(INPUT_1_DIR, "PPC_*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]

    if not files:
        raise FileNotFoundError(
            f"[ERROR] Không tìm thấy PPC_*.xlsx trong: {INPUT_1_DIR}\n"
            f"→ Hãy đặt file Excel vào thư mục đó."
        )

    if len(files) > 1:
        files.sort(key=os.path.getmtime, reverse=True)
        print(f"\n  [WARNING] Nhiều file → chọn mới nhất: {os.path.basename(files[0])}")

    INPUT_FILE = files[0]
    print(f"\n  Input file : {INPUT_FILE}")
    print(f"  Output     : {OUTPUT_FILE}")

    xls = pd.ExcelFile(INPUT_FILE)
    available_sheets = xls.sheet_names
    print(f"  Sheets     : {available_sheets}")

    sku_rows = step1_load_listing(xls)
    if not sku_rows:
        print("\n[DONE] Không có SKU nào cần xử lý.")
        return

    step2_validate_portfolio(xls, sku_rows)

    print("\n" + "=" * 60)
    print("[BƯỚC 3] Xử lý từng SKU (Skip Active → Expand Keywords → Generate Name)")
    print("=" * 60)

    processed_results: dict[str, pd.DataFrame] = {}

    for item in sku_rows:
        sku = item["sku"]
        print(f"\n  --- SKU: {sku} ---")
        if sku not in available_sheets:
            print(f"  [WARN] Không có sheet tương ứng → Bỏ qua.")
            continue
        try:
            df_result = step3_load_sku_sheet(xls, sku)
        except Exception as e:
            print(f"  [ERROR] {e} → Bỏ qua.")
            continue
        if df_result is not None and not df_result.empty:
            processed_results[sku] = df_result
            print(f"  [OK] → {len(df_result)} dòng sẽ được xuất.")
        else:
            print(f"  [SKIP] Không có dòng mới → Bỏ qua file này.")

    if not processed_results:
        print("\n[WARNING] Không có SKU nào có dòng mới cần tạo.")
        return

    # ── Xuất ra input 2 ──────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print(f"[OUTPUT] Xuất → {OUTPUT_FILE}")
    print("=" * 60)

    while True:
        try:
            with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
                for sku, df in processed_results.items():
                    sheet_name = sku[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    workbook  = writer.book
                    worksheet = writer.sheets[sheet_name]
                    text_fmt  = workbook.add_format({"num_format": "@"})
                    for col_idx in range(len(df.columns)):
                        worksheet.set_column(col_idx, col_idx, 30, text_fmt)
            print(f"\n  [OK] Đã lưu: {OUTPUT_FILE}")
            break
        except PermissionError:
            input(f"\n  [!] File đang mở. Đóng Excel rồi nhấn Enter...")

    print("\n" + "=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    total_rows = sum(len(df) for df in processed_results.values())
    print(f"  SKU có dòng mới   : {len(processed_results)}")
    print(f"  SKU không có mới  : {len(sku_rows) - len(processed_results)}")
    print(f"  Tổng keywords mới : {total_rows}")
    print("=" * 60)
    print("\n  → Chạy tiếp: python mass_sop_factory.py\n")


if __name__ == "__main__":
    main()
