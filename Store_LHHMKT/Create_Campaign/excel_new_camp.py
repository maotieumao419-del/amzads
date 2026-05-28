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

# Status hợp lệ mặc định (để tham khảo)
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
    """
    Sử dụng if-else liên tục để bao phủ các trường hợp người dùng nhập sai chính tả 
    hoặc dùng từ đồng nghĩa trong sheet Listing.
    """
    s = status_str.strip().lower()
    if not s:
        return False
        
    # Các trường hợp hợp lệ để BẮT ĐẦU xử lý
    if "công việc mới" in s or "new" in s or "mới" in s:
        return True
    elif "tiếp nhận" in s or "pending" in s or "đang xử lý" in s:
        return True
    elif "đã tạo" in s or "done" in s or "hoàn thành" in s or "xong" in s:
        return True
    elif "update" in s or "cập nhật" in s or "sửa" in s:
        return True
    else:
        # Trường hợp không rõ ràng thì mặc định bỏ qua để an toàn
        return False


def is_active_row(trang_thai_str: str) -> bool:
    """
    Kiểm tra xem dòng trong SKU sheet đã được xử lý/chạy chưa.
    - Trả về True (Bỏ qua) nếu dòng đó đã có trạng thái hoàn thành.
    - Trả về False (Xử lý) nếu rỗng hoặc yêu cầu tạo mới.
    """
    s = trang_thai_str.strip().lower()
    
    if not s or "chưa tạo" in s or "công việc mới" in s or "cần tạo" in s:
        return False
        
    if "active" in s or "đang chạy" in s or "running" in s:
        return True
    elif "done" in s or "hoàn thành" in s or "xong" in s:
        return True
    elif "đã tạo" in s or "created" in s:
        return True
    elif "pause" in s or "dừng" in s or "ngừng" in s:
        return True
    elif "archive" in s or "xóa" in s or "lưu trữ" in s:
        return True
    else:
        # Nếu ghi chú linh tinh không khớp bất kỳ rules nào nhưng có chữ, 
        # ta tạm coi là chưa xử lý (False) để phòng hờ sót việc.
        return False


# =============================================================================
# NOTE PARSER
# =============================================================================

def parse_note(note: str, campaign_name: str = "") -> dict:
    """
    Phân tích Ghi chú → dict với match_type, bid, placements, placement_tag.
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
        "match_type":    match_type,
        "bid":           round(float(bid), 4),
        "placements":    placements,
        "placement_tag": placement_tag,
    }


# =============================================================================
# CAMPAIGN NAME GENERATOR
# =============================================================================

# Dùng một dictionary (cache) để lưu SPR cho các keyword giống nhau
# tránh việc hỏi đi hỏi lại cùng 1 keyword nhiều lần trong 1 lần chạy
_SPR_CACHE = {}

def build_campaign_name(sku: str, type_code: str, match_type: str,
                        raw_keyword: str, placement_tag: str, creator_name: str = "LHH") -> str:
    """
    Tạo tên Campaign mới. (Scenario 2 - C_mass_sop_factory)
    - Keyword (Exact): {SKU}_SP00_{SPRxx}_{Keyword}_{Matchtype}_{Creator}_{DDMMYY}
    - Keyword (Khác): {SKU}_SP00_{Keyword}_{Matchtype}_{Creator}_{DDMMYY}
    - Auto: {SKU}_SP02_AUTO_{Creator}_{DDMMYY}
    - PT: {SKU}_SP03_PT_{Target}_{Creator}_{DDMMYY}
    """
    from datetime import datetime
    current_date = datetime.now().strftime("%d%m%y")
    
    mt = str(match_type).lower().strip() if match_type else "unknown"
    tc = str(type_code).upper().strip() if type_code else "KT"
    
    is_auto = (tc == "AU" or mt == "auto")
    is_pt = (tc == "PT" or mt == "targeting" or mt == "pt")
    
    mt_capitalized = mt.capitalize() if mt else "Unknown"
    
    if is_auto:
        components = [sku, "SP02", "AUTO", creator_name, current_date]
    elif is_pt:
        clean_keyword = re.sub(r'[^\w\s-]', '', raw_keyword).strip()
        pt_string = f"PT_{clean_keyword}"
        components = [sku, "SP03", pt_string, creator_name, current_date]
    else:
        # Keyword Campaign (SP00)
        clean_keyword = re.sub(r'[^\w\s-]', '', raw_keyword).strip()
        keyword_slug = " ".join(clean_keyword.split()[:5])
        
        # Nếu là exact, yêu cầu user nhập SPR qua console
        if mt == "exact":
            if keyword_slug not in _SPR_CACHE:
                while True:
                    spr_input = input(f"\n[?] Nhập SPR cho keyword '{raw_keyword}' (2 chữ số): ").strip()
                    if not spr_input:
                        spr_input = "00"
                    
                    val_to_check = spr_input.upper()
                    if val_to_check.startswith("SPR"):
                        val_to_check = val_to_check[3:]
                        
                    if len(val_to_check) > 2:
                        print("  [!] Lỗi: SPR chỉ được phép tối đa 2 chữ số. Vui lòng nhập lại.")
                        continue
                        
                    if val_to_check.isdigit() and len(val_to_check) == 1:
                        val_to_check = f"0{val_to_check}"
                    elif not val_to_check:
                        val_to_check = "00"
                        
                    spr_val = f"SPR{val_to_check}"
                    _SPR_CACHE[keyword_slug] = spr_val
                    break
            
            spr_final = _SPR_CACHE[keyword_slug]
            components = [sku, "SP00", spr_final, keyword_slug, mt_capitalized, creator_name, current_date]
        else:
            components = [sku, "SP00", keyword_slug, mt_capitalized, creator_name, current_date]
        
    raw_name = "_".join(components)
    final_name = re.sub(r'_+', '_', raw_name)
    
    return final_name


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

    # Dùng dictionary để gộp các SKU trùng lặp
    sku_dict = {}
    for _, row in df.iterrows():
        sku    = str(row[col_sku]).replace("\n", "").strip()
        status = str(row[col_status]).strip()

        if not sku or sku.lower() == "nan":
            continue

        pid   = str(row[col_pid]).strip() if col_pid else ""
        store = str(row[col_store]).strip() if col_store else ""

        if sku not in sku_dict:
            sku_dict[sku] = {
                "sku": sku, 
                "portfolio_id": pid, 
                "store": store, 
                "statuses": [status]
            }
        else:
            sku_dict[sku]["statuses"].append(status)
            # Nếu portfolio_id rỗng ở dòng trước, cập nhật nếu dòng này có
            if not sku_dict[sku]["portfolio_id"] and pid:
                sku_dict[sku]["portfolio_id"] = pid

    results = []
    for sku, info in sku_dict.items():
        # Kiểm tra xem có BẤT KỲ status nào hợp lệ không
        valid_status = None
        for s in info["statuses"]:
            if is_valid_listing_status(s):
                valid_status = s
                break
                
        if valid_status is not None:
            results.append({"sku": sku, "portfolio_id": info["portfolio_id"], "store": info["store"]})
            print(f"  [OK] {sku!r:42s}  Portfolio={info['portfolio_id']}  Status='{valid_status}' (Gộp từ {len(info['statuses'])} dòng)")
        else:
            print(f"  [SKIP] '{sku}': Statuses={info['statuses']} → bỏ qua")

    print(f"\n  → Tổng SKU cần xử lý (Unique): {len(results)}")
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

def step3_load_sku_sheet(xls: pd.ExcelFile, sku: str, creator_name: str = "LHH", sheet_name: str = None) -> pd.DataFrame:
    """
    Đọc sheet SKU:
    1. Tự động nhận diện dòng Header (chống lỗi do dòng Full SKU ở trên cùng).
    2. Forward-fill STT / Ghi chú / Loại Campaign theo nhóm.
    3. SKIP dòng có Trạng thái = Active/Done.
    4. Tách Target multi-line.
    5. Parse Ghi chú → Campaign Name.
    """
    if sheet_name is None:
        sheet_name = sku
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    
    # --- CHUẨN HÓA CẤU TRÚC SHEET (DYNAMIC HEADER DETECTION) ---
    # Kiểm tra xem header chuẩn có nằm ngay ở df_raw.columns không?
    has_target = False
    for c in df_raw.columns:
        if "target" in str(c).strip().lower():
            has_target = True
            break
            
    if not has_target:
        # Nếu không có, quét 5 dòng đầu của data để tìm dòng thực sự chứa Header
        header_row_idx = -1
        for idx, row in df_raw.head(5).iterrows():
            row_str = " ".join([str(x).lower() for x in row.values])
            if "target" in row_str and "stt" in row_str:
                header_row_idx = idx
                break
                
        if header_row_idx != -1:
            # Lấy dòng đó làm header
            df_raw.columns = df_raw.iloc[header_row_idx].astype(str).str.strip()
            # Lấy dữ liệu từ dòng bên dưới header trở đi
            df_raw = df_raw.iloc[header_row_idx + 1:].reset_index(drop=True)
        else:
            raise ValueError(f"Không tìm thấy dòng Header chứa cột 'Target' trong sheet {sku}")
    else:
        df_raw.columns = df_raw.columns.str.strip()

    df_raw.fillna("", inplace=True)

    col_target = find_col(df_raw, COL_TARGET)
    col_note   = find_col(df_raw, COL_NOTE)
    col_camp   = normalize_col(df_raw, COL_CAMP_NAME)
    col_type   = normalize_col(df_raw, COL_LOOP_TYPE)
    
    # Hỗ trợ nhiều tên cột trạng thái, nếu không có thì tạo mới
    col_ts = None
    for name in ["Status", "Trạng thái", "Trạng Thái"]:
        col_ts = normalize_col(df_raw, name)
        if col_ts:
            break
            
    if not col_ts:
        col_ts = "Status"
        df_raw[col_ts] = ""

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

    # ── 3b. Forward-fill STT / Ghi chú / Loại Campaign ───────────
    # STT, Ghi chú và Loại Campaign được share trong 1 nhóm dòng → cần ffill.
    # !! Campaign Name và Trạng thái KHÔNG ffill để đảm bảo tính độc nhất !!
    if col_stt:
        df_raw[col_stt] = df_raw[col_stt].replace("", None).ffill().fillna("")
        
    if col_note:
        df_raw[col_note] = df_raw[col_note].replace("", None).ffill().fillna("")

    if col_type:
        df_raw[col_type] = df_raw[col_type].replace("", None).ffill().fillna("")

    # ── 3c. Đếm dòng SKIP vs dòng MỚI ──────────────────────────────────────
    skip_active = 0
    new_rows    = 0

    expanded_rows = []
    seen_targets = set()

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
        camp_val = str(row[col_camp]).strip() if col_camp else ""
        
        # Parse note để lấy placement_tag và match_type
        parsed = parse_note(note_val, camp_val)

        for kw in keywords:
            # --- LỌC TRÙNG (DEDUPLICATION) ---
            # Nếu 1 keyword xuất hiện 2 lần với cùng 1 match type thì lọc bỏ đi
            current_match_type = parsed["match_type"] or "unknown"
            target_key = (kw.lower(), current_match_type.lower())
            if target_key in seen_targets:
                continue
            seen_targets.add(target_key)
            # ---------------------------------

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
                raw_keyword   = kw,
                placement_tag = parsed["placement_tag"],
                creator_name  = creator_name,
            )

            expanded_rows.append(new_row)
            new_rows += 1

    if not expanded_rows:
        print(f"  [BƯỚC 3] SKU '{sku}': Không có dòng mới nào (skip_active={skip_active})")
        return pd.DataFrame(columns=df_raw.columns)

    df_out = pd.DataFrame(expanded_rows, columns=df_raw.columns)
    df_out = df_out.reset_index(drop=True)

    # Lọc 6 cột chuẩn (Sử dụng tên cột trạng thái thực tế col_ts)
    out_cols = ["STT", "Campaign Name", "Loại Campaign", "Target", "Note", col_ts]
    df_final = pd.DataFrame(columns=out_cols)

    df_final["STT"] = df_out[col_stt] if col_stt else range(1, len(df_out) + 1)
    df_final["Campaign Name"] = df_out[col_camp] if col_camp else ""
    df_final["Loại Campaign"] = df_out[col_type] if col_type else ""
    df_final["Target"] = df_out[col_target] if col_target else ""
    df_final["Note"] = df_out[col_note] if col_note else ""
    df_final[col_ts] = df_out[col_ts]

    print(f"  [BƯỚC 3] SKU '{sku}': skip_active={skip_active} | new_keywords={new_rows} dòng")
    return df_final


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def main():
    print("\n" + "=" * 60)
    print("  EXCEL NEW CAMP — Bước 1/2 Pipeline C_mass_sop_factory")
    print(f"  Input 1 : {INPUT_1_DIR}")
    print(f"  Input 2 : {INPUT_2_DIR}")
    print("=" * 60)

    creator_name = input("\n[?] Nhập tên người tạo campaign (VD: LHH, Nguyen...): ").strip()
    if not creator_name:
        creator_name = "LHH"
        print(f"  -> Không nhập, mặc định sử dụng: {creator_name}")

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
        
        target_sheet = sku
        if sku not in available_sheets:
            if sku[:31] in available_sheets:
                target_sheet = sku[:31]
            else:
                print(f"  [WARN] Không có sheet tương ứng → Bỏ qua.")
                continue
        try:
            df_result = step3_load_sku_sheet(xls, sku, creator_name=creator_name, sheet_name=target_sheet)
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
                    # Xuất data từ dòng 3 (startrow=2), bỏ qua header mặc định
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=2)
                    workbook  = writer.book
                    worksheet = writer.sheets[sheet_name]

                    # 1. Row 1: Tiêu đề Sheet gộp A1:F1
                    format_title = workbook.add_format({
                        'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'
                    })
                    worksheet.merge_range('A1:F1', sku, format_title)

                    # 2. Row 2: Tiêu đề Cột
                    format_header = workbook.add_format({'bold': True, 'border': 1})
                    format_note_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#ACD1EC'})

                    headers = df.columns.tolist()
                    for col_num, value in enumerate(headers):
                        if value == 'Note':
                            worksheet.write(1, col_num, value, format_note_header)
                        else:
                            worksheet.write(1, col_num, value, format_header)

                    # 3. Chỉnh kích thước cột
                    text_fmt  = workbook.add_format({"num_format": "@"})
                    worksheet.set_column(0, 0, 5, text_fmt)   # STT
                    worksheet.set_column(1, 1, 40, text_fmt)  # Campaign Name
                    worksheet.set_column(2, 2, 20, text_fmt)  # Loại Campaign
                    worksheet.set_column(3, 3, 20, text_fmt)  # Target
                    worksheet.set_column(4, 4, 35, text_fmt)  # Note
                    worksheet.set_column(5, 5, 12, text_fmt)  # Trạng thái
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
