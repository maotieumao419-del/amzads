# =============================================================================
# FILE: excel_new_camp.py  (Self-contained — thuộc C_mass_sop_factory)
# MỤC ĐÍCH: Bước 1/2 pipeline hàng loạt — đọc PPC_NGUYEN.xlsx, xử lý 5 bước,
#           xuất PPC_PROCESSED_OUTPUT.xlsx vào cùng thư mục data/input/
#           để mass_sop_factory.py đọc tiếp ngay mà không cần copy tay.
#
# PIPELINE 5 BƯỚC:
#   Bước 1 → Quét Listing & Làm sạch SKU
#   Bước 2 → Xác thực Portfolio (Log Only)
#   Bước 3 → Làm sạch Từ khóa & Nhân bản (Cross Join 3 Match Types)
#   Bước 4 → Chuẩn hóa cột "Loại Campaign" + tạo type_code
#   Bước 5 → Naming Convention động & Đánh STT
#
# INPUT:  data/input/PPC_NGUYEN.xlsx          ← đặt file vào đây
# OUTPUT: data/input/PPC_PROCESSED_OUTPUT.xlsx ← mass_sop_factory.py sẽ đọc từ đây
# =============================================================================

import os
import re
import pandas as pd

# ---------------------------------------------------------------------------
# PATHS — tất cả đều nằm trong C_mass_sop_factory/data/
# ---------------------------------------------------------------------------
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR    = os.path.join(SCRIPT_DIR, "data", "input")

INPUT_FILE   = os.path.join(INPUT_DIR, "PPC_NGUYEN.xlsx")
# Output nằm cùng thư mục input để mass_sop_factory đọc trực tiếp
OUTPUT_FILE  = os.path.join(INPUT_DIR, "PPC_PROCESSED_OUTPUT.xlsx")

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------
LISTING_SHEET    = "Listing"
PORTFOLIO_SHEET  = "Portfolio ID"
STATUS_TRIGGER_NEW = "Công việc mới"
STATUS_TRIGGER_RECV = "Tiếp nhận"

COL_SKU          = "SKU"
COL_PORTFOLIO_ID = "Portfolio Id"
COL_STATUS       = "Trạng thái"
COL_TARGET       = "Target"
COL_NOTE         = "Ghi chú"
COL_CAMP_NAME    = "Campaign Name"
COL_LOOP_TYPE    = "Loại Campaign"
COL_STT          = "STT"


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


# =============================================================================
# BƯỚC 1 – QUÉT LISTING & LÀM SẠCH SKU
# =============================================================================

def step1_load_listing(xls: pd.ExcelFile) -> list[dict]:
    print("\n" + "=" * 60)
    print("[BƯỚC 1] Quét sheet Listing & Làm sạch SKU")
    print("=" * 60)

    df_listing = pd.read_excel(xls, sheet_name=LISTING_SHEET, dtype=str)
    df_listing.columns = df_listing.columns.str.strip()
    df_listing.fillna("", inplace=True)

    col_sku    = find_col(df_listing, COL_SKU)
    col_pid    = find_col(df_listing, COL_PORTFOLIO_ID)
    col_status = find_col(df_listing, COL_STATUS)
    col_store  = normalize_col(df_listing, "Store") # Thêm cột Store theo yêu cầu

    status_series = df_listing[col_status].str.strip().str.lower()
    mask = status_series.str.contains(STATUS_TRIGGER_NEW.lower(), na=False) | \
           status_series.str.contains(STATUS_TRIGGER_RECV.lower(), na=False)
    
    df_filtered = df_listing[mask].copy()

    if df_filtered.empty:
        print(f"  [WARNING] Không có SKU nào có Trạng thái = '{STATUS_TRIGGER_NEW}' hoặc '{STATUS_TRIGGER_RECV}'.")
        return []

    results = []
    for _, row in df_filtered.iterrows():
        sku = row[col_sku].replace("\n", "").strip()
        pid = str(row[col_pid]).strip() if col_pid else ""
        store = str(row[col_store]).strip() if col_store else ""
        
        if not sku:
            print("  [SKIP] Dòng có SKU rỗng — bỏ qua.")
            continue
            
        results.append({"sku": sku, "portfolio_id": pid, "store": store})
        print(f"  [OK] SKU: {sku!r:30s}  |  Portfolio ID: {pid}  |  Store: {store}")

    print(f"\n  → Tổng SKU cần xử lý: {len(results)}")
    return results


# =============================================================================
# BƯỚC 2 – XÁC THỰC PORTFOLIO (LOG ONLY, BYPASS)
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

    all_pid_values = set()
    for col in df_pid.columns:
        all_pid_values.update(df_pid[col].astype(str).str.strip().tolist())
    all_pid_values.discard("")

    for item in sku_rows:
        pid = item["portfolio_id"]
        if not pid or pid.lower() in ("", "nan"):
            print(f"  [WARN] SKU '{item['sku']}' — Portfolio ID rỗng.")
        elif pid in all_pid_values:
            print(f"  [OK]   SKU '{item['sku']}' — Portfolio '{pid}' tìm thấy.")
        else:
            print(f"  [WARN] SKU '{item['sku']}' — Portfolio '{pid}' KHÔNG tìm thấy!")


# =============================================================================
# BƯỚC 3 – LÀM SẠCH TỪ KHÓA & NHÂN BẢN (CROSS JOIN 3 MATCH TYPES)
# =============================================================================

def _extract_note_patterns(notes_series: pd.Series) -> tuple[str, str, str]:
    re_exact  = re.compile(r"(?i).*\bexact\b.*")
    re_phrase = re.compile(r"(?i).*\bphrase\b.*")
    re_broad  = re.compile(r"(?i).*\bbroad\b.*")

    exact_note = phrase_note = broad_note = None

    for note in notes_series.dropna():
        note_str = str(note).strip()
        if not note_str:
            continue
        if exact_note  is None and re_exact.match(note_str):  exact_note  = note_str
        if phrase_note is None and re_phrase.match(note_str): phrase_note = note_str
        if broad_note  is None and re_broad.match(note_str):  broad_note  = note_str
        if exact_note and phrase_note and broad_note:
            break

    return (
        exact_note  or "Exact_pending",
        phrase_note or "Phrase_pending",
        broad_note  or "Broad_pending",
    )


def step3_clean_and_crossjoin(xls: pd.ExcelFile, sku: str) -> pd.DataFrame:
    df_sku = pd.read_excel(xls, sheet_name=sku, dtype=str)
    df_sku.columns = df_sku.columns.str.strip()
    df_sku.fillna("", inplace=True)

    col_target = find_col(df_sku, COL_TARGET)
    col_note   = find_col(df_sku, COL_NOTE)

    # Đếm số dòng ghi chú không rỗng
    non_empty_notes = df_sku[col_note].str.strip().replace("", float("nan")).dropna()
    is_fully_annotated = len(non_empty_notes) > 3

    exact_note, phrase_note, broad_note = _extract_note_patterns(df_sku[col_note])

    df_sku = df_sku[df_sku[col_target].str.strip() != ""].copy()
    
    if df_sku.empty:
        return df_sku

    df_sku[col_target] = (
        df_sku[col_target]
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\n", " ", regex=True)
        .str.strip()
    )

    if is_fully_annotated:
        print(f"  [BƯỚC 3] SKU '{sku}': Cột Ghi chú có {len(non_empty_notes)} dòng (>3) → Giữ nguyên các dòng, không nhân 3.")
        return df_sku
        
    print(f"  [BƯỚC 3] SKU '{sku}': Cột Ghi chú có <= 3 dòng → Lọc Target Unique & Nhân 3.")
    
    seen, unique_targets = set(), []
    for t in df_sku[col_target]:
        if t not in seen:
            seen.add(t)
            unique_targets.append(t)

    chunks = []
    for note_val in [exact_note, phrase_note, broad_note]:
        chunk = pd.DataFrame({col_target: unique_targets})
        chunk[col_note] = note_val
        chunks.append(chunk)

    df_crossed = pd.concat(chunks, ignore_index=True)

    for col in df_sku.columns:
        if col not in df_crossed.columns:
            df_crossed[col] = ""

    original_cols = [c for c in df_sku.columns if c in df_crossed.columns]
    extra_cols    = [c for c in df_crossed.columns if c not in original_cols]
    return df_crossed[original_cols + extra_cols]


# =============================================================================
# BƯỚC 4 – CHUẨN HÓA CỘT "LOẠI CAMPAIGN" & TẠO TYPE_CODE
# =============================================================================

def step4_normalize_campaign_type(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    TYPE_CODE_MAP = {
        "keyword":           "KT",
        "keyword (từ khóa)": "KT",
        "product targeting": "PT",
        "auto":              "AU",
    }
    DEFAULT_TYPE = "Keyword (Từ khóa)"
    DEFAULT_CODE = "KT"

    col_type = normalize_col(df, COL_LOOP_TYPE)
    if col_type is None:
        df[COL_LOOP_TYPE] = DEFAULT_TYPE
        return df, DEFAULT_CODE

    first_valid = next(
        (str(v).strip() for v in df[col_type]
         if str(v).strip() and str(v).strip().lower() != "nan"),
        None
    )
    first_valid = first_valid or DEFAULT_TYPE
    df[col_type] = first_valid
    type_code    = TYPE_CODE_MAP.get(first_valid.strip().lower(), "KT")
    print(f"  [BƯỚC 4] Loại Campaign = '{first_valid}'  →  type_code = '{type_code}'")
    return df, type_code


# =============================================================================
# BƯỚC 5 – NAMING CONVENTION ĐỘNG & ĐÁNH STT
# =============================================================================

def _parse_tpr(note: str) -> str:
    m = re.search(r"(\d+)\s*(?:TRP|TPR|T(?![A-Z]))", note, re.IGNORECASE)
    return f"{m.group(1)}TRP" if m else "00TRP"


def _parse_match_type(note: str) -> str:
    m = re.search(r"\b(exact|phrase|broad)\b", note, re.IGNORECASE)
    return m.group(1).lower() if m else "unknown"


def step5_naming_and_stt(df: pd.DataFrame, sku: str, type_code: str) -> pd.DataFrame:
    col_note      = find_col(df, COL_NOTE)
    col_target    = find_col(df, COL_TARGET)
    col_camp_name = normalize_col(df, COL_CAMP_NAME)
    col_stt       = normalize_col(df, COL_STT)

    if col_stt is None:
        df.insert(0, COL_STT, range(1, len(df) + 1))
    else:
        df[col_stt] = range(1, len(df) + 1)

    if col_camp_name is None:
        col_camp_name = COL_CAMP_NAME
        df[col_camp_name] = ""

    new_names = []
    for _, row in df.iterrows():
        note       = str(row[col_note]).strip()
        target_val = str(row[col_target]).strip()
        camp_name  = f"{sku}_{type_code}_{_parse_match_type(note)}_{target_val}_{_parse_tpr(note)}"
        new_names.append(camp_name)

    df[col_camp_name] = new_names
    return df


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def process_sku(xls: pd.ExcelFile, sku: str) -> pd.DataFrame | None:
    print(f"\n  --- Xử lý SKU: {sku} ---")
    try:
        df = step3_clean_and_crossjoin(xls, sku)
    except ValueError as e:
        print(f"  [WARN] {e}")
        return None

    if df.empty:
        print(f"  [WARN] Sheet '{sku}' không có dữ liệu Target hợp lệ.")
        return None

    df, type_code = step4_normalize_campaign_type(df)
    df = step5_naming_and_stt(df, sku, type_code)
    print(f"  [OK] SKU '{sku}' → {len(df)} dòng sau cross-join.")
    return df


def main():
    print("\n" + "=" * 60)
    print("  EXCEL NEW CAMP — Bước 1/2 Pipeline C_mass_sop_factory")
    print("=" * 60)

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"[ERROR] Không tìm thấy file input: {INPUT_FILE}\n"
            f"→ Đặt PPC_NGUYEN.xlsx vào: {INPUT_DIR}"
        )

    os.makedirs(INPUT_DIR, exist_ok=True)

    print(f"\n  Input  : {INPUT_FILE}")
    print(f"  Output : {OUTPUT_FILE}")

    xls = pd.ExcelFile(INPUT_FILE)
    available_sheets = xls.sheet_names
    print(f"  Sheets : {available_sheets}")

    sku_rows = step1_load_listing(xls)
    if not sku_rows:
        print("\n[DONE] Không có SKU nào cần xử lý.")
        return

    step2_validate_portfolio(xls, sku_rows)

    print("\n" + "=" * 60)
    print("[BƯỚC 3-5] Xử lý từng SKU (Clean → Cross-Join → Naming)")
    print("=" * 60)

    processed_results: dict[str, pd.DataFrame] = {}

    for item in sku_rows:
        sku = item["sku"]
        if sku not in available_sheets:
            print(f"\n  [WARN] SKU '{sku}' không có sheet tương ứng → Bỏ qua.")
            continue
        try:
            df_result = process_sku(xls, sku)
        except Exception as e:
            print(f"\n  [ERROR] SKU '{sku}': {e} → Bỏ qua.")
            continue
        if df_result is not None:
            processed_results[sku] = df_result

    if not processed_results:
        print("  [ERROR] Không có SKU nào được xử lý thành công.")
        return

    print("\n" + "=" * 60)
    print("[OUTPUT] Xuất PPC_PROCESSED_OUTPUT.xlsx")
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
                        worksheet.set_column(col_idx, col_idx, 25, text_fmt)
            print(f"\n  [OK] Đã lưu: {OUTPUT_FILE}")
            break
        except PermissionError:
            input(f"\n  [!] File đang mở. Đóng Excel rồi nhấn Enter...")

    print("\n" + "=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    total_rows = sum(len(df) for df in processed_results.values())
    print(f"  SKU thành công : {len(processed_results)}")
    print(f"  SKU bỏ qua     : {len(sku_rows) - len(processed_results)}")
    print(f"  Tổng dòng      : {total_rows}")
    print("=" * 60)
    print("\n  → Chạy tiếp: python mass_sop_factory.py\n")


if __name__ == "__main__":
    main()
