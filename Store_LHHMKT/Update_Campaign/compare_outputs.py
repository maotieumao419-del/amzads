"""
compare_outputs.py
==================
VALIDATION TOOL – Compare Master Output vs. Raw Input BulkSheet
This script performs verification on:
1. QUANTITY & COMPLETENESS: Checks if all standard campaigns from the input
   have been renamed and included, and complex ones quarantined.
2. STRUCTURAL INTEGRITY: Checks if the columns match the official Amazon template
   exactly in order and count, static columns have correct values, and ID columns are clean.
"""

import os
import sys
import glob
import pandas as pd

# Load shared utilities
from utils import clean_id, safe_str, ID_COLUMNS, find_latest_bulksheet, load_bulksheet

# Define colors/symbols for pretty printing
OK_MSG = "\033[92m✓ OK\033[0m"
WARN_MSG = "\033[93m⚠️ WARNING\033[0m"
FAIL_MSG = "\033[91m❌ FAILED\033[0m"
CRIT_MSG = "\033[41m\033[37m🔥 CRITICAL ERROR\033[0m"

# Official Amazon Template Column Order (30 columns)
TEMPLATE_COLUMNS = [
    "Product",
    "Entity",
    "Operation",
    "Campaign ID",
    "Ad Group ID",
    "Portfolio ID",
    "Ad ID",
    "Keyword ID",
    "Product Targeting ID",
    "Campaign Name",
    "Ad Group Name",
    "Start Date",
    "End Date",
    "Targeting Type",
    "State",
    "Daily Budget",
    "SKU",
    "Ad Group Default Bid",
    "Bid",
    "Keyword Text",
    "Native Language Keyword",
    "Native Language Locale",
    "Match Type",
    "Bidding Strategy",
    "Placement",
    "Percentage",
    "Product Targeting Expression",
    "Audience ID",
    "Shopper Cohort Percentage",
    "Shopper Cohort Type",
]

def print_banner(title: str):
    print()
    print("=" * 80)
    print(f" {title.center(78)} ")
    print("=" * 80)

def classify_campaign(rows: list[dict]) -> tuple[str, str]:
    """
    Classify a campaign as standard or complex.
    Matches the exact logic of module1_isolation.py.
    """
    product_ad_rows = [r for r in rows if safe_str(r.get("Entity")).lower() == "product ad"]
    keyword_rows = [r for r in rows if safe_str(r.get("Entity")).lower() == "keyword"]
    pt_rows = [r for r in rows if safe_str(r.get("Entity")).lower() == "product targeting"]

    skus = {safe_str(r.get("SKU")) for r in product_ad_rows if safe_str(r.get("SKU"))}

    if len(skus) > 1:
        return f"Multiple SKUs ({len(skus)}): " + ", ".join(sorted(skus)), ""

    has_kw = len(keyword_rows) > 0
    has_pt = len(pt_rows) > 0
    if has_kw and has_pt:
        return f"Mixed targeting: {len(keyword_rows)} Keyword row(s) and {len(pt_rows)} Product Targeting row(s)", ""

    sku = next(iter(skus)) if skus else "UNKNOWN_SKU"
    return "", sku

def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
    # ── 1. Locate Files ──────────────────────────────────────────────────────────
    print_banner("1. ĐỊNH VỊ CÁC FILE ĐẦU VÀO VÀ ĐẦU RA")
    
    try:
        input_path = find_latest_bulksheet(BASE_DIR)
        print(f"  {OK_MSG} Tìm thấy file INPUT (mới nhất): {os.path.basename(input_path)}")
    except FileNotFoundError as exc:
        print(f"  {FAIL_MSG} Không tìm thấy file BulkSheet đầu vào trong data/input/ hoặc data/!")
        print(exc)
        sys.exit(1)

    master_dir = os.path.join(BASE_DIR, "data", "master")
    if not os.path.exists(master_dir):
        master_dir = os.path.join(BASE_DIR, "data", "upload_files")
        
    candidates = glob.glob(os.path.join(master_dir, "*_UpdateCampaigns_*.xlsx"))
    if not candidates:
        candidates = glob.glob(os.path.join(master_dir, "Master_Amazon_Bulksheet_Update_Upload.xlsx"))
        
    if candidates:
        master_xlsx_path = sorted(candidates, key=os.path.getmtime, reverse=True)[0]
        print(f"  {OK_MSG} Tìm thấy file OUTPUT (Master): {os.path.abspath(master_xlsx_path)}")
    else:
        print(f"  {FAIL_MSG} Không tìm thấy file output Master tại `data/master/`!")
        sys.exit(1)

    # ── 2. Read Files ────────────────────────────────────────────────────────────
    print_banner("2. ĐỌC DỮ LIỆU VÀ PHÂN LOẠI CAMPAIGNS TỪ FILE INPUT")
    
    df_in = load_bulksheet(input_path)
    df_out = load_bulksheet(master_xlsx_path)

    # Resolve campaign name column in input file
    name_candidates = ["Campaign Name (Informational only)", "Campaign Name", "Name", "Campaign"]
    in_name_col = next((col for col in name_candidates if col in df_in.columns), None)
    if not in_name_col:
        print(f"  {FAIL_MSG} Không tìm thấy cột chứa tên chiến dịch trong file đầu vào!")
        sys.exit(1)

    # Group input rows by Campaign ID
    by_campaign_in = {}
    for _, row in df_in.iterrows():
        cid = clean_id(row.get("Campaign ID"))
        if cid:
            by_campaign_in.setdefault(cid, []).append(row.to_dict())

    total_in_cids = len(by_campaign_in)
    print(f"  - Tổng số chiến dịch (Campaign IDs) tìm thấy trong file input: {total_in_cids:,}")

    # Classify input campaigns
    std_campaigns = {}  # cid -> {sku, old_name, portfolio_id}
    complex_campaigns = {}  # cid -> {reason, old_name}

    for cid, rows in by_campaign_in.items():
        # Get campaign name from Campaign row
        camp_row = next((r for r in rows if safe_str(r.get("Entity")).lower() == "campaign"), None)
        old_name = safe_str(camp_row.get(in_name_col)) if camp_row else safe_str(rows[0].get(in_name_col))
        portfolio_id = clean_id(camp_row.get("Portfolio ID")) if camp_row else clean_id(rows[0].get("Portfolio ID"))
        
        reason, sku = classify_campaign(rows)
        if reason:
            complex_campaigns[cid] = {"reason": reason, "old_name": old_name}
        else:
            std_campaigns[cid] = {"sku": sku, "old_name": old_name, "portfolio_id": portfolio_id}

    print(f"  - Số chiến dịch TIÊU CHUẨN (Standard - Cần đổi tên): {len(std_campaigns):,}")
    print(f"  - Số chiến dịch PHỨC TẠP (Complex - Cần review tay, giữ nguyên): {len(complex_campaigns):,}")

    # Group output rows
    df_out_campaigns = df_out[df_out["Entity"].str.strip().str.lower() == "campaign"].copy()
    out_cids = set(df_out_campaigns["Campaign ID"].apply(clean_id))
    print(f"  - Tổng số chiến dịch được ghi nhận trong file output Master: {len(out_cids):,}")

    # ── 3. Quantity & Completeness Verification ──────────────────────────────────
    print_banner("3. KIỂM TRA SỐ LƯỢNG VÀ TÍNH ĐẦY ĐỦ CỦA DANH SÁCH ĐỔI TÊN")
    
    warnings_count = 0
    errors_count = 0

    # 3.1. Standard campaigns missing from Output?
    missing_std = set(std_campaigns.keys()) - out_cids
    if missing_std:
        print(f"  {FAIL_MSG} Có {len(missing_std)} chiến dịch tiêu chuẩn BỊ THIẾU trong file output:")
        errors_count += len(missing_std)
        for idx, cid in enumerate(sorted(missing_std), start=1):
            info = std_campaigns[cid]
            print(f"    {idx}. ID: {cid} | SKU: {info['sku']} | Tên cũ: {info['old_name']}")
    else:
        print(f"  {OK_MSG} Tất cả chiến dịch tiêu chuẩn đã được cập nhật vào file output.")

    # 3.2. Complex campaigns leaked into Output?
    leaked_complex = set(complex_campaigns.keys()) & out_cids
    if leaked_complex:
        print(f"  {FAIL_MSG} Phát hiện {len(leaked_complex)} chiến dịch phức tạp bị ghi đè vào file output (Không đúng thiết kế):")
        errors_count += len(leaked_complex)
        for idx, cid in enumerate(sorted(leaked_complex), start=1):
            info = complex_campaigns[cid]
            print(f"    {idx}. ID: {cid} | Tên cũ: {info['old_name']} | Lý do: {info['reason']}")
    else:
        print(f"  {OK_MSG} Không có chiến dịch phức tạp nào bị ghi nhầm vào file output.")

    # 3.3. Unrecognized campaigns in Output?
    unknown_out = out_cids - set(by_campaign_in.keys())
    if unknown_out:
        print(f"  {FAIL_MSG} Phát hiện {len(unknown_out)} Campaign ID trong file output KHÔNG tồn tại trong file input:")
        errors_count += len(unknown_out)
        for idx, cid in enumerate(sorted(unknown_out), start=1):
            print(f"    {idx}. ID: {cid}")
    else:
        print(f"  {OK_MSG} Không có Campaign ID lạ nào xuất hiện trong file output.")

    # ── 4. Structural Verification ───────────────────────────────────────────────
    print_banner("4. KIỂM TRA CẤU TRÚC FILE OUTPUT THEO TEMPLATE AMAZON")
    
    # 4.1. Sheet Name Check
    xls_out = pd.ExcelFile(master_xlsx_path, engine="openpyxl")
    expected_sheet = "Sponsored Products Campaigns"
    if expected_sheet not in xls_out.sheet_names:
        print(f"  {FAIL_MSG} Tên Sheet không chính xác! Yêu cầu: '{expected_sheet}', thực tế: '{xls_out.sheet_names[0]}'")
        errors_count += 1
    else:
        print(f"  {OK_MSG} Tên Sheet chính xác: '{expected_sheet}'")

    # 4.2. Column Count and Order Check
    out_cols = list(df_out.columns)
    if len(out_cols) != len(TEMPLATE_COLUMNS):
        print(f"  {FAIL_MSG} Số lượng cột không khớp! Yêu cầu: {len(TEMPLATE_COLUMNS)} cột, thực tế: {len(out_cols)} cột")
        errors_count += 1
    else:
        print(f"  {OK_MSG} Đủ số lượng cột ({len(TEMPLATE_COLUMNS)} cột)")

    mismatched_cols = []
    for idx, (expected, actual) in enumerate(zip(TEMPLATE_COLUMNS, out_cols)):
        if expected != actual:
            mismatched_cols.append((idx + 1, expected, actual))

    if mismatched_cols:
        print(f"  {FAIL_MSG} Phát hiện {len(mismatched_cols)} cột sai vị trí hoặc sai tên:")
        errors_count += len(mismatched_cols)
        for idx_col, expected_name, actual_name in mismatched_cols:
            print(f"    - Cột số {idx_col}: Yêu cầu '{expected_name}', thực tế là '{actual_name}'")
    else:
        print(f"  {OK_MSG} Thứ tự và tên các cột khớp 100% với Amazon template chuẩn.")

    # 4.3. Static Fields Verification (Product, Entity, Operation, State)
    static_errors = 0
    row_idx = 2  # Starting row in Excel (Header is row 1)
    
    for _, row in df_out.iterrows():
        p = safe_str(row.get("Product"))
        e = safe_str(row.get("Entity"))
        o = safe_str(row.get("Operation"))
        s = safe_str(row.get("State"))

        if p != "Sponsored Products":
            print(f"    - Dòng {row_idx}: Cột 'Product' sai! Có giá trị '{p}' (yêu cầu: 'Sponsored Products')")
            static_errors += 1
        if e.lower() != "campaign":
            print(f"    - Dòng {row_idx}: Cột 'Entity' sai! Có giá trị '{e}' (yêu cầu: 'Campaign')")
            static_errors += 1
        if o != "Update":
            print(f"    - Dòng {row_idx}: Cột 'Operation' sai! Có giá trị '{o}' (yêu cầu: 'Update')")
            static_errors += 1
        if s != "enabled":
            print(f"    - Dòng {row_idx}: Cột 'State' sai! Có giá trị '{s}' (yêu cầu: 'enabled')")
            static_errors += 1
        row_idx += 1

    if static_errors:
        print(f"  {FAIL_MSG} Phát hiện {static_errors} lỗi tại các trường dữ liệu tĩnh (Product, Entity, Operation, State)")
        errors_count += static_errors
    else:
        print(f"  {OK_MSG} Các trường dữ liệu tĩnh (Product, Entity, Operation, State) đã được set đúng giá trị cho mọi dòng.")

    # 4.4. Clean IDs Verification
    id_errors = 0
    row_idx = 2
    for _, row in df_out.iterrows():
        cid_raw = str(row.get("Campaign ID", ""))
        port_raw = str(row.get("Portfolio ID", ""))
        
        # Check Campaign ID format
        if not cid_raw or cid_raw.lower() == "nan":
            print(f"    - Dòng {row_idx}: 'Campaign ID' bị rỗng hoặc NaN!")
            id_errors += 1
        elif ".0" in cid_raw or "e+" in cid_raw.lower():
            print(f"    - Dòng {row_idx}: 'Campaign ID' bị lỗi format float/scientific: '{cid_raw}'")
            id_errors += 1

        # Check Portfolio ID format
        if port_raw and port_raw.lower() != "nan":
            if ".0" in port_raw or "e+" in port_raw.lower():
                print(f"    - Dòng {row_idx}: 'Portfolio ID' bị lỗi format float/scientific: '{port_raw}'")
                id_errors += 1
        row_idx += 1

    if id_errors:
        print(f"  {FAIL_MSG} Phát hiện {id_errors} lỗi về định dạng ID (bị đuôi .0 hoặc dạng e+17)")
        errors_count += id_errors
    else:
        print(f"  {OK_MSG} Tất cả Campaign ID và Portfolio ID đều sạch (không bị lỗi float / sci-notation).")

    # ── 5. Data Integrity & Naming SOP Verification ─────────────────────────────
    print_banner("5. KIỂM TRA TÍNH TOÀN VẸN DỮ LIỆU & QUY TẮC ĐẶT TÊN SOP")
    
    name_errors = 0
    portfolio_unlinked = 0
    
    for idx, row in df_out_campaigns.iterrows():
        cid = clean_id(row.get("Campaign ID"))
        new_name = safe_str(row.get("Campaign Name"))
        port_out = clean_id(row.get("Portfolio ID"))
        
        # Verify Campaign ID exists in std_campaigns
        if cid in std_campaigns:
            old_name = std_campaigns[cid]["old_name"]
            port_in = std_campaigns[cid]["portfolio_id"]
            sku = std_campaigns[cid]["sku"]
            
            # Check name changed
            if old_name == new_name:
                print(f"    [CẢNH BÁO] Campaign {cid} (SKU: {sku}) không thay đổi tên: '{old_name}'")
                warnings_count += 1
                
            # Check double underscores
            if "__" in new_name:
                print(f"    [CẢNH BÁO] Campaign {cid} có chứa dấu gạch dưới kép '__': '{new_name}'")
                warnings_count += 1
                
            # Check portfolio unlinked
            if port_in and not port_out:
                print(f"    {CRIT_MSG} Campaign {cid} bị mất Portfolio ID! (Trong input: {port_in}, trong output: Rỗng)")
                portfolio_unlinked += 1

    if portfolio_unlinked:
        print(f"  {FAIL_MSG} Phát hiện {portfolio_unlinked} lỗi nghiêm trọng làm ngắt kết nối Portfolio ID!")
        errors_count += portfolio_unlinked
    else:
        print(f"  {OK_MSG} Toàn bộ liên kết Portfolio ID được giữ nguyên vẹn.")

    # ── 6. Verification Summary ──────────────────────────────────────────────────
    print_banner("6. TỔNG KẾT KẾT QUẢ KIỂM TRA (SUMMARY)")
    
    print(f"  - File Bulk đầu vào      : {os.path.basename(input_path)}")
    print(f"  - File Master đầu ra     : {os.path.basename(master_xlsx_path)}")
    print(f"  - Tổng số Standard CIDs  : {len(std_campaigns)}")
    print(f"  - Số CIDs trong Output   : {len(out_cids)}")
    print()
    print(f"  - Số CẢNH BÁO (Warnings)  : {warnings_count}")
    print(f"  - Số LỖI (Errors)         : {errors_count}")
    print()

    # Show a few naming examples
    print("  Một số ví dụ đổi tên tiêu biểu trong file output:")
    sample_cids = list(out_cids & std_campaigns.keys())[:5]
    for cid in sample_cids:
        old = std_campaigns[cid]["old_name"]
        new_row = df_out_campaigns[df_out_campaigns["Campaign ID"].apply(clean_id) == cid].iloc[0]
        new = new_row.get("Campaign Name")
        print(f"    • ID: {cid}")
        print(f"      - Cũ: {old}")
        print(f"      - Mới: {new}")
        
    print()
    if errors_count == 0:
        print(f"  \033[92m=> KẾT LUẬN: FILE CHUẨN TEMPLATE VÀ KHỚP SỐ LƯỢNG. KHÔNG CÓ LỖI. SẴN SÀNG UPLOAD! \033[0m")
    else:
        print(f"  \033[91m=> KẾT LUẬN: CÓ {errors_count} LỖI CẦN ĐƯỢC XỬ LÝ TRƯỚC KHI UPLOAD LÊN AMAZON ADS! \033[0m")
    print("=" * 80)
    print()

if __name__ == "__main__":
    main()
