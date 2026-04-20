# =============================================================================
# FILE: amz_schema_validator.py
# MỤC ĐÍCH: Quality Control (QC) - Đối chiếu cấu trúc cột của các file Bulk_*
#           với file template chuẩn.
#
# THUẬT TOÁN:
#   1. Tìm file template chuẩn (mass_sop_part_1.xlsx) và đọc header thành Set.
#   2. Quét các file Bulk_*.xlsx hoặc Bulk_*.csv.
#   3. Đọc header của từng file target và đối chiếu bằng phép toán Set.
#   4. In kết quả PASS/FAILED chi tiết và ghi ra file log.
# =============================================================================

import os
import glob
import pandas as pd

# ---------------------------------------------------------------------------
# CẤU HÌNH ĐƯỜNG DẪN & TÊN FILE
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Xác định thư mục chứa file. Vì mass_sop_factory.py lưu output vào data/output
# ta sẽ kiểm tra cả thư mục script và thư mục output.
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "data", "output")

# Tên file template chuẩn
TEMPLATE_FILENAME = "mass_sop_part_1.xlsx"

# File log xuất ra
LOG_FILENAME = "validation_report.txt"

def get_header_columns(file_path: str) -> set:
    """
    Đọc file (XLSX hoặc CSV) và chỉ lấy dòng header (tên cột) để tối ưu RAM.
    Trả về một Set chứa các tên cột (đã loại bỏ khoảng trắng thừa).
    """
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.csv':
            # nrows=0 chỉ đọc đúng dòng header với CSV
            df = pd.read_csv(file_path, nrows=0)
        elif ext in ['.xlsx', '.xls']:
            # nrows=0 hỗ trợ lấy header đối với Excel trong pandas
            df = pd.read_excel(file_path, nrows=0)
        else:
            print(f"[ERROR] Định dạng không được hỗ trợ: {file_path}")
            return set()
        
        # Chuyển tên cột thành string, strip khoảng trắng và đưa vào Set
        return set(str(col).strip() for col in df.columns)
    except Exception as e:
        print(f"[ERROR] Không đọc được header của file {file_path}: {e}")
        return set()

def find_file_in_dirs(filename: str, dirs_to_search: list) -> str:
    """Tìm một file trong danh sách các thư mục, trả về đường dẫn tuyệt đối nếu có."""
    for d in dirs_to_search:
        path = os.path.join(d, filename)
        if os.path.exists(path):
            return path
    return ""

def main():
    print("=== AMZ BULKSHEET SCHEMA VALIDATOR ===")
    
    # 1. Tìm file template chuẩn (thử tìm ở SCRIPT_DIR, sau đó là OUTPUT_DIR)
    search_dirs = [SCRIPT_DIR, OUTPUT_DIR]
    template_path = find_file_in_dirs(TEMPLATE_FILENAME, search_dirs)
    
    if not template_path:
        print(f"[LỖI] Không tìm thấy file template chuẩn: '{TEMPLATE_FILENAME}'")
        print(f"Đã tìm trong: {search_dirs}")
        return

    # Lấy tập hợp cột của template
    template_cols = get_header_columns(template_path)
    if not template_cols:
        print("[LỖI] Template không có cột nào hoặc lỗi khi đọc template.")
        return

    print(f"Baseline Template: {TEMPLATE_FILENAME} loaded. Total columns: {len(template_cols)}\n")

    # 2. Tìm tất cả các file target có tiền tố Bulk_
    # Quét cả ở trong SCRIPT_DIR và OUTPUT_DIR để đảm bảo không bỏ sót
    target_files = []
    for d in search_dirs:
        if os.path.exists(d):
            # Tìm *.xlsx và *.csv
            target_files.extend(glob.glob(os.path.join(d, "Bulk_*.xlsx")))
            target_files.extend(glob.glob(os.path.join(d, "Bulk_*.csv")))
    
    # Xóa trùng lặp nếu có (dựa theo đường dẫn)
    target_files = list(set(target_files))
    
    if not target_files:
        print("[ẢNH BÁO] Không tìm thấy file output nào có tiền tố 'Bulk_' để đối chiếu.")
        return

    # Sắp xếp để log dễ nhìn hơn
    target_files.sort()

    log_lines = []
    log_lines.append("=== AMZ BULKSHEET VALIDATION REPORT ===")
    log_lines.append(f"Baseline Template: {TEMPLATE_FILENAME} (Total columns: {len(template_cols)})")
    log_lines.append("-" * 40)

    # 3. Thuật toán Đối chiếu cho từng file
    pass_count = 0
    fail_count = 0

    for file_path in target_files:
        filename = os.path.basename(file_path)
        print(f"Analyzing file: {filename}")
        
        target_cols = get_header_columns(file_path)
        
        if not target_cols:
            status_msg = f"-> Status: [FAILED]\n   - Lỗi: Không thể đọc file hoặc file không có cột."
            print(status_msg + "\n")
            log_lines.append(f"File: {filename}\n{status_msg}\n")
            fail_count += 1
            continue

        # Dùng phép toán Set để tìm Missing và Extra columns
        missing_columns = template_cols - target_cols
        extra_columns = target_cols - template_cols

        if not missing_columns and not extra_columns:
            # Khớp 100%
            status_msg = "-> Status: [PASS] - 100% Match"
            print(status_msg + "\n")
            log_lines.append(f"File: {filename}\n{status_msg}\n")
            pass_count += 1
        else:
            # Có lệch cấu trúc
            status_msg = "-> Status: [FAILED]"
            print(status_msg)
            
            missing_text = f"   - Missing Columns: {list(missing_columns)}" if missing_columns else ""
            extra_text   = f"   - Extra Columns: {list(extra_columns)}" if extra_columns else ""
            
            if missing_text: print(missing_text)
            if extra_text:   print(extra_text)
            print("") # New line
            
            log_lines.append(f"File: {filename}\n{status_msg}")
            if missing_text: log_lines.append(missing_text)
            if extra_text:   log_lines.append(extra_text)
            log_lines.append("")
            
            fail_count += 1

    # In báo cáo tổng quan ra terminal
    print("========================================")
    print(f"Validation Summary: {pass_count} Passed | {fail_count} Failed")
    print("========================================")

    # 4. Lưu ra file txt
    log_lines.append("-" * 40)
    log_lines.append(f"Validation Summary: {pass_count} Passed | {fail_count} Failed")
    
    log_path = os.path.join(SCRIPT_DIR, LOG_FILENAME)
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines))
        print(f"\nReport saved to: {log_path}")
    except Exception as e:
        print(f"\n[LỖI] Không thể lưu file log: {e}")

if __name__ == "__main__":
    main()
