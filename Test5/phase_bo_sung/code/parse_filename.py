import os
import re
from datetime import datetime

def parse_bulk_filename(filename):
    """
    Phân tích tên file Bulk để trích xuất số ngày và khoảng thời gian.
    Ví dụ tên file:
    - "BulkSheetExport_17-20 (1).xlsx" -> Có thể là từ ngày 17 đến 20 của tháng hiện tại.
    - "Bulk_20260401_20260415.xlsx" -> Format chuẩn YYYYMMDD.
    - "Bulk Operations 04-20-2026 to 05-20-2026.xlsx" -> Format mặc định của Amazon.
    """
    filename_clean = os.path.basename(filename)
    result = {
        "filename": filename_clean,
        "start_date": None,
        "end_date": None,
        "days_duration": None,
        "note": ""
    }

    # Pattern 1: Format mặc định của Amazon "Bulk Operations MM-DD-YYYY to MM-DD-YYYY"
    match_amazon = re.search(r'(\d{2}-\d{2}-\d{4})\s*to\s*(\d{2}-\d{2}-\d{4})', filename_clean)
    if match_amazon:
        try:
            start_dt = datetime.strptime(match_amazon.group(1), "%m-%d-%Y")
            end_dt = datetime.strptime(match_amazon.group(2), "%m-%d-%Y")
            result["start_date"] = start_dt.strftime("%Y-%m-%d")
            result["end_date"] = end_dt.strftime("%Y-%m-%d")
            result["days_duration"] = (end_dt - start_dt).days + 1
            result["note"] = "Amazon Default Format"
            return result
        except Exception:
            pass

    # Pattern 2: Format ngắn gọn "BulkSheetExport_17-20.xlsx" (Từ ngày X đến ngày Y trong tháng)
    # Cần logic bổ sung vì thiếu tháng/năm. Tạm lấy tháng/năm hiện tại.
    match_short = re.search(r'_(\d{1,2})-(\d{1,2})', filename_clean)
    if match_short:
        try:
            start_day = int(match_short.group(1))
            end_day = int(match_short.group(2))
            
            # Lấy tháng năm hiện tại làm gốc
            now = datetime.now()
            start_dt = datetime(now.year, now.month, start_day)
            end_dt = datetime(now.year, now.month, end_day)
            
            # Xử lý trường hợp lùi ngày qua tháng trước (VD: 25 đến 05)
            if end_day < start_day:
                start_dt = datetime(now.year, now.month - 1 if now.month > 1 else 12, start_day)
                if now.month == 1:
                    start_dt = start_dt.replace(year=now.year - 1)

            result["start_date"] = start_dt.strftime("%Y-%m-%d")
            result["end_date"] = end_dt.strftime("%Y-%m-%d")
            result["days_duration"] = (end_dt - start_dt).days + 1
            result["note"] = "Short Format (Assumed current month)"
            return result
        except Exception:
            pass

    # Pattern 3: Format hệ thống dài "bulk-*-YYYYMMDD-YYYYMMDD-*"
    match_long = re.search(r'(\d{8})-(\d{8})', filename_clean)
    if match_long:
        try:
            start_dt = datetime.strptime(match_long.group(1), "%Y%m%d")
            end_dt = datetime.strptime(match_long.group(2), "%Y%m%d")
            result["start_date"] = start_dt.strftime("%Y-%m-%d")
            result["end_date"] = end_dt.strftime("%Y-%m-%d")
            result["days_duration"] = (end_dt - start_dt).days + 1
            result["note"] = "Long Format (YYYYMMDD)"
            return result
        except Exception:
            pass

    # Nếu không parse được
    result["note"] = "Unknown Format"
    return result

if __name__ == "__main__":
    import glob
    import shutil
    import json
    
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    INPUT_DIR = os.path.join(BASE_DIR, "input")
    
    # Thư mục đích ở phase0
    PHASE0_DIR = os.path.join(os.path.dirname(BASE_DIR), "phase0")
    PHASE0_INPUT_DIR = os.path.join(PHASE0_DIR, "input")
    os.makedirs(PHASE0_INPUT_DIR, exist_ok=True)
    
    print("--- CHUAN BI CHUYEN TIEP DATA SANG PHASE 0 ---")
    
    input_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    if not input_files:
        print("Khong tim thay file .xlsx nao trong thu muc input/")
    else:
        for f in input_files:
            res = parse_bulk_filename(f)
            print(f"\nPhan tich file: {res['filename']}")
            print(f"Khoang thoi gian: {res['days_duration']} ngay")
            
            # Copy file sang phase0 input
            dest_file = os.path.join(PHASE0_INPUT_DIR, res['filename'])
            shutil.copy2(f, dest_file)
            print(f"Da copy sang: {dest_file}")
            
            # Luu metadata json
            meta_file = os.path.join(PHASE0_INPUT_DIR, "phase0_input_meta.json")
            with open(meta_file, "w", encoding="utf-8") as json_file:
                json.dump(res, json_file, indent=4)
            print(f"Da luu metadata sang: {meta_file}")
