import os
import glob
import re
import pandas as pd
import numpy as np

# ==============================================================================
# CẤU HÌNH ĐƯỜNG DẪN FILE (USER CÓ THỂ TÙY CHỈNH THEO THỰC TẾ)
# ==============================================================================
# Đường dẫn tới file Master xuất từ Amazon (Sponsored Products Campaigns.csv)
AMAZON_MASTER_FILE = r"BulkSheetExport_Sponsored_Products_Campaigns.csv"

# Thư mục chứa các file tracking nội bộ (tên bắt đầu bằng "PPC_Musemory_UPDATED.xlsx - ... .csv")
# Để dấu "." mặc định quét trong thư mục hiện tại
INTERNAL_CSV_DIR = "."

# Tên file Excel báo cáo đầu ra
OUTPUT_EXCEL_FILE = "Placement_CrossCheck_Report.xlsx"


def clean_percentage(val):
    """
    Làm sạch giá trị Percentage từ file Amazon.
    Xử lý mượt mà các trường hợp NaN, chuỗi rỗng, có ký tự % hoặc chuỗi không hợp lệ.
    """
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip().replace('%', '')
    if not val_str or val_str.lower() == 'nan':
        return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0


def extract_ppc_placements(note_str):
    """
    Hàm Regex bóc tách chính xác giá trị Top (T) và Rest of Search (R) từ cột Ghi chú.
    Format kỳ vọng: [[Match Type]-[Bid]-[TopT-RestR-ProductP]]
    Ví dụ: 
      - "[[Exact]-[0.56]-[20T-20R-20P]]" => T = 20.0, R = 20.0
      - "[[Broad]-[0.35]-[0T-0R-0P]]"    => T = 0.0,  R = 0.0
    """
    if not isinstance(note_str, str) or not note_str.strip():
        return 0.0, 0.0
    
    # 1. Tìm cụm liền kề có định dạng chuẩn (VD: 20T-20R)
    # Hỗ trợ số thập phân bằng (?:\.\d+)? và không phân biệt hoa thường (re.IGNORECASE)
    match = re.search(r'(\d+(?:\.\d+)?)T\s*-\s*(\d+(?:\.\d+)?)R', note_str, re.IGNORECASE)
    if match:
        return float(match.group(1)), float(match.group(2))
    
    # 2. Phân rã tìm độc lập T và R trong trường hợp chuỗi bị gõ thiếu/sai dấu phân cách
    match_t = re.search(r'(\d+(?:\.\d+)?)T', note_str, re.IGNORECASE)
    match_r = re.search(r'(\d+(?:\.\d+)?)R', note_str, re.IGNORECASE)
    
    t_val = float(match_t.group(1)) if match_t else 0.0
    r_val = float(match_r.group(1)) if match_r else 0.0
    
    return t_val, r_val


def determine_status(top_ppc, top_amz, rest_ppc, rest_amz):
    """
    Xác định trạng thái so sánh (Status) dựa trên logic chéo giữa PPC và Amazon.
    """
    # Khớp hoàn toàn (chấp nhận sai số nhỏ của phép tính số thập phân)
    if np.isclose(top_ppc, top_amz, atol=0.01) and np.isclose(rest_ppc, rest_amz, atol=0.01):
        return "MATCH"
    
    # Theo chuẩn đề bài yêu cầu:
    # "MISMATCH - LỖI CODE": Nếu Top_Amazon > 0 nhưng Top_PPC = 0 (file master có nhưng trích xuất sai)
    if top_amz > 0 and top_ppc == 0:
        return "MISMATCH - LỖI CODE"
        
    # "MISMATCH - LỖI AMAZON/SYNC": Nếu Top_PPC > 0 nhưng Top_Amazon = 0 (nội bộ có nhưng Amazon bị mất/chưa set)
    if top_ppc > 0 and top_amz == 0:
        return "MISMATCH - LỖI AMAZON/SYNC"
        
    # Bổ sung kiểm tra chi tiết cho Rest of Search hoặc trường hợp bị lệch số liệu (VD: 20 vs 50)
    if rest_amz > 0 and rest_ppc == 0:
        return "MISMATCH - LỖI CODE (Rest)"
    if rest_ppc > 0 and rest_amz == 0:
        return "MISMATCH - LỖI AMAZON/SYNC (Rest)"
        
    return "MISMATCH - LỆCH GIÁ TRỊ"


def cross_check_placements():
    print("==========================================================")
    print(" SCRIPT KIỂM TRA CHÉO (CROSS-CHECK) PLACEMENT PERCENTAGE  ")
    print("==========================================================")
    
    # --- Tự động định vị file Amazon Master nếu cấu hình chưa chính xác ---
    master_file = AMAZON_MASTER_FILE
    if not os.path.exists(master_file):
        print(f"[Thông báo] Không tìm thấy file tại đường dẫn tĩnh '{master_file}'.")
        print("            Đang tự động quét tìm file BulkSheetExport CSV...")
        # Tìm file kết thúc bằng Sponsored Products Campaigns.csv
        candidates = glob.glob("**/*Sponsored Products Campaigns.csv", recursive=True)
        if not candidates:
            candidates = glob.glob("**/BulkSheetExport*.csv", recursive=True)
        
        if candidates:
            master_file = candidates[0]
            print(f"[Thành công] Đã tự động tìm thấy file Master: '{master_file}'")
        else:
            print("[Lỗi] Không tìm thấy file Amazon Master CSV nào trong hệ thống.")
            print("      Vui lòng cập nhật chính xác đường dẫn biến AMAZON_MASTER_FILE.")
            return

    # --------------------------------------------------------------------------
    # BƯỚC 1: XỬ LÝ FILE AMAZON MASTER (BulkSheetExport)
    # --------------------------------------------------------------------------
    print("\n--- BƯỚC 1: ĐỌC VÀ MAP DỮ LIỆU FILE AMAZON MASTER ---")
    try:
        df_amz = pd.read_csv(master_file, low_memory=False)
    except Exception as e:
        print(f"[Lỗi] Không thể đọc file Amazon Master: {e}")
        return

    # Xóa khoảng trắng thừa ở tên cột
    df_amz.columns = df_amz.columns.str.strip()
    
    # Kiểm tra các cột thiết yếu
    required_cols = ['Entity', 'Campaign Name', 'Placement', 'Percentage']
    missing_cols = [c for c in required_cols if c not in df_amz.columns]
    if missing_cols:
        print(f"[Lỗi] File Master bị thiếu các cột sau: {missing_cols}")
        return

    # Lọc các dòng Bidding Adjustment (không phân biệt hoa thường)
    df_amz['Entity_clean'] = df_amz['Entity'].astype(str).str.strip().str.lower()
    bidding_adj_df = df_amz[df_amz['Entity_clean'] == 'bidding adjustment']
    
    # Dictionary lưu trữ cấu hình Placement của từng chiến dịch
    # Format: { 'Campaign_A': {'Top': 20.0, 'Rest': 0.0} }
    amazon_mapping = {}
    
    # Khởi tạo mặc định bằng 0 cho tất cả Campaign Name xuất hiện trong file
    # để đảm bảo các chiến dịch không có dòng Bidding Adjustment vẫn hiểu là 0
    all_campaigns = df_amz['Campaign Name'].dropna().unique()
    for camp in all_campaigns:
        camp_clean = str(camp).strip()
        if camp_clean and camp_clean.lower() != 'nan':
            amazon_mapping[camp_clean] = {'Top': 0.0, 'Rest': 0.0}

    # Lặp qua các dòng Bidding Adjustment để ghi nhận phần trăm
    for _, row in bidding_adj_df.iterrows():
        camp = str(row['Campaign Name']).strip()
        if not camp or camp.lower() == 'nan':
            continue
            
        placement_type = str(row['Placement']).strip().lower()
        pct_val = clean_percentage(row['Percentage'])
        
        if camp not in amazon_mapping:
            amazon_mapping[camp] = {'Top': 0.0, 'Rest': 0.0}
            
        if 'top' in placement_type:
            amazon_mapping[camp]['Top'] = pct_val
        elif 'rest of search' in placement_type or 'rest_of_search' in placement_type:
            amazon_mapping[camp]['Rest'] = pct_val

    print(f"[Thành công] Đã tạo bảng mapping Placement cho {len(amazon_mapping)} chiến dịch từ Amazon.")

    # --------------------------------------------------------------------------
    # BƯỚC 2: XỬ LÝ CÁC FILE PPC NỘI BỘ (PPC_Musemory_UPDATED)
    # --------------------------------------------------------------------------
    print("\n--- BƯỚC 2: QUÉT VÀ TRÍCH XUẤT CÁC FILE CSV NỘI BỘ ---")
    # Quét tìm các file CSV bắt đầu bằng "PPC_Musemory_UPDATED.xlsx - "
    search_pattern = os.path.join(INTERNAL_CSV_DIR, "PPC_Musemory_UPDATED.xlsx - *.csv")
    internal_files = glob.glob(search_pattern)
    
    if not internal_files:
        # Thử tìm đệ quy nếu thư mục hiện tại không thấy
        internal_files = glob.glob("**/PPC_Musemory_UPDATED.xlsx - *.csv", recursive=True)
        
    if not internal_files:
        print("[Cảnh báo] Không tìm thấy file CSV nội bộ nào khớp mẫu 'PPC_Musemory_UPDATED.xlsx - *.csv'.")
        print("           Đảm bảo bạn đã trích xuất/lưu các sheet Excel ra định dạng CSV tương ứng.")
        return
        
    internal_records = []
    exclude_keywords = ['Listing.csv', 'Portfolio ID.csv']
    
    files_processed = 0
    for file_path in internal_files:
        file_name = os.path.basename(file_path)
        
        # Bỏ qua các file cấu hình chung theo yêu cầu
        if any(file_name.endswith(ex) for ex in exclude_keywords):
            continue
            
        # Lấy tên file source (tên sheet nội bộ)
        source_sheet = file_name.replace("PPC_Musemory_UPDATED.xlsx - ", "").replace(".csv", "").strip()
        
        try:
            df_int = pd.read_csv(file_path, low_memory=False)
            df_int.columns = df_int.columns.str.strip()
            
            # Linh hoạt tìm tên cột Ghi chú (tránh lỗi font/dấu)
            note_col = None
            for col in ['Ghi chú', 'Ghi chu', 'Ghi ch\u00fa']:
                if col in df_int.columns:
                    note_col = col
                    break
                    
            if 'Campaign Name' not in df_int.columns:
                continue
                
            # Xử lý cột Ghi chú bị thiếu hoặc NaN
            if not note_col:
                df_int['Ghi_chú_clean'] = ""
            else:
                df_int['Ghi_chú_clean'] = df_int[note_col].fillna("")
                
            # Duyệt các dòng chiến dịch
            for _, row in df_int.iterrows():
                camp = str(row['Campaign Name']).strip()
                if not camp or camp.lower() == 'nan':
                    continue
                    
                note_str = str(row['Ghi_chú_clean']).strip()
                top_ppc, rest_ppc = extract_ppc_placements(note_str)
                
                internal_records.append({
                    'File Source': source_sheet,
                    'Campaign Name': camp,
                    'Chuỗi Ghi Chú Gốc': note_str,
                    'Top_PPC': top_ppc,
                    'Rest_PPC': rest_ppc
                })
            files_processed += 1
            
        except Exception as e:
            print(f"[Lỗi] Không thể đọc file nội bộ '{file_name}': {e}")
            
    if not internal_records:
        print("[Lỗi] Không thu thập được dòng dữ liệu hợp lệ nào từ các file tracking nội bộ.")
        return
        
    df_internal_all = pd.DataFrame(internal_records)
    
    # TỐI ƯU HÓA: Loại bỏ các dòng trùng lặp hoàn toàn theo (File Source, Campaign Name, Ghi chú gốc)
    # Giúp báo cáo gọn gàng, tránh lặp lại 50 dòng giống hệt nhau nếu 1 chiến dịch chứa 50 từ khóa
    df_internal_unique = df_internal_all.drop_duplicates(subset=['File Source', 'Campaign Name', 'Chuỗi Ghi Chú Gốc']).copy()
    
    print(f"[Thành công] Đã quét {files_processed} file CSV hợp lệ.")
    print(f"            Trích xuất được {len(df_internal_unique)} cấu hình chiến dịch duy nhất.")

    # --------------------------------------------------------------------------
    # BƯỚC 3 & 4: SO SÁNH CHÉO (CROSS-CHECK) VÀ XUẤT OUTPUT
    # --------------------------------------------------------------------------
    print("\n--- BƯỚC 3 & 4: SO SÁNH CHÉO VÀ KẾT XUẤT BÁO CÁO EXCEL ---")
    
    # Map dữ liệu Top và Rest từ Master file sang
    df_internal_unique['Top_Amazon'] = df_internal_unique['Campaign Name'].apply(
        lambda c: amazon_mapping.get(c, {}).get('Top', 0.0)
    )
    df_internal_unique['Rest_Amazon'] = df_internal_unique['Campaign Name'].apply(
        lambda c: amazon_mapping.get(c, {}).get('Rest', 0.0)
    )
    
    # Áp dụng logic so sánh để xác định Status
    status_list = []
    for _, row in df_internal_unique.iterrows():
        st = determine_status(row['Top_PPC'], row['Top_Amazon'], row['Rest_PPC'], row['Rest_Amazon'])
        status_list.append(st)
        
    df_internal_unique['Status'] = status_list
    
    # Sắp xếp đúng thứ tự cột Output theo yêu cầu
    output_columns = [
        'File Source',
        'Campaign Name',
        'Chuỗi Ghi Chú Gốc',
        'Top_PPC',
        'Top_Amazon',
        'Rest_PPC',
        'Rest_Amazon',
        'Status'
    ]
    df_final_report = df_internal_unique[output_columns].copy()
    
    # Sắp xếp các dòng để dễ theo dõi (Đẩy các dòng MISMATCH lên trên cùng)
    df_final_report.sort_values(
        by=['Status', 'File Source', 'Campaign Name'], 
        ascending=[False, True, True], 
        inplace=True
    )
    
    # Ghi file Excel
    try:
        df_final_report.to_excel(OUTPUT_EXCEL_FILE, index=False, engine='openpyxl')
        print(f"\n[HOÀN THÀNH] Báo cáo đã được xuất thành công ra file:")
        print(f"             → os.path.abspath(OUTPUT_EXCEL_FILE): {os.path.abspath(OUTPUT_EXCEL_FILE)}")
        
        # Thống kê tổng quan kết quả
        match_count = (df_final_report['Status'] == 'MATCH').sum()
        err_code    = df_final_report['Status'].str.contains('LỖI CODE').sum()
        err_sync    = df_final_report['Status'].str.contains('LỖI AMAZON/SYNC').sum()
        err_diff    = len(df_final_report) - match_count - err_code - err_sync
        
        print("\n========================================")
        print("    BẢNG TÓM TẮT KẾT QUẢ ĐỐI CHIẾU      ")
        print("========================================")
        print(f" • Tổng số dòng kiểm tra : {len(df_final_report)}")
        print(f" • Khớp nhau (MATCH)     : {match_count}")
        print(f" • MISMATCH - LỖI CODE   : {err_code}")
        print(f" • MISMATCH - LỖI SYNC   : {err_sync}")
        if err_diff > 0:
            print(f" • MISMATCH - Lệch số liệu: {err_diff}")
        print("========================================")
        print("Gợi ý phân tích kết quả:")
        print(" - Nếu có nhiều 'MISMATCH - LỖI CODE': Cần kiểm tra lại hàm trích xuất/đọc file Bulk.")
        print(" - Nếu có nhiều 'MISMATCH - LỖI AMAZON/SYNC': Dữ liệu cấu hình nội bộ có set nhưng")
        print("   chưa được đẩy thành công lên Amazon, hoặc đã bị ghi đè/mất trên Seller Central.")
        
    except Exception as e:
        print(f"\n[Lỗi] Không thể ghi file báo cáo Excel: {e}")
        print("      Vui lòng đảm bảo file không bị mở bởi ứng dụng khác (Excel) khi chạy script.")


if __name__ == "__main__":
    cross_check_placements()
