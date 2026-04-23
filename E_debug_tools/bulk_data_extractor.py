import pandas as pd
import os
import glob
from datetime import datetime

def find_column(df, search_names):
    """
    Hàm linh hoạt để tìm tên cột trong DataFrame dựa trên danh sách các tên khả thi.
    """
    for name in search_names:
        if name in df.columns:
            return name
        # Kiểm tra không phân biệt hoa thường
        for col in df.columns:
            if str(col).lower() == name.lower():
                return col
    return None

def extract_sku_mapping(file_path):
    """
    Đọc file Amazon Bulk, trích xuất Campaign Mapping và Metrics.
    """
    print("\n[STEP 1] Dang xu ly file: " + os.path.basename(file_path))
    
    try:
        # 1. Doc cu the sheet "Sponsored Products Campaigns"
        # Su dung engine openpyxl de xu ly file .xlsx
        df = pd.read_excel(file_path, sheet_name='Sponsored Products Campaigns')
        
        # 2. Loc Entity la 'Campaign'
        if 'Entity' not in df.columns:
            print(f"Error: Khong tim thay cot 'Entity' trong file.")
            return
            
        df_campaign = df[df['Entity'].str.lower() == 'campaign'].copy()
        
        if df_campaign.empty:
            print(f"Warning: Khong tim thay dong 'Campaign' nao trong sheet.")
            return

        # 3. Tim cac cot can thiet (Fallback logic)
        col_portfolio = find_column(df_campaign, ['Portfolio ID', 'Portfolio Id'])
        col_campaign = find_column(df_campaign, ['Campaign Name', 'Campaign Name (Informational only)'])
        col_state = find_column(df_campaign, ['State', 'Campaign State (Informational only)', 'Campaign State'])
        
        # Metrics columns
        col_imp = find_column(df_campaign, ['Impressions'])
        col_click = find_column(df_campaign, ['Clicks'])
        col_order = find_column(df_campaign, ['Orders', 'Seven Day Total Units Ordered'])

        # Kiem tra cac cot bat buoc
        if not col_campaign:
            print("Error: Khong tim thay cot Campaign Name.")
            return

        # 4. Trich xuat va lam sach du lieu
        # Tao dictionary mapping de doi ten cot cho dong nhat
        extract_map = {
            col_campaign: 'Campaign Name'
        }
        if col_portfolio: extract_map[col_portfolio] = 'Portfolio ID'
        if col_state: extract_map[col_state] = 'State'
        if col_imp: extract_map[col_imp] = 'Impressions'
        if col_click: extract_map[col_click] = 'Clicks'
        if col_order: extract_map[col_order] = 'Orders'

        # Loc lay cac cot can thiet
        df_result = df_campaign[list(extract_map.keys())].copy()
        df_result = df_result.rename(columns=extract_map)

        # Xu ly cac gia tri NaN cho Metrics (chuyen ve 0)
        metric_cols = ['Impressions', 'Clicks', 'Orders']
        for col in metric_cols:
            if col in df_result.columns:
                df_result[col] = df_result[col].fillna(0)

        # 5. Loai bo trung lap (Unique list)
        df_result = df_result.drop_duplicates()

        # 6. Xuat Benchmark/Output
        original_name = os.path.basename(file_path).replace('.xlsx', '')
        output_name = f"Mapping_Data_{original_name}.xlsx"
        output_path = os.path.join(os.path.dirname(file_path), output_name)
        
        df_result.to_excel(output_path, index=False)
        
        print(f"Completed! Da xuat du lieu ra: {output_name}")
        print(f"Total campaigns: {len(df_result)}")
        return output_path

    except Exception as e:
        print(f"Error khi xu ly file: {str(e)}")
        return None

if __name__ == "__main__":
    # Mac dinh quet folder D neu khong chay voi argument
    DEFAULT_INPUT_DIR = os.path.join("..", "D_auto_bulk_updater", "data", "input 1")
    
    # Kiem tra xem folder co ton tai khong (neu chay tu E_debug_tools)
    if not os.path.exists(DEFAULT_INPUT_DIR):
        # Thu lai voi path truc tiep neu chay tu root
        DEFAULT_INPUT_DIR = os.path.join("D_auto_bulk_updater", "data", "input 1")

    print("--- AMAZON BULK DATA EXTRACTOR ---")
    
    # Tim cac file .xlsx trong folder input
    files = glob.glob(os.path.join(DEFAULT_INPUT_DIR, "*.xlsx"))
    # Loai bo cac file tam cua Excel (~$) va file Mapping cu
    files = [f for f in files if not os.path.basename(f).startswith('~$') and not os.path.basename(f).startswith('Mapping_Data')]

    if not files:
        print(f"Empty: Khong tim thay file Bulk nao trong: {os.path.abspath(DEFAULT_INPUT_DIR)}")
    else:
        print(f"Found {len(files)} file Bulk. Bat dau xu ly...")
        for file in files:
            extract_sku_mapping(file)
    
    print("\nFinished.")
