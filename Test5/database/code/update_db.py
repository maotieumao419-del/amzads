import os
import glob
import pandas as pd
from db_manager import init_db, upsert_daily_data

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_DIR = os.path.join(BASE_DIR, "input")
DB_PATH = os.path.join(BASE_DIR, "amazon_ads_history.db")

def process_daily_report(file_path):
    print(f"Dang xu ly file: {os.path.basename(file_path)}")
    
    # Đọc file (có thể là CSV hoặc Excel)
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Lỗi khi đọc file: {e}")
        return
        
    # Xóa khoảng trắng ở tên cột
    df.columns = df.columns.str.strip()
    
    # Chuẩn hóa tên cột
    col_map = {
        'Date': 'date',
        'Campaign Name': 'campaign_name',
        'Impressions': 'impressions',
        'Clicks': 'clicks',
        'Spend': 'spend',
        '7 Day Total Sales': 'sales',
        '14 Day Total Sales': 'sales',
        '30 Day Total Sales': 'sales',
        'Sales': 'sales',
        '7 Day Total Orders (#)': 'orders',
        '14 Day Total Orders (#)': 'orders',
        '30 Day Total Orders (#)': 'orders',
        'Orders': 'orders'
    }
    
    df.rename(columns=lambda x: col_map.get(x, x.lower()), inplace=True)
    
    # Lọc ra các cột cần thiết, nếu thiếu thì báo lỗi
    if 'date' not in df.columns or 'campaign_name' not in df.columns:
        print(f"File {os.path.basename(file_path)} không có cột Date hoặc Campaign Name. Bỏ qua.")
        return
        
    required_cols = ['date', 'campaign_name', 'impressions', 'clicks', 'spend', 'sales', 'orders']
    for col in required_cols:
        if col not in df.columns:
            df[col] = 0
            
    df = df[required_cols].copy()
    
    # Chuẩn hóa định dạng ngày (ép về YYYY-MM-DD)
    df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')
    
    # Xóa khoảng trắng ở campaign_name
    df['campaign_name'] = df['campaign_name'].astype(str).str.strip()
    
    # Điền giá trị Na cho số học
    df.fillna(0, inplace=True)
    
    # Upsert vào DB
    upsert_daily_data(DB_PATH, df)
    print(f"=> Đã nạp thành công {len(df)} dòng vào cơ sở dữ liệu.")

def main():
    # 1. Đảm bảo DB tồn tại
    init_db(DB_PATH)
    
    # 2. Quét các file trong input
    input_files = glob.glob(os.path.join(INPUT_DIR, "*.*"))
    valid_files = [f for f in input_files if (f.endswith('.csv') or f.endswith('.xlsx')) and not f.endswith('.processed')]
    
    if not valid_files:
        print("Khong tim thay file Bao cao Daily moi (CSV/Excel) nao trong thu muc input/")
        return
        
    for file_path in valid_files:
        try:
            process_daily_report(file_path)
            # Sau khi nạp xong, đổi tên file để đánh dấu đã xử lý
            processed_path = file_path + ".processed"
            if os.path.exists(processed_path):
                os.remove(processed_path)
            os.rename(file_path, processed_path)
        except Exception as e:
            print(f"Loi khi xu ly file {file_path}: {str(e)}")

if __name__ == "__main__":
    main()
