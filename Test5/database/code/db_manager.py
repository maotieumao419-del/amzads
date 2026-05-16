import sqlite3
import pandas as pd
import os
from datetime import timedelta

def get_db_connection(db_path):
    conn = sqlite3.connect(db_path)
    return conn

def init_db(db_path):
    """Khởi tạo cấu trúc bảng nếu chưa tồn tại"""
    conn = get_db_connection(db_path)
    cursor = conn.cursor()
    
    # Bảng chứa dữ liệu daily
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS campaign_daily (
            date TEXT,
            campaign_name TEXT,
            impressions REAL,
            clicks REAL,
            spend REAL,
            sales REAL,
            orders REAL,
            PRIMARY KEY (date, campaign_name)
        )
    ''')
    conn.commit()
    conn.close()

def upsert_daily_data(db_path, df):
    """
    Nhận một DataFrame chuẩn hóa (date, campaign_name, impressions, clicks, spend, sales, orders)
    và UPSERT vào Database.
    """
    conn = get_db_connection(db_path)
    cursor = conn.cursor()
    
    # Sử dụng REPLACE INTO để Upsert (nếu đã có date+campaign_name thì ghi đè, nếu chưa thì thêm mới)
    records = df[['date', 'campaign_name', 'impressions', 'clicks', 'spend', 'sales', 'orders']].values.tolist()
    
    cursor.executemany('''
        REPLACE INTO campaign_daily (date, campaign_name, impressions, clicks, spend, sales, orders)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', records)
    
    conn.commit()
    conn.close()

def ingest_bulk_to_db(db_path, df_camp, start_date_str, end_date_str, days_duration):
    """
    Chia đều số liệu của file Bulk ra các ngày tương ứng và lưu vào Database.
    Ví dụ file Bulk 4 ngày có Spend=40 -> Tạo 4 dòng, mỗi dòng Spend=10.
    """
    if days_duration <= 0:
        return
        
    start_date = pd.to_datetime(start_date_str)
    
    records = []
    for _, row in df_camp.iterrows():
        camp_name = row.get('Campaign Name', '')
        if not camp_name:
            continue
            
        daily_impressions = row.get('Impressions', 0) / days_duration
        daily_clicks = row.get('Clicks', 0) / days_duration
        daily_spend = row.get('Spend', 0) / days_duration
        daily_sales = row.get('Sales', 0) / days_duration
        daily_orders = row.get('Orders', 0) / days_duration
        
        for i in range(days_duration):
            current_date = (start_date + timedelta(days=i)).strftime('%Y-%m-%d')
            records.append({
                'date': current_date,
                'campaign_name': camp_name,
                'impressions': daily_impressions,
                'clicks': daily_clicks,
                'spend': daily_spend,
                'sales': daily_sales,
                'orders': daily_orders
            })
            
    if records:
        df_daily = pd.DataFrame(records)
        upsert_daily_data(db_path, df_daily)

def get_time_series_data(db_path, start_date, end_date):
    """
    Truy vấn dữ liệu từ DB giữa start_date và end_date.
    Sau đó áp dụng Zero-fill (điền 0) cho những ngày bị thiếu của từng Campaign.
    Trả về DataFrame chuẩn để vẽ biểu đồ.
    """
    conn = get_db_connection(db_path)
    
    query = '''
        SELECT date, campaign_name, impressions, clicks, spend, sales, orders
        FROM campaign_daily
        WHERE date >= ? AND date <= ?
    '''
    df_raw = pd.read_sql_query(query, conn, params=(start_date, end_date))
    conn.close()
    
    if df_raw.empty:
        return pd.DataFrame()
        
    df_raw['date'] = pd.to_datetime(df_raw['date'])
    
    # Tạo dải ngày chuẩn (Reference Date Range) từ start_date đến end_date
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    
    # Lấy danh sách tất cả các Campaign có trong khoảng thời gian này
    campaigns = df_raw['campaign_name'].unique()
    
    # Tạo MultiIndex chuẩn chứa tất cả sự kết hợp giữa Ngày và Campaign
    multi_idx = pd.MultiIndex.from_product([date_range, campaigns], names=['date', 'campaign_name'])
    
    # Đặt index cho df_raw để merge
    df_raw = df_raw.set_index(['date', 'campaign_name'])
    
    # Reindex (Đây là kỹ thuật Zero-fill)
    df_filled = df_raw.reindex(multi_idx, fill_value=0).reset_index()
    
    # Ép kiểu lại ngày thành dạng chuỗi yyyy-mm-dd cho đồng nhất
    df_filled['date'] = df_filled['date'].dt.strftime('%Y-%m-%d')
    
    return df_filled
