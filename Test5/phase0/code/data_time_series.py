import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DB_DIR = os.path.join(BASE_DIR, "database", "code")
sys.path.append(DB_DIR)

from db_manager import get_time_series_data

DB_PATH = os.path.join(BASE_DIR, "database", "amazon_ads_history.db")

def generate_mock_ts_data(df_camp, today, days=14):
    """
    Giả lập dữ liệu Time Series (Fallback nếu Database trống)
    """
    records = []
    
    for _, row in df_camp.iterrows():
        camp_id = row['Campaign ID']
        camp_name = row.get('Campaign Name', '')
        total_spend = row.get('Spend', 0)
        total_sales = row.get('Sales', 0)
        
        if total_spend == 0 and total_sales == 0:
            spend_arr = np.zeros(days)
            sales_arr = np.zeros(days)
        else:
            spend_weights = np.random.rand(days)
            sales_weights = np.random.rand(days)
            spend_arr = spend_weights / spend_weights.sum() * total_spend if spend_weights.sum() > 0 else np.zeros(days)
            sales_arr = sales_weights / sales_weights.sum() * total_sales if sales_weights.sum() > 0 else np.zeros(days)
        
        for i in range(days):
            date_val = today - timedelta(days=(days - i - 1))
            records.append({
                'Date': date_val,
                'Campaign Name': camp_name,
                'Campaign ID': camp_id,
                'Spend': spend_arr[i],
                'Sales': sales_arr[i]
            })
            
    return pd.DataFrame(records)

def process_time_series(df_camp, today=None, days=14):
    if today is None:
        today = datetime.now().date()
        
    end_date_str = today.strftime('%Y-%m-%d')
    start_date_str = (today - timedelta(days=days-1)).strftime('%Y-%m-%d')
    
    # 1. Gọi dữ liệu từ Database
    try:
        df_ts_real = get_time_series_data(DB_PATH, start_date_str, end_date_str)
    except Exception as e:
        print(f"Khong the ket noi Database ({e}). Dung Mock Data.")
        df_ts_real = pd.DataFrame()
    
    # 2. Xử lý logic Điền khuyết (Zero-fill) và Nối ID
    if df_ts_real.empty:
        print(f"CANH BAO: Database trong tu {start_date_str} den {end_date_str}. Dang dung Mock Data.")
        df_ts = generate_mock_ts_data(df_camp, today, days=days)
    else:
        # Chuẩn hóa tên cột
        df_ts_real.rename(columns={
            'date': 'Date',
            'campaign_name': 'Campaign Name',
            'spend': 'Spend',
            'sales': 'Sales',
            'impressions': 'Impressions',
            'clicks': 'Clicks',
            'orders': 'Orders'
        }, inplace=True)
        
        # Ánh xạ Campaign ID từ file Bulk
        mapping = df_camp[['Campaign Name', 'Campaign ID']].drop_duplicates()
        df_ts = pd.merge(df_ts_real, mapping, on='Campaign Name', how='left')
        
        # Bỏ đi những Campaign cũ trong lịch sử nhưng không có trong file Bulk hiện tại
        df_ts.dropna(subset=['Campaign ID'], inplace=True)
        df_ts['Date'] = pd.to_datetime(df_ts['Date']).dt.date
    
    # 3. Pivot dữ liệu ra 2 dạng Wide (Sparklines) và Long (Deep Dive)
    df_ts['Date_Str'] = df_ts['Date'].apply(lambda x: x.strftime('%Y%m%d'))
    
    # Pivot Spend
    pivot_spend = df_ts.pivot_table(index='Campaign ID', columns='Date_Str', values='Spend', aggfunc='sum').fillna(0)
    pivot_spend.columns = [f"TS_Spend_{c}" for c in pivot_spend.columns]
    
    # Pivot Sales
    pivot_sales = df_ts.pivot_table(index='Campaign ID', columns='Date_Str', values='Sales', aggfunc='sum').fillna(0)
    pivot_sales.columns = [f"TS_Sales_{c}" for c in pivot_sales.columns]
    
    df_ts_agg = pivot_spend.join(pivot_sales)
    
    return df_ts, df_ts_agg
