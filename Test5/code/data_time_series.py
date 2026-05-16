import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def generate_mock_ts_data(df_camp, days=14):
    """
    Giả lập dữ liệu Time Series (14 ngày) dựa trên tổng Spend/Sales của Campaign.
    Thực tế: Bạn cần tải file 'Sponsored Products Campaigns report' (Daily) từ Amazon.
    """
    records = []
    today = datetime.now().date()
    
    for _, row in df_camp.iterrows():
        camp_id = row['Campaign ID']
        total_spend = row.get('Spend', 0)
        total_sales = row.get('Sales', 0)
        
        # Tạo chuỗi ngẫu nhiên nhưng tổng gần bằng tổng trong Bulk
        # Để đơn giản, ta chia đều rồi cộng trừ nhiễu ngẫu nhiên
        if total_spend == 0 and total_sales == 0:
            spend_arr = np.zeros(days)
            sales_arr = np.zeros(days)
        else:
            # Sinh mảng ngẫu nhiên
            spend_weights = np.random.rand(days)
            sales_weights = np.random.rand(days)
            # Chuẩn hóa để tổng = total
            spend_arr = spend_weights / spend_weights.sum() * total_spend if spend_weights.sum() > 0 else np.zeros(days)
            sales_arr = sales_weights / sales_weights.sum() * total_sales if sales_weights.sum() > 0 else np.zeros(days)
        
        for i in range(days):
            date_val = today - timedelta(days=(days - i))
            records.append({
                'Date': date_val,
                'Campaign ID': camp_id,
                'Spend': spend_arr[i],
                'Sales': sales_arr[i]
            })
            
    return pd.DataFrame(records)

def process_time_series(df_camp, days=14):
    # Trong thực tế, bạn sẽ dùng pd.read_csv("input/Daily_Report.csv") 
    # Ở đây ta gọi hàm mock để có data demo
    df_ts = generate_mock_ts_data(df_camp, days=days)
    
    # Pivot dữ liệu để mỗi Campaign ID có 1 dòng, 14 cột Spend và 14 cột Sales
    df_ts['Date_Str'] = df_ts['Date'].apply(lambda x: x.strftime('%Y%m%d'))
    
    pivot_spend = df_ts.pivot(index='Campaign ID', columns='Date_Str', values='Spend').fillna(0)
    pivot_spend.columns = [f"TS_Spend_{c}" for c in pivot_spend.columns]
    
    pivot_sales = df_ts.pivot(index='Campaign ID', columns='Date_Str', values='Sales').fillna(0)
    pivot_sales.columns = [f"TS_Sales_{c}" for c in pivot_sales.columns]
    
    df_ts_agg = pivot_spend.join(pivot_sales)
    
    # Trả về cả dạng Long (để Deep Dive) và Wide (để vẽ Sparklines)
    return df_ts, df_ts_agg
