import pandas as pd

def build_daily_summary_sheet(workbook, writer, df_daily_summary):
    sheet_name = '📅 DAILY TRENDS'
    ws_daily = workbook.add_worksheet(sheet_name)
    
    fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'color': '#FFFFFF', 'bg_color': '#44546A'})
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    fmt_date = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})
    fmt_money = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
    fmt_pct = workbook.add_format({'num_format': '0.00%', 'border': 1})
    fmt_number = workbook.add_format({'border': 1, 'align': 'center'})
    
    ws_daily.set_column('A:A', 15)
    ws_daily.set_column('B:E', 15)
    
    # Hướng dẫn
    ws_daily.merge_range('A1:E1', 'TỔNG QUAN CHI TIÊU TOÀN TÀI KHOẢN THEO NGÀY', fmt_title)
    
    # Header bảng chi tiết
    headers = ['Date', 'Spend', 'Sales', 'Orders', 'ACOS']
    for col_idx, header in enumerate(headers):
        ws_daily.write(3, col_idx, header, fmt_header)
        
    # Sắp xếp giảm dần theo Date để hiển thị ngày mới nhất ở trên
    df_sorted = df_daily_summary.sort_values('Date', ascending=False).reset_index(drop=True)
    
    row_start = 4 # Row 5 in Excel
    for i, row in df_sorted.iterrows():
        current_row = row_start + i
        date_val = row['Date']
        
        if isinstance(date_val, str):
            ws_daily.write_string(current_row, 0, str(date_val), workbook.add_format({'border': 1, 'align': 'center'}))
        else:
            ws_daily.write_datetime(current_row, 0, pd.to_datetime(date_val), fmt_date)
            
        ws_daily.write_number(current_row, 1, row.get('Spend', 0), fmt_money)
        ws_daily.write_number(current_row, 2, row.get('Sales', 0), fmt_money)
        ws_daily.write_number(current_row, 3, row.get('Orders', 0), fmt_number)
        
        acos = row.get('Spend', 0) / row.get('Sales', 1) if row.get('Sales', 0) > 0 else 0
        ws_daily.write_number(current_row, 4, acos, fmt_pct)
        
    # Thêm Line Chart
    chart = workbook.add_chart({'type': 'line'})
    last_row = row_start + len(df_sorted) - 1
    
    chart.add_series({
        'name': 'Spend',
        'categories': f"='{sheet_name}'!$A$5:$A${last_row+1}",
        'values': f"='{sheet_name}'!$B$5:$B${last_row+1}",
        'line': {'color': '#ED7D31', 'width': 2.25}
    })
    
    chart.add_series({
        'name': 'Sales',
        'categories': f"='{sheet_name}'!$A$5:$A${last_row+1}",
        'values': f"='{sheet_name}'!$C$5:$C${last_row+1}",
        'line': {'color': '#4472C4', 'width': 2.25}
    })
    
    chart.set_title({'name': 'Biểu đồ Chi phí & Doanh thu Tài khoản'})
    chart.set_x_axis({'name': 'Ngày', 'reverse': True})
    chart.set_y_axis({'name': 'USD ($)'})
    
    ws_daily.insert_chart('G4', chart, {'x_scale': 1.5, 'y_scale': 1.5})
    
    return ws_daily
