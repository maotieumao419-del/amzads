import pandas as pd

def build_deep_dive_sheet(workbook, writer, df_ts, df_camp_out):
    # 1. Lưu RAW_TS (dạng dọc) vào sheet ẩn
    raw_sheet_name = 'RAW_TS'
    df_ts.to_excel(writer, sheet_name=raw_sheet_name, index=False)
    ws_raw = writer.sheets[raw_sheet_name]
    ws_raw.hide()
    
    # 2. Tạo sheet DEEP DIVE
    sheet_name = '🔍 DEEP DIVE'
    ws_dd = workbook.add_worksheet(sheet_name)
    
    fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'color': '#FFFFFF', 'bg_color': '#44546A'})
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center'})
    fmt_date = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})
    fmt_money = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
    fmt_pct = workbook.add_format({'num_format': '0.00%', 'border': 1})
    fmt_dropdown = workbook.add_format({'border': 1, 'bg_color': '#FFFFCC', 'bold': True})
    
    ws_dd.set_column('A:A', 15)
    ws_dd.set_column('B:D', 15)
    
    # Hướng dẫn
    ws_dd.merge_range('A1:D1', 'CHI TIẾT LỊCH SỬ CHI TIÊU THEO NGÀY', fmt_title)
    ws_dd.write('A2', 'Chọn Campaign ID:')
    
    # Lấy danh sách Campaign ID để làm Data Validation
    # Ta tham chiếu đến cột Campaign ID ở sheet 🎯 CAMPAIGNS (Giả sử nằm ở cột F)
    # Lấy index cột Campaign ID trong df_camp_out
    col_idx = df_camp_out.columns.get_loc('Campaign ID')
    from xlsxwriter.utility import xl_col_to_name
    camp_col_letter = xl_col_to_name(col_idx)
    max_row = len(df_camp_out) + 1
    validation_range = f"='🎯 CAMPAIGNS'!${camp_col_letter}$2:${camp_col_letter}${max_row}"
    
    ws_dd.data_validation('B2', {
        'validate': 'list',
        'source': validation_range,
        'input_title': 'Chọn Campaign',
        'input_message': 'Hãy chọn một Campaign từ danh sách thả xuống.'
    })
    
    # Set default value (Campaign đầu tiên)
    first_camp_id = df_camp_out.iloc[0]['Campaign ID']
    ws_dd.write('B2', first_camp_id, fmt_dropdown)
    
    # Header bảng chi tiết
    ws_dd.write('A4', 'Date', fmt_header)
    ws_dd.write('B4', 'Spend', fmt_header)
    ws_dd.write('C4', 'Sales', fmt_header)
    ws_dd.write('D4', 'ACOS', fmt_header)
    
    # Lấy 14 ngày duy nhất từ df_ts
    unique_dates = sorted(df_ts['Date'].unique(), reverse=True)
    
    row_start = 4 # Row 5 in Excel (index 4)
    for i, date_val in enumerate(unique_dates):
        current_row = row_start + i
        excel_row = current_row + 1
        
        ws_dd.write_datetime(current_row, 0, pd.to_datetime(date_val), fmt_date)
        
        # SUMIFS formulas
        # Spend = SUMIFS(RAW_TS!$C:$C, RAW_TS!$B:$B, $B$2, RAW_TS!$A:$A, $A5)
        # RAW_TS column mapping: A:Date, B:Campaign ID, C:Spend, D:Sales
        spend_formula = f"=SUMIFS('RAW_TS'!$C:$C, 'RAW_TS'!$B:$B, $B$2, 'RAW_TS'!$A:$A, $A{excel_row})"
        ws_dd.write_formula(current_row, 1, spend_formula, fmt_money)
        
        sales_formula = f"=SUMIFS('RAW_TS'!$D:$D, 'RAW_TS'!$B:$B, $B$2, 'RAW_TS'!$A:$A, $A{excel_row})"
        ws_dd.write_formula(current_row, 2, sales_formula, fmt_money)
        
        acos_formula = f"=IF(C{excel_row}>0, B{excel_row}/C{excel_row}, 0)"
        ws_dd.write_formula(current_row, 3, acos_formula, fmt_pct)
        
    # Thêm Bar Chart cho biểu đồ 14 ngày
    chart = workbook.add_chart({'type': 'column'})
    last_row = row_start + len(unique_dates)
    
    # Cấu hình Series Spend
    chart.add_series({
        'name': 'Spend',
        'categories': f"='🔍 DEEP DIVE'!$A$5:$A${last_row}",
        'values': f"='🔍 DEEP DIVE'!$B$5:$B${last_row}",
        'fill': {'color': '#ED7D31'}
    })
    
    # Cấu hình Series Sales
    chart.add_series({
        'name': 'Sales',
        'categories': f"='🔍 DEEP DIVE'!$A$5:$A${last_row}",
        'values': f"='🔍 DEEP DIVE'!$C$5:$C${last_row}",
        'fill': {'color': '#4472C4'}
    })
    
    chart.set_title({'name': 'Biểu đồ Chi phí & Doanh thu 14 Ngày'})
    chart.set_x_axis({'name': 'Ngày'})
    chart.set_y_axis({'name': 'USD ($)'})
    
    ws_dd.insert_chart('F4', chart, {'x_scale': 1.5, 'y_scale': 1.5})
    
    return ws_dd
