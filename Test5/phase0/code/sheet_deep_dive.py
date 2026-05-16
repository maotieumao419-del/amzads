import pandas as pd

def build_deep_dive_sheet(workbook, writer, df_ts, df_camp_out, df_kw, data_type='bulk'):
    df_raw = df_ts.copy()
    
    df_raw = df_raw[['Date', 'Campaign ID', 'Spend', 'Sales', 'Campaign Name']]

    # 1. Lưu RAW_TS (dạng dọc) vào sheet ẩn
    raw_sheet_name = 'RAW_TS'
    df_raw.to_excel(writer, sheet_name=raw_sheet_name, index=False)
    ws_raw = writer.sheets[raw_sheet_name]
    ws_raw.hide()
    
    # 1.5 Lưu RAW_KW (dạng dọc) vào sheet ẩn
    raw_kw_sheet_name = 'RAW_KW'
    cols_kw = ["Campaign ID", "Campaign Name (Informational only)", "Date", "Keyword / Target", "Match Type", "Bid", "State", "Impressions", "Clicks", "Spend", "Sales"]
    cols_to_use = [c for c in cols_kw if c in df_kw.columns]
    df_kw_raw = df_kw[cols_to_use].copy()
    
    # Thêm cột ACOS
    if 'Spend' in df_kw_raw.columns and 'Sales' in df_kw_raw.columns:
        df_kw_raw['ACOS'] = df_kw_raw.apply(lambda row: row['Spend'] / row['Sales'] if pd.notna(row['Sales']) and row['Sales'] > 0 else 0, axis=1)
    else:
        df_kw_raw['ACOS'] = 0
    
    df_kw_raw.to_excel(writer, sheet_name=raw_kw_sheet_name, index=False)
    ws_kw = writer.sheets[raw_kw_sheet_name]
    ws_kw.hide()
    
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
    ws_dd.write('A2', 'Chọn Campaign Name:')
    ws_dd.write('A3', 'Hoặc chọn Campaign ID:')
    
    from xlsxwriter.utility import xl_col_to_name
    max_row = len(df_camp_out) + 1

    # Validation cho Campaign Name (B2)
    col_idx_name = df_camp_out.columns.get_loc('Campaign Name')
    camp_col_letter_name = xl_col_to_name(col_idx_name)
    validation_range_name = f"='🎯 CAMPAIGNS'!${camp_col_letter_name}$2:${camp_col_letter_name}${max_row}"
    
    ws_dd.data_validation('B2', {
        'validate': 'list',
        'source': validation_range_name,
        'input_title': 'Chọn Campaign Name',
        'input_message': 'Hãy chọn một Campaign Name từ danh sách.'
    })
    
    # Validation cho Campaign ID (B3)
    col_idx_id = df_camp_out.columns.get_loc('Campaign ID')
    camp_col_letter_id = xl_col_to_name(col_idx_id)
    validation_range_id = f"='🎯 CAMPAIGNS'!${camp_col_letter_id}$2:${camp_col_letter_id}${max_row}"
    
    ws_dd.data_validation('B3', {
        'validate': 'list',
        'source': validation_range_id,
        'input_title': 'Chọn Campaign ID',
        'input_message': 'Xóa ô B2 (Name) để lọc theo ID này.'
    })
    
    # Set default value
    first_camp_name = df_camp_out.iloc[0]['Campaign Name']
    ws_dd.write('B2', first_camp_name, fmt_dropdown)
    ws_dd.write('B3', '', fmt_dropdown)
    
    # Header bảng chi tiết
    ws_dd.write('A5', 'Date', fmt_header)
    ws_dd.write('B5', 'Spend', fmt_header)
    ws_dd.write('C5', 'Sales', fmt_header)
    ws_dd.write('D5', 'ACOS', fmt_header)
    
    # Lấy ngày từ df_raw (nếu bulk thì chỉ có 1 row date range)
    unique_dates = sorted(df_raw['Date'].unique(), reverse=True)
    
    # Định dạng chuỗi cho Date nếu là bulk
    fmt_date_str = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    
    row_start = 5 # Row 6 in Excel (index 5)
    for i, date_val in enumerate(unique_dates):
        current_row = row_start + i
        excel_row = current_row + 1
        
        if data_type == 'bulk':
            ws_dd.write_string(current_row, 0, str(date_val), fmt_date_str)
        else:
            ws_dd.write_datetime(current_row, 0, pd.to_datetime(date_val), fmt_date)
        
        # SUMIFS formulas
        # Ưu tiên B2 (Name). Nếu B2 trống thì dùng B3 (ID).
        spend_formula = f'=IF($B$2<>"", SUMIFS(\'RAW_TS\'!$C:$C, \'RAW_TS\'!$E:$E, $B$2, \'RAW_TS\'!$A:$A, $A{excel_row}), SUMIFS(\'RAW_TS\'!$C:$C, \'RAW_TS\'!$B:$B, $B$3, \'RAW_TS\'!$A:$A, $A{excel_row}))'
        ws_dd.write_formula(current_row, 1, spend_formula, fmt_money)
        
        sales_formula = f'=IF($B$2<>"", SUMIFS(\'RAW_TS\'!$D:$D, \'RAW_TS\'!$E:$E, $B$2, \'RAW_TS\'!$A:$A, $A{excel_row}), SUMIFS(\'RAW_TS\'!$D:$D, \'RAW_TS\'!$B:$B, $B$3, \'RAW_TS\'!$A:$A, $A{excel_row}))'
        ws_dd.write_formula(current_row, 2, sales_formula, fmt_money)
        
        acos_formula = f"=IF(C{excel_row}>0, B{excel_row}/C{excel_row}, 0)"
        ws_dd.write_formula(current_row, 3, acos_formula, fmt_pct)
        
    # Thêm Bar Chart cho biểu đồ 14 ngày
    chart = workbook.add_chart({'type': 'column'})
    last_row = row_start + len(unique_dates)
    
    # Cấu hình Series Spend
    chart.add_series({
        'name': 'Spend',
        'categories': f"='🔍 DEEP DIVE'!$A$6:$A${last_row}",
        'values': f"='🔍 DEEP DIVE'!$B$6:$B${last_row}",
        'fill': {'color': '#ED7D31'}
    })
    
    # Cấu hình Series Sales
    chart.add_series({
        'name': 'Sales',
        'categories': f"='🔍 DEEP DIVE'!$A$6:$A${last_row}",
        'values': f"='🔍 DEEP DIVE'!$C$6:$C${last_row}",
        'fill': {'color': '#4472C4'}
    })
    
    if data_type == 'bulk':
        chart.set_title({'name': f'Biểu đồ Chi phí & Doanh thu (Khoảng ngày)'})
    else:
        chart.set_title({'name': 'Biểu đồ Chi phí & Doanh thu 14 Ngày'})
        
    chart.set_x_axis({'name': 'Khoảng thời gian' if data_type == 'bulk' else 'Ngày'})
    chart.set_y_axis({'name': 'USD ($)'})
    
    ws_dd.insert_chart('F4', chart, {'x_scale': 1.5, 'y_scale': 1.5})
    
    # 3. Thêm bảng thông số Keyword
    # Đảm bảo bảng keyword không bị đè bởi biểu đồ (biểu đồ scale 1.5 bắt đầu từ row 4 chiếm khoảng 22-25 rows)
    kw_start_row = max(last_row + 4, 30)
    ws_dd.merge_range(kw_start_row, 0, kw_start_row, 9, 'LỊCH SỬ THÔNG SỐ KEYWORD THEO TỪNG BULK', fmt_title)
    
    kw_headers = ["Date", "Keyword / Target", "Match Type", "Bid", "State", "Impressions", "Clicks", "Spend", "Sales", "ACOS"]
    for i, header in enumerate(kw_headers):
        ws_dd.write(kw_start_row + 2, i, header, fmt_header)
        
    max_kw_row = len(df_kw_raw) + 1
    # Dữ liệu xuất từ cột C (Date) đến cột L (ACOS)
    # A=CampID, B=CampName, C=Date, D=Keyword, E=Match, F=Bid, G=State, H=Imp, I=Clicks, J=Spend, K=Sales, L=ACOS
    col_c_to_l = f"'RAW_KW'!C2:L{max_kw_row}"
    
    # Dùng hàm FILTER của Excel
    filter_formula = f'=IF($B$2<>"", FILTER({col_c_to_l}, \'RAW_KW\'!B2:B{max_kw_row}=$B$2, "Không có dữ liệu"), FILTER({col_c_to_l}, \'RAW_KW\'!A2:A{max_kw_row}=$B$3, "Không có dữ liệu"))'
    
    ws_dd.write_dynamic_array_formula(kw_start_row + 3, 0, kw_start_row + 3, 9, filter_formula)
    
    # Căn chỉnh kích thước cột cho đẹp
    ws_dd.set_column('E:L', 12)
    
    return ws_dd
