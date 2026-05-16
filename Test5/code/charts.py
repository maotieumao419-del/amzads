def add_sparkline_sheet(workbook, writer, df_ts_agg):
    # Tạo sheet ẩn để lưu data vẽ sparklines
    sheet_name = 'TS_DATA'
    df_ts_agg.reset_index(inplace=True)
    df_ts_agg.to_excel(writer, sheet_name=sheet_name, index=False)
    
    ws_ts = writer.sheets[sheet_name]
    ws_ts.hide() # Ẩn sheet này đi vì chỉ để chứa data
    
    return df_ts_agg

def draw_campaign_sparklines(workbook, writer, df_camp_out, df_ts_agg):
    ws_camp = writer.sheets['🎯 CAMPAIGNS']
    
    # Tìm cột "Campaign ID" trong CAMPAIGNS
    col_idx_cid = df_camp_out.columns.get_loc("Campaign ID")
    
    # Thêm 2 cột mới cho Sparkline ở cuối
    num_cols = len(df_camp_out.columns)
    col_spend_spark = num_cols
    col_sales_spark = num_cols + 1
    
    fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2'})
    ws_camp.write(0, col_spend_spark, "14-Day Spend Trend", fmt_header)
    ws_camp.write(0, col_sales_spark, "14-Day Sales Trend", fmt_header)
    
    ws_camp.set_column(col_spend_spark, col_sales_spark, 20)
    
    # Lấy danh sách cột Spend và Sales trong TS_DATA
    spend_cols = [c for c in df_ts_agg.columns if c.startswith('TS_Spend_')]
    sales_cols = [c for c in df_ts_agg.columns if c.startswith('TS_Sales_')]
    
    if not spend_cols or not sales_cols:
        return
        
    ts_sheet_name = 'TS_DATA'
    
    # Duyệt từng dòng trong CAMPAIGNS sheet
    for row_idx in range(len(df_camp_out)):
        camp_id = df_camp_out.iloc[row_idx, col_idx_cid]
        
        # Tìm dòng tương ứng trong TS_DATA
        ts_match = df_ts_agg.index[df_ts_agg['Campaign ID'] == camp_id].tolist()
        if ts_match:
            ts_row = ts_match[0] + 2  # Excel row (1-based, +1 for header)
            
            # Cột đầu và cuối của Spend trong TS_DATA
            spend_start_col = df_ts_agg.columns.get_loc(spend_cols[0])
            spend_end_col = df_ts_agg.columns.get_loc(spend_cols[-1])
            
            # Cột đầu và cuối của Sales trong TS_DATA
            sales_start_col = df_ts_agg.columns.get_loc(sales_cols[0])
            sales_end_col = df_ts_agg.columns.get_loc(sales_cols[-1])
            
            # Hàm chuyển đổi index số sang chữ (vd: 0 -> A, 1 -> B)
            from xlsxwriter.utility import xl_col_to_name
            spend_range = f"{ts_sheet_name}!{xl_col_to_name(spend_start_col)}{ts_row}:{xl_col_to_name(spend_end_col)}{ts_row}"
            sales_range = f"{ts_sheet_name}!{xl_col_to_name(sales_start_col)}{ts_row}:{xl_col_to_name(sales_end_col)}{ts_row}"
            
            # Vẽ Sparkline
            excel_row = row_idx + 1 # +1 for header
            ws_camp.add_sparkline(excel_row, col_spend_spark, {
                'range': spend_range,
                'type': 'column'
            })
            
            ws_camp.add_sparkline(excel_row, col_sales_spark, {
                'range': sales_range,
                'type': 'line',
                'markers': True
            })
