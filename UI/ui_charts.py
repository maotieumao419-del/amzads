# ui_charts.py

from xlsxwriter.utility import xl_col_to_name
from config import SHEET_TS_DATA


def add_ts_data_sheet(workbook, writer, df_camp_ts):
    """
    Lưu time-series aggregated data (Campaign level) vào sheet ẩn TS_DATA.
    df_camp_ts: DataFrame với index=Campaign Name, cols=TS_Spend_XXXX, TS_Sales_XXXX
    Trả về df_camp_ts đã reset_index để dùng cho sparklines.
    """
    df_out = df_camp_ts.reset_index()
    df_out.to_excel(writer, sheet_name=SHEET_TS_DATA, index=False)
    writer.sheets[SHEET_TS_DATA].hide()
    return df_out


def draw_campaign_sparklines(workbook, writer, df_camp, df_ts_agg):
    """
    Vẽ sparklines Spend cho từng Campaign trong CAMPAIGNS sheet.
    df_camp: DataFrame campaign với cột 'Campaign Name'
    df_ts_agg: DataFrame wide với cột 'Campaign Name', TS_Spend_*, TS_Sales_*
    """
    from config import SHEET_CAMPAIGNS
    if SHEET_CAMPAIGNS not in writer.sheets:
        return
    if df_ts_agg is None or df_ts_agg.empty:
        return

    ws_camp = writer.sheets[SHEET_CAMPAIGNS]

    spend_cols = [c for c in df_ts_agg.columns if c.startswith('TS_Spend_')]
    sales_cols = [c for c in df_ts_agg.columns if c.startswith('TS_Sales_')]

    if not spend_cols:
        return

    # Thêm 2 cột header sparkline ở cuối CAMPAIGNS sheet
    num_cols = df_camp.shape[1] if hasattr(df_camp, 'shape') else 14
    col_spark_spend = num_cols      # ngay sau cột cuối
    col_spark_sales = num_cols + 1

    fmt_hdr = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2',
                                   'align': 'center', 'color': '#232F3E'})
    ws_camp.write(1, col_spark_spend, 'Spend Trend', fmt_hdr)
    ws_camp.write(1, col_spark_sales, 'Sales Trend', fmt_hdr)
    ws_camp.set_column(col_spark_spend, col_spark_sales, 18)

    # Map Campaign Name → row index trong TS_DATA
    camp_to_ts_row = {}
    for ts_ri, ts_row in df_ts_agg.iterrows():
        cn = ts_row.get('Campaign Name', '')
        if cn:
            camp_to_ts_row[cn] = ts_ri

    # Vị trí cột trong TS_DATA
    ts_cols = df_ts_agg.columns.tolist()
    if spend_cols:
        spend_start = ts_cols.index(spend_cols[0])
        spend_end   = ts_cols.index(spend_cols[-1])
    if sales_cols:
        sales_start = ts_cols.index(sales_cols[0])
        sales_end   = ts_cols.index(sales_cols[-1])

    for camp_ri, camp_row in df_camp.reset_index(drop=True).iterrows():
        cn = str(camp_row.get('Campaign Name', ''))
        if cn not in camp_to_ts_row:
            continue

        ts_data_row = camp_to_ts_row[cn] + 2   # +1 header, +1 Excel 1-based

        spend_range = (f"{SHEET_TS_DATA}!"
                       f"{xl_col_to_name(spend_start)}{ts_data_row}:"
                       f"{xl_col_to_name(spend_end)}{ts_data_row}")
        sales_range = (f"{SHEET_TS_DATA}!"
                       f"{xl_col_to_name(sales_start)}{ts_data_row}:"
                       f"{xl_col_to_name(sales_end)}{ts_data_row}")

        excel_row = camp_ri + 2   # +2 (title row + header row)

        ws_camp.add_sparkline(excel_row, col_spark_spend, {
            'range': spend_range,
            'type': 'column',
            'high_point': True,
            'series_color': '#FF9900'
        })
        ws_camp.add_sparkline(excel_row, col_spark_sales, {
            'range': sales_range,
            'type': 'line',
            'markers': True,
            'series_color': '#232F3E'
        })


def draw_profitability_combo_chart(workbook, worksheet, chart_position, data_ranges, theme):
    """
    Vẽ biểu đồ Combo (Bar + Line) phục vụ phân tích lợi nhuận.
    - Trục chính: Spend (Cam), Sales (Xanh Đậm) dạng Cột.
    - Trục phụ: ACOS (Đỏ) dạng Đường (y2_axis).
    """
    column_chart = workbook.add_chart({'type': 'column'})

    column_chart.add_series({
        'name': 'Total Spend',
        'categories': data_ranges['categories'],
        'values':     data_ranges['spend'],
        'fill': {'color': theme['chart_colors']['spend']}
    })
    column_chart.add_series({
        'name': 'Total Sales',
        'categories': data_ranges['categories'],
        'values':     data_ranges['sales'],
        'fill': {'color': theme['chart_colors']['sales']}
    })

    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'name': 'ACOS %',
        'categories': data_ranges['categories'],
        'values':     data_ranges['acos'],
        'line':   {'color': theme['chart_colors']['acos'], 'width': 2},
        'marker': {'type': 'circle', 'size': 5,
                   'fill': {'color': theme['chart_colors']['acos']}},
        'y2_axis': True
    })

    column_chart.combine(line_chart)
    column_chart.set_title({'name': 'Phân Tích Lợi Nhuận: Spend vs Sales vs ACOS'})
    column_chart.set_x_axis({'name': 'Campaign'})
    column_chart.set_y_axis({'name': 'USD ($)', 'major_gridlines': {'visible': True}})
    column_chart.set_y2_axis({'name': 'ACOS (%)'})

    worksheet.insert_chart(chart_position, column_chart, {'x_scale': 1.6, 'y_scale': 1.2})
