# ui_charts.py

def draw_profitability_combo_chart(workbook, worksheet, chart_position, data_ranges, theme):
    """
    Vẽ biểu đồ Combo (Bar + Line) phục vụ phân tích lợi nhuận.
    Không xử lý tính toán data, chỉ nhận reference (data_ranges).
    - Trục chính: Spend (Cam), Sales (Xanh Đậm) dạng Cột.
    - Trục phụ: ACOS (Đỏ) dạng Đường (y2_axis).
    """
    # 1. Tạo Bar Chart cho Spend và Sales
    column_chart = workbook.add_chart({'type': 'column'})
    
    # Series: Total Spend
    column_chart.add_series({
        'name': 'Total Spend',
        'categories': data_ranges['categories'],
        'values': data_ranges['spend'],
        'fill': {'color': theme['chart_colors']['spend']}
    })
    
    # Series: Total Sales
    column_chart.add_series({
        'name': 'Total Sales',
        'categories': data_ranges['categories'],
        'values': data_ranges['sales'],
        'fill': {'color': theme['chart_colors']['sales']}
    })

    # 2. Tạo Line Chart cho ACOS (Trục Y2)
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'name': 'ACOS %',
        'categories': data_ranges['categories'],
        'values': data_ranges['acos'],
        'line': {'color': theme['chart_colors']['acos'], 'width': 2},
        'marker': {'type': 'circle', 'size': 5, 'fill': {'color': theme['chart_colors']['acos']}},
        'y2_axis': True
    })

    # 3. Combine hai chart
    column_chart.combine(line_chart)

    # 4. Định dạng tiêu đề và các trục
    column_chart.set_title({'name': 'Phân Tích Lợi Nhuận: Spend vs Sales vs ACOS'})
    column_chart.set_x_axis({'name': 'Campaign'})
    column_chart.set_y_axis({'name': 'USD ($)', 'major_gridlines': {'visible': True}})
    column_chart.set_y2_axis({'name': 'ACOS (%)'})

    # 5. Chèn chart vào worksheet
    worksheet.insert_chart(chart_position, column_chart, {'x_scale': 1.6, 'y_scale': 1.2})
