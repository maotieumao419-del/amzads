# ui_theme.py

def get_ui_formats(workbook):
    """
    Trả về dictionary chứa các format tĩnh (Theme) cho UI Dashboard.
    """
    return {
        'title': workbook.add_format({
            'bold': True, 'font_size': 14, 'bg_color': '#232F3E', 
            'color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter'
        }),
        'header': workbook.add_format({
            'bold': True, 'border': 1, 'bg_color': '#232F3E', 
            'color': '#FFFFFF', 'align': 'center'
        }),
        'dropdown': workbook.add_format({
            'border': 1, 'bg_color': '#FFFFCC', 'bold': True, 'align': 'center'
        }),
        'metric_num': workbook.add_format({
            'border': 1, 'align': 'center', 'num_format': '#,##0'
        }),
        'metric_currency': workbook.add_format({
            'border': 1, 'align': 'center', 'num_format': '$#,##0.00'
        }),
        'metric_percent': workbook.add_format({
            'border': 1, 'align': 'center', 'num_format': '0.00%'
        }),
        'cell_border': workbook.add_format({
            'border': 1, 'align': 'left'
        }),
        # Màu sắc Conditional Formatting
        'alert_red': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}),
        'alert_yellow': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}),
        'alert_green': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}),
        
        # Bảng màu cho Chart
        'chart_colors': {
            'spend': '#FF9900',  # Cam Amazon
            'sales': '#232F3E',  # Xanh đậm Amazon
            'acos': '#FF0000'    # Đỏ
        }
    }

def apply_conditional_formatting(worksheet, formats, target_acos):
    """
    Áp dụng luật chẩn đoán phễu PPC bằng conditional formatting của xlsxwriter.
    Giả định bảng dữ liệu nằm từ dòng 6 đến 50.
    Cột D: CTR, Cột H: CVR, Cột I: ACOS
    """
    red_fmt = formats['alert_red']
    yellow_fmt = formats['alert_yellow']
    green_fmt = formats['alert_green']

    # 1. Cột CTR (D) < 0.003 -> Báo Đỏ
    worksheet.conditional_format('D6:D50', {
        'type': 'cell',
        'criteria': '<',
        'value': 0.003,
        'format': red_fmt
    })
    
    # 2. Cột CVR (H) < 0.05 -> Báo Vàng
    worksheet.conditional_format('H6:H50', {
        'type': 'cell',
        'criteria': '<',
        'value': 0.05,
        'format': yellow_fmt
    })
    
    # 3. Logic ACOS (I) dựa vào ô Giai đoạn ở $B$3
    # Nếu B3="Maintain" VÀ ACOS > target_acos -> Báo Đỏ
    worksheet.conditional_format('I6:I50', {
        'type': 'formula',
        'criteria': f'=AND($B$3="Maintain", $I6>{target_acos})',
        'format': red_fmt
    })
    
    # Nếu B3="Maintain" VÀ ACOS <= target_acos (và lớn hơn 0 để bỏ qua ô trống) -> Báo Xanh
    worksheet.conditional_format('I6:I50', {
        'type': 'formula',
        'criteria': f'=AND($B$3="Maintain", $I6<={target_acos}, $I6>0)',
        'format': green_fmt
    })
