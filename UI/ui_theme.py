# ui_theme.py

def get_ui_formats(workbook):
    """
    Trả về dictionary chứa toàn bộ format cells cho Dashboard.
    Brand palette: #232F3E (Amazon Navy), #FF9900 (Amazon Orange), #FEBD69 (Amazon Gold)
    """
    return {
        # ── Tiêu đề sheet ──────────────────────────────────────────────────
        'title': workbook.add_format({
            'bold': True, 'font_size': 14,
            'bg_color': '#232F3E', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter'
        }),
        'subtitle': workbook.add_format({
            'bold': True, 'font_size': 11,
            'bg_color': '#37475A', 'color': '#FEBD69',
            'align': 'left', 'valign': 'vcenter'
        }),

        # ── Header bảng ────────────────────────────────────────────────────
        'header': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#232F3E', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter',
            'text_wrap': True
        }),
        'header_orange': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#FF9900', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter'
        }),
        'header_light': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#D9E1F2', 'color': '#232F3E',
            'align': 'center', 'valign': 'vcenter'
        }),

        # ── Ô nhập liệu (Dropdown) ─────────────────────────────────────────
        'dropdown': workbook.add_format({
            'border': 2, 'bg_color': '#FFFFCC', 'bold': True,
            'align': 'left', 'valign': 'vcenter'
        }),
        'label': workbook.add_format({
            'bold': True, 'bg_color': '#F2F2F2',
            'border': 1, 'align': 'right', 'valign': 'vcenter'
        }),

        # ── Metric formats ─────────────────────────────────────────────────
        'metric_num': workbook.add_format({
            'border': 1, 'align': 'center', 'num_format': '#,##0'
        }),
        'metric_currency': workbook.add_format({
            'border': 1, 'align': 'center', 'num_format': '$#,##0.00'
        }),
        'metric_percent': workbook.add_format({
            'border': 1, 'align': 'center', 'num_format': '0.00%'
        }),
        'metric_text': workbook.add_format({
            'border': 1, 'align': 'left', 'valign': 'vcenter'
        }),
        'metric_center': workbook.add_format({
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        }),

        # ── Health Tags ────────────────────────────────────────────────────
        'health_star': workbook.add_format({
            'border': 1, 'align': 'center', 'bold': True,
            'bg_color': '#C6EFCE', 'font_color': '#006100'
        }),
        'health_bleeder': workbook.add_format({
            'border': 1, 'align': 'center', 'bold': True,
            'bg_color': '#FFC7CE', 'font_color': '#9C0006'
        }),
        'health_watch': workbook.add_format({
            'border': 1, 'align': 'center', 'bold': True,
            'bg_color': '#FFEB9C', 'font_color': '#9C6500'
        }),
        'health_dead': workbook.add_format({
            'border': 1, 'align': 'center', 'bold': True,
            'bg_color': '#808080', 'font_color': '#FFFFFF'
        }),
        'health_sleep': workbook.add_format({
            'border': 1, 'align': 'center',
            'bg_color': '#EDEDED', 'font_color': '#666666'
        }),
        'health_new': workbook.add_format({
            'border': 1, 'align': 'center',
            'bg_color': '#DDEEFF', 'font_color': '#004080'
        }),

        # ── Suggested Action ───────────────────────────────────────────────
        'suggest_increase': workbook.add_format({
            'border': 1, 'align': 'center',
            'bg_color': '#E2EFDA', 'font_color': '#375623'
        }),
        'suggest_decrease': workbook.add_format({
            'border': 1, 'align': 'center',
            'bg_color': '#FCE4D6', 'font_color': '#833C00'
        }),
        'suggest_pause': workbook.add_format({
            'border': 1, 'align': 'center',
            'bg_color': '#FFC7CE', 'font_color': '#9C0006'
        }),
        'suggest_monitor': workbook.add_format({
            'border': 1, 'align': 'center',
            'bg_color': '#F2F2F2', 'font_color': '#595959'
        }),

        # ── Conditional Formatting colors ──────────────────────────────────
        'alert_red':    workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}),
        'alert_yellow': workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}),
        'alert_green':  workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}),

        # ── Chart palette (dict, không phải format object) ─────────────────
        'chart_colors': {
            'spend': '#FF9900',   # Amazon Orange
            'sales': '#232F3E',   # Amazon Navy
            'acos':  '#CC0000',   # Đỏ đậm
            'ctr':   '#5B9BD5',   # Xanh dương
        }
    }


def apply_conditional_formatting(worksheet, formats, target_acos):
    """
    Áp dụng conditional formatting cho bảng chi tiết campaign/keyword.
    Dùng cho cả CAMPAIGNS sheet và DEEP DIVE sheet.
    col_acos_letter: chữ cột Excel chứa ACOS (vd: 'I')
    """
    red_fmt    = formats['alert_red']
    yellow_fmt = formats['alert_yellow']
    green_fmt  = formats['alert_green']

    # ACOS > target → đỏ
    worksheet.conditional_format('I7:I100', {
        'type': 'cell', 'criteria': '>', 'value': target_acos,
        'format': red_fmt
    })
    # ACOS > 0 và ≤ target → xanh
    worksheet.conditional_format('I7:I100', {
        'type': 'formula',
        'criteria': f'=AND($I7>{0}, $I7<={target_acos})',
        'format': green_fmt
    })
    # CTR < 0.3% → vàng
    worksheet.conditional_format('D7:D100', {
        'type': 'cell', 'criteria': '<', 'value': 0.003,
        'format': yellow_fmt
    })
