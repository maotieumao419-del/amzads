# UI/sheets/sheet_keywords.py
"""
Sheet 🔑 KEYWORDS — Hiệu suất từng keyword/target với Health Tag & Gợi ý.
"""

from config import SHEET_KEYWORDS


def build_keywords_sheet(workbook, writer, df_kw):
    """Tạo sheet KEYWORDS."""
    ws = workbook.add_worksheet(SHEET_KEYWORDS)
    fmts = _get_formats(workbook)

    ws.set_row(0, 28)
    ws.merge_range('A1:P1', '🔑 KEYWORDS — Phân tích từ khóa & Gợi ý tối ưu', fmts['title'])

    col_defs = [
        ('SKU',              12),
        ('Campaign Name',    35),
        ('Loại Campaign',    16),
        ('Target',           30),
        ('Match Type',       10),
        ('Status',           10),
        ('Impressions',      13),
        ('Clicks',           9),
        ('Click-through Rate', 8),
        ('Spend',            12),
        ('Sales',            12),
        ('Orders',           9),
        ('ACOS',             9),
        ('CPC',              9),
        ('Health Tag',       13),
        ('Suggested Action', 16),
    ]

    for ci, (name, width) in enumerate(col_defs):
        ws.set_column(ci, ci, width)
        ws.write(1, ci, name, fmts['header'])

    cols_out = [c[0] for c in col_defs]

    for ri, row in df_kw.sort_values('Spend', ascending=False).reset_index(drop=True).iterrows():
        er = ri + 2

        health    = str(row.get('Health Tag', '🆕 New'))
        suggest   = str(row.get('Suggested Action', '👀 Theo dõi'))
        health_fmt   = _health_fmt(health, fmts)
        suggest_fmt  = _suggest_fmt(suggest, fmts)

        row_vals = {
            'SKU':                 (str(row.get('SKU', '')),                   fmts['cell_text']),
            'Campaign Name':       (str(row.get('Campaign Name', '')),         fmts['cell_text']),
            'Loại Campaign':       (str(row.get('Loại Campaign', '')),         fmts['cell_center']),
            'Target':              (str(row.get('Target', '')),                fmts['cell_text']),
            'Match Type':          (str(row.get('Match Type', '')),            fmts['cell_center']),
            'Status':              (str(row.get('Status', '')),                fmts['cell_center']),
            'Impressions':         (row.get('Impressions', 0),                 fmts['cell_num']),
            'Clicks':              (row.get('Clicks', 0),                      fmts['cell_num']),
            'Click-through Rate':  (row.get('Click-through Rate', 0),         fmts['cell_pct']),
            'Spend':               (row.get('Spend', 0),                       fmts['cell_money']),
            'Sales':               (row.get('Sales', 0),                       fmts['cell_money']),
            'Orders':              (row.get('Orders', 0),                      fmts['cell_num']),
            'ACOS':                (row.get('ACOS', 0),                        fmts['cell_pct']),
            'CPC':                 (row.get('CPC', 0),                         fmts['cell_money']),
            'Health Tag':          (health,                                    health_fmt),
            'Suggested Action':    (suggest,                                   suggest_fmt),
        }

        for ci, col_name in enumerate(cols_out):
            val, fmt = row_vals[col_name]
            if isinstance(val, str):
                ws.write_string(er, ci, val, fmt)
            else:
                ws.write_number(er, ci, float(val), fmt)

    ws.freeze_panes(2, 0)

    # Conditional Formatting ACOS
    from xlsxwriter.utility import xl_col_to_name
    acos_ci = cols_out.index('ACOS')
    acos_letter = xl_col_to_name(acos_ci)
    n_rows = len(df_kw) + 2

    ws.conditional_format(f'{acos_letter}3:{acos_letter}{n_rows}', {
        'type': 'cell', 'criteria': '>', 'value': 0.40,
        'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    })
    ws.conditional_format(f'{acos_letter}3:{acos_letter}{n_rows}', {
        'type': 'formula',
        'criteria': f'=AND({acos_letter}3>0,{acos_letter}3<=0.30)',
        'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    })


def _health_fmt(health, fmts):
    if '⭐' in health:   return fmts['health_star']
    if '🔴' in health:   return fmts['health_bleed']
    if '⚠️' in health:  return fmts['health_watch']
    if '💀' in health:   return fmts['health_dead']
    if '😴' in health:   return fmts['health_sleep']
    return fmts['health_new']


def _suggest_fmt(suggest, fmts):
    if 'Tăng' in suggest:    return fmts['suggest_up']
    if 'Giảm' in suggest:    return fmts['suggest_down']
    if 'dừng' in suggest:    return fmts['suggest_pause']
    return fmts['suggest_monitor']


def _get_formats(workbook):
    return {
        'title': workbook.add_format({
            'bold': True, 'font_size': 13,
            'bg_color': '#232F3E', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter'
        }),
        'header': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#37475A', 'color': '#FEBD69',
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        'cell_text':   workbook.add_format({'border': 1, 'align': 'left',   'valign': 'vcenter'}),
        'cell_center': workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        'cell_num':    workbook.add_format({'border': 1, 'num_format': '#,##0',     'align': 'center'}),
        'cell_money':  workbook.add_format({'border': 1, 'num_format': '$#,##0.00', 'align': 'center'}),
        'cell_pct':    workbook.add_format({'border': 1, 'num_format': '0.00%',     'align': 'center'}),
        # Health
        'health_star':  workbook.add_format({'border': 1, 'align': 'center', 'bold': True,
                                              'bg_color': '#C6EFCE', 'font_color': '#006100'}),
        'health_bleed': workbook.add_format({'border': 1, 'align': 'center', 'bold': True,
                                              'bg_color': '#FFC7CE', 'font_color': '#9C0006'}),
        'health_watch': workbook.add_format({'border': 1, 'align': 'center', 'bold': True,
                                              'bg_color': '#FFEB9C', 'font_color': '#9C6500'}),
        'health_dead':  workbook.add_format({'border': 1, 'align': 'center', 'bold': True,
                                              'bg_color': '#808080', 'font_color': '#FFFFFF'}),
        'health_sleep': workbook.add_format({'border': 1, 'align': 'center',
                                              'bg_color': '#EDEDED', 'font_color': '#666666'}),
        'health_new':   workbook.add_format({'border': 1, 'align': 'center',
                                              'bg_color': '#DDEEFF', 'font_color': '#004080'}),
        # Suggestion
        'suggest_up':      workbook.add_format({'border': 1, 'align': 'center',
                                                'bg_color': '#E2EFDA', 'font_color': '#375623'}),
        'suggest_down':    workbook.add_format({'border': 1, 'align': 'center',
                                                'bg_color': '#FCE4D6', 'font_color': '#833C00'}),
        'suggest_pause':   workbook.add_format({'border': 1, 'align': 'center',
                                                'bg_color': '#FFC7CE', 'font_color': '#9C0006'}),
        'suggest_monitor': workbook.add_format({'border': 1, 'align': 'center',
                                                'bg_color': '#F2F2F2', 'font_color': '#595959'}),
    }
