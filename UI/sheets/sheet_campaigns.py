# UI/sheets/sheet_campaigns.py
"""
Sheet 🎯 CAMPAIGNS — Chi tiết campaign với Health Tag và Sparklines.
"""

from config import SHEET_CAMPAIGNS


def build_campaigns_sheet(workbook, writer, df_camp) -> object:
    """
    Tạo sheet CAMPAIGNS. Trả về df_camp_out để dùng cho sparklines và Deep Dive.
    """
    ws = workbook.add_worksheet(SHEET_CAMPAIGNS)
    fmts = _get_formats(workbook)

    # ── Columns & Layout ──────────────────────────────────────────────────
    ws.set_row(0, 28)
    ws.merge_range('A1:N1', '🎯 CAMPAIGNS — Hiệu suất & Phân loại chiến dịch', fmts['title'])

    col_defs = [
        ('SKU',           12),
        ('Campaign Name', 38),
        ('Loại Campaign', 16),
        ('Status',        10),
        ('Impressions',   13),
        ('Clicks',        9),
        ('CTR',           8),
        ('Spend',         12),
        ('Sales',         12),
        ('Orders',        9),
        ('ACOS',          9),
        ('ROAS',          9),
        ('CVR',           8),
        ('Health Tag',    13),
    ]

    for ci, (name, width) in enumerate(col_defs):
        ws.set_column(ci, ci, width)
        ws.write(1, ci, name, fmts['header'])

    # ── Data rows ─────────────────────────────────────────────────────────
    cols_out = [c[0] for c in col_defs]

    for ri, row in df_camp.iterrows():
        er = ri + 2   # excel row index (0-based, +2 for title+header)

        health = str(row.get('Health Tag', '🆕 New'))
        health_fmt = _health_fmt(health, fmts)

        vals = {
            'SKU':           (str(row.get('SKU', '')),              fmts['cell_text']),
            'Campaign Name': (str(row.get('Campaign Name', '')),    fmts['cell_text']),
            'Loại Campaign': (str(row.get('Loại Campaign', '')),    fmts['cell_center']),
            'Status':        (str(row.get('Status', '')),           fmts['cell_center']),
            'Impressions':   (row.get('Impressions', 0),            fmts['cell_num']),
            'Clicks':        (row.get('Clicks', 0),                 fmts['cell_num']),
            'CTR':           (row.get('CTR', 0),                    fmts['cell_pct']),
            'Spend':         (row.get('Spend', 0),                  fmts['cell_money']),
            'Sales':         (row.get('Sales', 0),                  fmts['cell_money']),
            'Orders':        (row.get('Orders', 0),                 fmts['cell_num']),
            'ACOS':          (row.get('ACOS', 0),                   fmts['cell_pct']),
            'ROAS':          (row.get('ROAS', 0),                   fmts['cell_num2']),
            'CVR':           (row.get('CVR', 0),                    fmts['cell_pct']),
            'Health Tag':    (health,                               health_fmt),
        }

        for ci, col_name in enumerate(cols_out):
            val, fmt = vals[col_name]
            if isinstance(val, str):
                ws.write_string(er, ci, val, fmt)
            elif isinstance(val, float) or isinstance(val, int):
                ws.write_number(er, ci, float(val), fmt)
            else:
                ws.write(er, ci, val, fmt)

    ws.freeze_panes(2, 0)

    # ── Conditional Formatting trên cột ACOS (cột K = index 10) ──────────
    acos_col = cols_out.index('ACOS')
    from xlsxwriter.utility import xl_col_to_name
    acos_letter = xl_col_to_name(acos_col)
    n_rows = len(df_camp) + 2

    ws.conditional_format(f'{acos_letter}3:{acos_letter}{n_rows}', {
        'type': 'cell', 'criteria': '>', 'value': 0.30,
        'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    })
    ws.conditional_format(f'{acos_letter}3:{acos_letter}{n_rows}', {
        'type': 'formula',
        'criteria': f'=AND({acos_letter}3>0,{acos_letter}3<=0.30)',
        'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    })

    return df_camp.reset_index(drop=True), cols_out


def _health_fmt(health: str, fmts: dict):
    if '⭐' in health:   return fmts['health_star']
    if '🔴' in health:   return fmts['health_bleed']
    if '⚠️' in health:  return fmts['health_watch']
    if '💀' in health:   return fmts['health_dead']
    if '😴' in health:   return fmts['health_sleep']
    return fmts['health_new']


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
        'cell_num2':   workbook.add_format({'border': 1, 'num_format': '#,##0.00',  'align': 'center'}),
        'cell_money':  workbook.add_format({'border': 1, 'num_format': '$#,##0.00', 'align': 'center'}),
        'cell_pct':    workbook.add_format({'border': 1, 'num_format': '0.00%',     'align': 'center'}),
        # Health Tag formats
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
    }
