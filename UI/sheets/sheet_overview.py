# UI/sheets/sheet_overview.py
"""
Sheet 📊 OVERVIEW — Tổng quan hiệu suất toàn tài khoản.
"""

from config import SHEET_OVERVIEW, TARGET_ACOS


def _safe_div(a, b):
    return a / b if b and b > 0 else 0


def build_overview_sheet(workbook, writer, df_camp, df_kw, skus: list, file_name: str):
    """
    Tạo sheet OVERVIEW với KPI tổng quát và breakdown theo SKU.
    """
    ws = workbook.add_worksheet(SHEET_OVERVIEW)
    fmts = _get_formats(workbook)

    ws.set_column('A:A', 30)
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 22)
    ws.set_column('D:D', 18)
    ws.set_column('E:I', 14)

    # ── Row 1: Title ──────────────────────────────────────────────────────
    ws.merge_range('A1:I1', '📊 AMAZON ADS — TỔNG QUAN TÀI KHOẢN', fmts['title'])
    ws.set_row(0, 32)

    # ── Row 2: Source file info ───────────────────────────────────────────
    ws.write('A2', f'Nguồn dữ liệu: {file_name}', fmts['subtitle'])
    ws.merge_range('A2:I2', f'Nguồn dữ liệu: {file_name}', fmts['subtitle'])
    ws.set_row(1, 20)

    # ── Row 4-12: KPI Cards ────────────────────────────────────────────────
    total_spend  = df_camp['Spend'].sum()
    total_sales  = df_camp['Sales'].sum()
    total_orders = df_camp['Orders'].sum()
    total_clicks = df_camp['Clicks'].sum()
    total_impr   = df_camp['Impressions'].sum()
    overall_acos = _safe_div(total_spend, total_sales)
    overall_roas = _safe_div(total_sales, total_spend)
    overall_ctr  = _safe_div(total_clicks, total_impr)
    overall_cvr  = _safe_div(total_orders, total_clicks)

    active_camps = len(df_camp[df_camp['Status'].str.lower() == 'enable']) if 'Status' in df_camp.columns else len(df_camp)
    total_camps  = len(df_camp)
    total_kw     = len(df_kw)

    ws.write('A4', 'CHỈ SỐ TỔNG QUAN', fmts['section_header'])
    ws.merge_range('A4:I4', 'CHỈ SỐ TỔNG QUAN', fmts['section_header'])

    kpi_rows = [
        ('📅 Số Campaign',           total_camps,  None,         fmts['kpi_num']),
        ('✅ Campaign đang chạy',     active_camps, None,         fmts['kpi_num']),
        ('🔑 Số Keyword/Target',      total_kw,     None,         fmts['kpi_num']),
        ('💰 Tổng Chi phí (Spend)',   total_spend,  None,         fmts['kpi_money']),
        ('💵 Tổng Doanh thu (Sales)', total_sales,  None,         fmts['kpi_money']),
        ('🛒 Tổng Đơn hàng (Orders)', total_orders, None,         fmts['kpi_num']),
        ('📊 ACOS Tổng thể',          overall_acos, TARGET_ACOS,  fmts['kpi_pct']),
        ('📈 ROAS Tổng thể',          overall_roas, None,         fmts['kpi_num2']),
        ('🖱️ CTR Tổng thể',           overall_ctr,  None,         fmts['kpi_pct']),
        ('🔄 CVR Tổng thể',           overall_cvr,  None,         fmts['kpi_pct']),
    ]

    for i, (label, value, benchmark, val_fmt) in enumerate(kpi_rows):
        row = 4 + i   # Excel rows 5–14 (0-indexed: 4–13)
        ws.write(row, 0, label, fmts['kpi_label'])
        ws.write(row, 1, value, val_fmt)
        if benchmark is not None:
            status = '✅ Đạt' if value <= benchmark else '⚠️ Vượt ngưỡng'
            status_fmt = fmts['status_ok'] if value <= benchmark else fmts['status_warn']
            ws.write(row, 2, f'Target: {benchmark:.0%}', fmts['kpi_label'])
            ws.write(row, 3, status, status_fmt)

    # ── Row 17+: Breakdown theo SKU ───────────────────────────────────────
    ws.write(15, 0, 'BREAKDOWN THEO SKU', fmts['section_header'])
    ws.merge_range(15, 0, 15, 8, 'BREAKDOWN THEO SKU', fmts['section_header'])

    sku_headers = ['SKU', 'Số Campaign', 'Spend', 'Sales', 'Orders', 'ACOS', 'ROAS', 'Health ⭐', 'Health 🔴']
    for ci, h in enumerate(sku_headers):
        ws.write(16, ci, h, fmts['header'])

    if not df_camp.empty and 'SKU' in df_camp.columns:
        sku_grp = df_camp.groupby('SKU', as_index=False).agg(
            n_camp=('Campaign Name', 'count'),
            Spend=('Spend', 'sum'),
            Sales=('Sales', 'sum'),
            Orders=('Orders', 'sum'),
        )
        sku_grp['ACOS'] = sku_grp.apply(
            lambda r: _safe_div(r['Spend'], r['Sales']), axis=1)
        sku_grp['ROAS'] = sku_grp.apply(
            lambda r: _safe_div(r['Sales'], r['Spend']), axis=1)

        # Count health tags per SKU
        if 'Health Tag' in df_camp.columns:
            star_count = df_camp[df_camp['Health Tag'] == '⭐ Star'].groupby('SKU').size().to_dict()
            bleed_count = df_camp[df_camp['Health Tag'] == '🔴 Bleeder'].groupby('SKU').size().to_dict()
        else:
            star_count = bleed_count = {}

        for ri, row in sku_grp.sort_values('Spend', ascending=False).iterrows():
            er = 17 + ri
            ws.write(er, 0, row['SKU'],    fmts['cell_text'])
            ws.write(er, 1, row['n_camp'], fmts['cell_num'])
            ws.write(er, 2, row['Spend'],  fmts['cell_money'])
            ws.write(er, 3, row['Sales'],  fmts['cell_money'])
            ws.write(er, 4, row['Orders'], fmts['cell_num'])
            ws.write(er, 5, row['ACOS'],   fmts['cell_pct'])
            ws.write(er, 6, row['ROAS'],   fmts['cell_num2'])
            ws.write(er, 7, star_count.get(row['SKU'], 0),  fmts['cell_num'])
            ws.write(er, 8, bleed_count.get(row['SKU'], 0), fmts['cell_num'])


def _get_formats(workbook):
    return {
        'title': workbook.add_format({
            'bold': True, 'font_size': 16,
            'bg_color': '#232F3E', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter'
        }),
        'subtitle': workbook.add_format({
            'bold': False, 'font_size': 10, 'italic': True,
            'bg_color': '#37475A', 'color': '#FEBD69',
            'align': 'left', 'valign': 'vcenter'
        }),
        'section_header': workbook.add_format({
            'bold': True, 'font_size': 11,
            'bg_color': '#FF9900', 'color': '#FFFFFF',
            'align': 'left', 'valign': 'vcenter', 'border': 1
        }),
        'header': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#232F3E', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter'
        }),
        'kpi_label': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#F2F2F2', 'align': 'left', 'valign': 'vcenter'
        }),
        'kpi_num':   workbook.add_format({'border': 1, 'num_format': '#,##0',      'align': 'center', 'bold': True}),
        'kpi_num2':  workbook.add_format({'border': 1, 'num_format': '#,##0.00',   'align': 'center', 'bold': True}),
        'kpi_money': workbook.add_format({'border': 1, 'num_format': '$#,##0.00',  'align': 'center', 'bold': True}),
        'kpi_pct':   workbook.add_format({'border': 1, 'num_format': '0.00%',      'align': 'center', 'bold': True}),
        'status_ok':   workbook.add_format({'border': 1, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'bold': True}),
        'status_warn': workbook.add_format({'border': 1, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'bold': True}),
        'cell_text':  workbook.add_format({'border': 1, 'align': 'left'}),
        'cell_num':   workbook.add_format({'border': 1, 'num_format': '#,##0',     'align': 'center'}),
        'cell_num2':  workbook.add_format({'border': 1, 'num_format': '#,##0.00',  'align': 'center'}),
        'cell_money': workbook.add_format({'border': 1, 'num_format': '$#,##0.00', 'align': 'center'}),
        'cell_pct':   workbook.add_format({'border': 1, 'num_format': '0.00%',     'align': 'center'}),
    }
