# UI/sheets/sheet_deep_dive.py
"""
Sheet 🔍 DEEP DIVE — Phân tích sâu theo SKU + Campaign với dropdown tương tác.

Layout:
  Row 1: Tiêu đề
  Row 2: Chọn SKU (B2 - dropdown)
  Row 3: Chọn Campaign (B3 - dropdown)
  Row 4: Chọn Giai đoạn (B4 - Launch / Maintain)
  Row 5: (trống)
  Row 6: KPI nhanh của campaign được chọn
  Row 8: Header bảng metrics theo kỳ ngày
  Row 9+: SUMIFS theo từng date block
  [Chart cột bên phải]
  [Phần dưới: Bảng Keyword FILTER()]
"""

import pandas as pd
from xlsxwriter.utility import xl_col_to_name
from config import SHEET_DEEP_DIVE, SHEET_CAMPAIGNS, SHEET_RAW_KW, TARGET_ACOS


def build_deep_dive_sheet(workbook, writer,
                           df_kw: pd.DataFrame,
                           df_camp: pd.DataFrame,
                           date_blocks: list,
                           skus: list):
    """
    Tạo sheet DEEP DIVE.
    - df_kw: keyword data đầy đủ (dùng để lưu vào RAW_KW và FILTER)
    - df_camp: campaign-level aggregated data
    - date_blocks: list các kỳ ngày, ví dụ ['3004-0405', '0506-1206']
    - skus: list SKU names
    """
    ws = workbook.add_worksheet(SHEET_DEEP_DIVE)
    fmts = _get_formats(workbook)

    ws.set_column('A:A', 22)
    ws.set_column('B:B', 38)
    ws.set_column('C:I', 14)

    # ── 1. Lưu RAW_KW vào sheet ẩn ───────────────────────────────────────
    raw_kw_cols = ['SKU', 'Campaign Name', 'Target', 'Match Type', 'Status',
                   'Impressions', 'Clicks', 'Spend', 'Sales', 'Orders', 'ACOS',
                   'Health Tag', 'Suggested Action']
    df_kw_raw = df_kw[[c for c in raw_kw_cols if c in df_kw.columns]].copy()
    df_kw_raw.to_excel(writer, sheet_name=SHEET_RAW_KW, index=False)
    writer.sheets[SHEET_RAW_KW].hide()

    # ── 2. Lưu danh sách SKU vào cột ẩn trong RAW_KW (workaround 255 chars) ─
    ws_raw_kw = writer.sheets[SHEET_RAW_KW]
    ws_raw_kw.write('ZZ1', 'SKU_LIST')
    for i, sku in enumerate(sorted(skus)):
        ws_raw_kw.write(i + 1, 701, sku)  # cột ZZ = index 701

    sku_range = f"='{SHEET_RAW_KW}'!$ZZ$2:$ZZ${len(skus)+1}"

    # ── 3. Layout Header ──────────────────────────────────────────────────
    ws.set_row(0, 30)
    ws.merge_range('A1:I1', '🔍 DEEP DIVE — Phân tích sâu theo Campaign', fmts['title'])

    # Labels
    ws.write('A2', 'Chọn SKU:', fmts['label'])
    ws.write('A3', 'Chọn Campaign:', fmts['label'])
    ws.write('A4', 'Giai đoạn:', fmts['label'])

    # ── 4. Dropdown SKU (B2) ──────────────────────────────────────────────
    ws.data_validation('B2', {
        'validate': 'list',
        'source': sku_range,
        'input_title': 'Chọn SKU',
        'input_message': 'Chọn SKU cần phân tích'
    })
    first_sku = sorted(skus)[0] if skus else ''
    ws.write('B2', first_sku, fmts['dropdown'])

    # ── 5. Dropdown Campaign (B3) — từ CAMPAIGNS sheet ───────────────────
    # Lấy tất cả campaign name vào cột ZZ+1 của RAW_KW
    camp_names = sorted(df_camp['Campaign Name'].unique().tolist()) if not df_camp.empty else []
    ws_raw_kw.write('AAA1', 'CAMP_LIST')
    for i, cn in enumerate(camp_names):
        ws_raw_kw.write(i + 1, 702, cn)  # cột AAA = 702

    camp_range = f"='{SHEET_RAW_KW}'!$AAA$2:$AAA${len(camp_names)+1}"
    ws.data_validation('B3', {
        'validate': 'list',
        'source': camp_range,
        'input_title': 'Chọn Campaign',
        'input_message': 'Chọn Campaign muốn phân tích sâu'
    })
    first_camp = camp_names[0] if camp_names else ''
    ws.write('B3', first_camp, fmts['dropdown'])

    # ── 6. Dropdown Giai đoạn (B4) ───────────────────────────────────────
    ws.data_validation('B4', {
        'validate': 'list',
        'source': ['Launch', 'Maintain', 'Scale'],
        'input_title': 'Giai đoạn',
        'input_message': 'Chọn giai đoạn vận hành'
    })
    ws.write('B4', 'Maintain', fmts['dropdown'])

    # ── 7. Header bảng metrics theo kỳ ───────────────────────────────────
    ws.set_row(6, 22)
    section_headers = ['Kỳ ngày', 'Spend', 'Sales', 'Orders', 'ACOS', 'ROAS', 'Clicks', 'Impressions']
    for ci, h in enumerate(section_headers):
        ws.write(6, ci, h, fmts['header'])

    # ── 8. Ghi dữ liệu theo từng date block (SUMIFS từ RAW_KW) ───────────
    max_kw_row = len(df_kw_raw) + 1

    # Tìm col letters trong RAW_KW
    raw_headers = df_kw_raw.columns.tolist()
    def _col_letter(col_name):
        if col_name in raw_headers:
            return xl_col_to_name(raw_headers.index(col_name))
        return 'A'

    c_sku   = _col_letter('SKU')
    c_camp  = _col_letter('Campaign Name')
    c_spend = _col_letter('Spend')
    c_sales = _col_letter('Sales')
    c_ord   = _col_letter('Orders')
    c_clk   = _col_letter('Clicks')
    c_imp   = _col_letter('Impressions')

    # Tạo dict: date_block → dict của camp metrics từ df_kw (precomputed)
    # Mỗi date block là 1 dòng trong bảng
    # Dùng static data vì SUMIFS không thể filter theo date block từ wide format trong sheet ẩn
    # Thay vào đó, ta precompute và ghi static values

    # Tổng hợp metrics của từng campaign theo từng date block từ df_kw_ts (nếu có)
    # Ở đây dùng cách đơn giản: ghi giá trị tổng kỳ gần nhất từ df_camp
    # và ghi tất cả date blocks với SUMIFS formula trên RAW_KW

    # Do RAW_KW chỉ có metrics của kỳ mới nhất, ta ghi bảng tóm tắt theo date blocks
    # với formula SUMIFS. Mỗi dòng = 1 date block.
    # Nếu RAW_KW chỉ có kỳ mới nhất, thì chỉ hiện 1 dòng = kỳ đó.

    period_label = date_blocks[-1] if date_blocks else 'Kỳ mới nhất'

    # Row 7 (index 7): dòng dữ liệu duy nhất (kỳ mới nhất từ RAW_KW)
    data_row = 7
    ws.write_string(data_row, 0, f'Ngày {period_label}', fmts['cell_center'])

    base_cond = f"'{SHEET_RAW_KW}'!{c_sku}:{c_sku},$B$2,'{SHEET_RAW_KW}'!{c_camp}:{c_camp},$B$3"

    def _sumifs(col_letter):
        return f"=SUMIFS('{SHEET_RAW_KW}'!{col_letter}:{col_letter},{base_cond})"

    ws.write_formula(data_row, 1, _sumifs(c_spend), fmts['cell_money'])
    ws.write_formula(data_row, 2, _sumifs(c_sales), fmts['cell_money'])
    ws.write_formula(data_row, 3, _sumifs(c_ord),   fmts['cell_num'])
    ws.write_formula(data_row, 4,
        f'=IF(C{data_row+1}=0,0,B{data_row+1}/C{data_row+1})', fmts['cell_pct'])
    ws.write_formula(data_row, 5,
        f'=IF(B{data_row+1}=0,0,C{data_row+1}/B{data_row+1})', fmts['cell_num2'])
    ws.write_formula(data_row, 6, _sumifs(c_clk),   fmts['cell_num'])
    ws.write_formula(data_row, 7, _sumifs(c_imp),   fmts['cell_num'])

    # ── 9. Bar Chart Spend vs Sales ────────────────────────────────────────
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Spend',
        'categories': f"='{SHEET_DEEP_DIVE}'!$A$8:$A${data_row+1}",
        'values':     f"='{SHEET_DEEP_DIVE}'!$B$8:$B${data_row+1}",
        'fill': {'color': '#FF9900'}
    })
    chart.add_series({
        'name': 'Sales',
        'categories': f"='{SHEET_DEEP_DIVE}'!$A$8:$A${data_row+1}",
        'values':     f"='{SHEET_DEEP_DIVE}'!$C$8:$C${data_row+1}",
        'fill': {'color': '#232F3E'}
    })
    chart.set_title({'name': 'Spend vs Sales — Campaign được chọn'})
    chart.set_x_axis({'name': 'Kỳ ngày'})
    chart.set_y_axis({'name': 'USD ($)'})
    ws.insert_chart('K2', chart, {'x_scale': 1.5, 'y_scale': 1.3})

    # ── 10. Bảng Keyword FILTER() ─────────────────────────────────────────
    kw_table_start = data_row + 5   # Đảm bảo không đè biểu đồ
    kw_table_start = max(kw_table_start, 18)

    ws.merge_range(kw_table_start, 0, kw_table_start, 8,
                   'LỊCH SỬ KEYWORD — Campaign được chọn', fmts['section_header'])

    kw_headers = ['SKU', 'Campaign Name', 'Target', 'Match Type', 'Status',
                  'Spend', 'Sales', 'Orders', 'ACOS']
    for ci, h in enumerate(kw_headers):
        ws.write(kw_table_start + 1, ci, h, fmts['header'])

    # Tìm col letters
    c_target  = _col_letter('Target')
    c_mt      = _col_letter('Match Type')
    c_status  = _col_letter('Status')
    c_acos    = _col_letter('ACOS')

    # FILTER formula: lọc theo SKU=$B$2 VÀ Campaign=$B$3
    # Các cột: A=SKU, B=CampName, C=Target, D=MatchType, E=Status, F=Imp, G=Clicks, H=Spend, I=Sales, J=Orders, K=ACOS
    raw_range_end = max_kw_row + 1

    filter_cols = [c_sku, c_camp, c_target, c_mt, c_status, c_spend, c_sales, c_ord, c_acos]
    combined = ','.join([f"'{SHEET_RAW_KW}'!{c}2:{c}{raw_range_end}" for c in filter_cols])

    filter_formula = (
        f"=IFERROR(FILTER(CHOOSE({{1,2,3,4,5,6,7,8,9}},"
        f"'{SHEET_RAW_KW}'!{c_sku}2:{c_sku}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_camp}2:{c_camp}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_target}2:{c_target}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_mt}2:{c_mt}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_status}2:{c_status}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_spend}2:{c_spend}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_sales}2:{c_sales}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_ord}2:{c_ord}{raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_acos}2:{c_acos}{raw_range_end}"
        f"),({{''{SHEET_RAW_KW}''!{c_sku}2:{c_sku}{raw_range_end}=$B$2}}"
        f"*({{''{SHEET_RAW_KW}''!{c_camp}2:{c_camp}{raw_range_end}=$B$3}})"
        f"),\"Không có dữ liệu\"),\"Không có dữ liệu\")"
    )

    # Viết FILTER formula dạng đơn giản hơn cho xlsxwriter
    simple_filter = (
        f"=IFERROR(FILTER("
        f"CHOOSE({{1,2,3,4,5,6,7,8,9}},"
        f"'{SHEET_RAW_KW}'!{c_sku}$2:{c_sku}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_camp}$2:{c_camp}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_target}$2:{c_target}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_mt}$2:{c_mt}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_status}$2:{c_status}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_spend}$2:{c_spend}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_sales}$2:{c_sales}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_ord}$2:{c_ord}${raw_range_end},"
        f"'{SHEET_RAW_KW}'!{c_acos}$2:{c_acos}${raw_range_end}"
        f"),("
        f"'{SHEET_RAW_KW}'!{c_sku}$2:{c_sku}${raw_range_end}=$B$2)*"
        f"('{SHEET_RAW_KW}'!{c_camp}$2:{c_camp}${raw_range_end}=$B$3)"
        f"),\"Không có dữ liệu\")"
    )

    ws.write_dynamic_array_formula(
        kw_table_start + 2, 0,
        kw_table_start + 2, 8,
        simple_filter,
        fmts['cell_text']
    )

    # Conditional Formatting ACOS trong bảng keyword
    ws.conditional_format(
        kw_table_start + 2, 8, kw_table_start + 50, 8,
        {'type': 'cell', 'criteria': '>', 'value': TARGET_ACOS,
         'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})}
    )
    ws.conditional_format(
        kw_table_start + 2, 8, kw_table_start + 50, 8,
        {'type': 'formula',
         'criteria': f'=AND(I{kw_table_start+3}>0,I{kw_table_start+3}<={TARGET_ACOS})',
         'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})}
    )


def _get_formats(workbook):
    return {
        'title': workbook.add_format({
            'bold': True, 'font_size': 13,
            'bg_color': '#232F3E', 'color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter'
        }),
        'section_header': workbook.add_format({
            'bold': True, 'font_size': 11,
            'bg_color': '#FF9900', 'color': '#FFFFFF',
            'align': 'left', 'valign': 'vcenter', 'border': 1
        }),
        'header': workbook.add_format({
            'bold': True, 'border': 1,
            'bg_color': '#37475A', 'color': '#FEBD69',
            'align': 'center', 'valign': 'vcenter'
        }),
        'label': workbook.add_format({
            'bold': True, 'bg_color': '#F2F2F2', 'border': 1,
            'align': 'right', 'valign': 'vcenter'
        }),
        'dropdown': workbook.add_format({
            'border': 2, 'bg_color': '#FFFFCC', 'bold': True,
            'align': 'left', 'valign': 'vcenter'
        }),
        'cell_text':   workbook.add_format({'border': 1, 'align': 'left',   'valign': 'vcenter'}),
        'cell_center': workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        'cell_num':    workbook.add_format({'border': 1, 'num_format': '#,##0',     'align': 'center'}),
        'cell_num2':   workbook.add_format({'border': 1, 'num_format': '#,##0.00',  'align': 'center'}),
        'cell_money':  workbook.add_format({'border': 1, 'num_format': '$#,##0.00', 'align': 'center'}),
        'cell_pct':    workbook.add_format({'border': 1, 'num_format': '0.00%',     'align': 'center'}),
    }
