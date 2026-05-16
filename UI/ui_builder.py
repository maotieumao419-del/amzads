# ui_builder.py

from config import TARGET_ACOS
from ui_theme import get_ui_formats, apply_conditional_formatting
from ui_charts import draw_profitability_combo_chart

def get_excel_col(index):
    """Chuyển đổi số index của cột (0-based) sang ký tự Excel (0->A, 25->Z, 26->AA)"""
    letter = ''
    while index >= 0:
        letter = chr(index % 26 + 65) + letter
        index = index // 26 - 1
    return letter


def build_interactive_dashboard(workbook, writer, df_merged):
    """
    View Controller: Dựng layout UI, thiết lập Global State và ghi các công thức động.
    """
    # 1. Lưu Data Layer vào sheet ẩn
    raw_sheet = 'RAW_SKU_DATA'
    df_merged.to_excel(writer, sheet_name=raw_sheet, index=False)
    ws_raw = writer.sheets[raw_sheet]
    ws_raw.hide()

    # 2. Setup UI View
    ui_sheet = 'Dashboard'
    ws_ui = workbook.add_worksheet(ui_sheet)
    formats = get_ui_formats(workbook)

    # Chỉnh độ rộng cột
    ws_ui.set_column('A:A', 35) # Tên Campaign
    ws_ui.set_column('B:I', 15) # Các cột Metric

    # 3. Khởi tạo Global State 1 (Dropdown Chọn SKU) ở dòng 1
    ws_ui.write('A1', 'CHỌN SKU:', formats['title'])
    skus = sorted(df_merged['SKU'].astype(str).unique().tolist())
    if not skus:
        skus = ["No Data"]
        
    # FIX LỖI 255 KÝ TỰ CỦA EXCEL DATA VALIDATION
    # Ghi danh sách SKU vào cột ZZ của sheet RAW_SKU_DATA (cột ẩn cực xa)
    ws_raw.write('ZZ1', 'SKU_LIST')
    for i, sku in enumerate(skus):
        ws_raw.write(f'ZZ{i+2}', sku)
        
    # Tham chiếu Data Validation tới vùng vừa ghi
    sku_range = f"='{raw_sheet}'!$ZZ$2:$ZZ${len(skus)+1}"
    ws_ui.data_validation('B1', {'validate': 'list', 'source': sku_range})
    ws_ui.write('B1', skus[0], formats['dropdown'])

    # 4. Khởi tạo Global State 2 (Dropdown Giai đoạn) ở dòng 2
    ws_ui.write('A2', 'GIAI ĐOẠN:', formats['title'])
    ws_ui.data_validation('B2', {'validate': 'list', 'source': ['Launch', 'Maintain']})
    ws_ui.write('B2', 'Maintain', formats['dropdown'])

    # Sinh Dynamic Mapping từ DataFrame headers
    col_map = {str(col).strip(): get_excel_col(i) for i, col in enumerate(df_merged.columns)}
    c_sku = col_map.get('SKU', 'A')
    c_camp = col_map.get('Campaign Name', col_map.get('Campaign', 'B'))
    c_imp = col_map.get('Impressions', 'C')
    c_clk = col_map.get('Clicks', 'D')
    c_spd = col_map.get('Spend', 'E')
    c_sls = col_map.get('Sales', 'F')
    c_ord = col_map.get('Orders', 'G')

    # 5. Dòng 4: Global Metrics (TỔNG QUAN)
    ws_ui.write('A4', 'TỔNG QUAN', formats['title'])
    
    # Hàm tính tổng quát cho Global Metrics
    def get_sumifs(col_letter):
        return f"=SUMIFS('{raw_sheet}'!{col_letter}:{col_letter}, '{raw_sheet}'!{c_sku}:{c_sku}, $B$1)"
        
    # Headers cho Global Metrics (ghi ở dòng 3 - index 2)
    ws_ui.write('B3', 'Spend', formats['header'])
    ws_ui.write('C3', 'Sales', formats['header'])
    ws_ui.write('D3', 'Clicks', formats['header'])
    ws_ui.write('E3', 'Impressions', formats['header'])
    ws_ui.write('F3', 'ACOS', formats['header'])
    ws_ui.write('G3', 'CVR', formats['header'])
    ws_ui.write('H3', 'CTR', formats['header'])

    # Ghi công thức cho dòng 4 (index 3)
    ws_ui.write_formula('B4', get_sumifs(c_spd), formats['metric_currency'])
    ws_ui.write_formula('C4', get_sumifs(c_sls), formats['metric_currency'])
    ws_ui.write_formula('D4', get_sumifs(c_clk), formats['metric_num'])
    ws_ui.write_formula('E4', get_sumifs(c_imp), formats['metric_num'])
    ws_ui.write_formula('F4', f"=IF(C4=0, 0, B4/C4)", formats['metric_percent'])
    ws_ui.write_formula('G4', f"=IF(D4=0, 0, {get_sumifs(c_ord)}/D4)", formats['metric_percent'])
    ws_ui.write_formula('H4', f"=IF(E4=0, 0, D4/E4)", formats['metric_percent'])

    # 6. Dòng 6: Header Bảng Chi Tiết (index 5)
    headers = ['Campaign', 'Impressions', 'Clicks', 'CTR', 'Spend', 'Sales', 'Orders', 'CVR', 'ACOS']
    for col_idx, header in enumerate(headers):
        ws_ui.write(5, col_idx, header, formats['header'])

    # 7. Ô A7 (index 6): Kích hoạt Dynamic Spill Array
    formula_campaign = f'=UNIQUE(FILTER(\'{raw_sheet}\'!{c_camp}:{c_camp}, \'{raw_sheet}\'!{c_sku}:{c_sku}=$B$1, ""))'
    ws_ui.write_dynamic_array_formula('A7', formula_campaign)

    # 8. Từ dòng 7 đến 50, viết công thức SUMIFS cho các cột chỉ số
    for row in range(6, 50): # index 6 tương ứng với dòng 7 trong Excel
        excel_row = row + 1
        camp_cell = f'$A${excel_row}'
        
        # Điều kiện SUMIFS: Cột SKU = $B$1, Cột Campaign = Tên Campaign hiện tại
        base_cond = f"'{raw_sheet}'!{c_sku}:{c_sku}, $B$1, '{raw_sheet}'!{c_camp}:{c_camp}, {camp_cell}"
        
        # Ghi các công thức có bọc tên sheet trong dấu nháy đơn
        ws_ui.write_formula(row, 1, f'=IF({camp_cell}="","", SUMIFS(\'{raw_sheet}\'!{c_imp}:{c_imp}, {base_cond}))', formats['metric_num'])
        ws_ui.write_formula(row, 2, f'=IF({camp_cell}="","", SUMIFS(\'{raw_sheet}\'!{c_clk}:{c_clk}, {base_cond}))', formats['metric_num'])
        ws_ui.write_formula(row, 3, f'=IF(OR({camp_cell}="", B{excel_row}=0), 0, C{excel_row}/B{excel_row})', formats['metric_percent'])
        ws_ui.write_formula(row, 4, f'=IF({camp_cell}="","", SUMIFS(\'{raw_sheet}\'!{c_spd}:{c_spd}, {base_cond}))', formats['metric_currency'])
        ws_ui.write_formula(row, 5, f'=IF({camp_cell}="","", SUMIFS(\'{raw_sheet}\'!{c_sls}:{c_sls}, {base_cond}))', formats['metric_currency'])
        ws_ui.write_formula(row, 6, f'=IF({camp_cell}="","", SUMIFS(\'{raw_sheet}\'!{c_ord}:{c_ord}, {base_cond}))', formats['metric_num'])
        ws_ui.write_formula(row, 7, f'=IF(OR({camp_cell}="", C{excel_row}=0), 0, G{excel_row}/C{excel_row})', formats['metric_percent'])
        ws_ui.write_formula(row, 8, f'=IF(OR({camp_cell}="", F{excel_row}=0), 0, E{excel_row}/F{excel_row})', formats['metric_percent'])

    # 9. Áp dụng luật Conditional Formatting
    apply_conditional_formatting(ws_ui, formats, TARGET_ACOS)

    # 10. Gọi Charting Layer để vẽ Combo Chart
    # Dòng bắt đầu là 7 (A7:A21) thay vì 6 như trước
    data_ranges = {
        'categories': f'={ui_sheet}!$A$7:$A$21',
        'spend': f'={ui_sheet}!$E$7:$E$21',
        'sales': f'={ui_sheet}!$F$7:$F$21',
        'acos': f'={ui_sheet}!$I$7:$I$21'
    }
    
    draw_profitability_combo_chart(workbook, ws_ui, 'K5', data_ranges, formats)
