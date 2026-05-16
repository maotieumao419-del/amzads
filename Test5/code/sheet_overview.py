from utils import safe_div

def build_overview_sheet(workbook, writer, df_camp, df_kw, today):
    ws_ov = workbook.add_worksheet('📊 OVERVIEW')
    ws_ov.set_column('B:C', 20)
    
    total_spend = df_camp["Spend"].sum()
    total_sales = df_camp["Sales"].sum()
    total_orders = df_camp["Orders"].sum()
    overall_acos = safe_div(total_spend, total_sales)
    overall_roas = safe_div(total_sales, total_spend)

    kpi_data = [
        ("Report Date", today.strftime("%Y-%m-%d %H:%M")),
        ("Total Campaigns", len(df_camp)),
        ("Active Campaigns", len(df_camp[df_camp['State'] == 'enabled'])),
        ("Total Keywords/Targets", len(df_kw)),
        ("Total Spend", total_spend),
        ("Total Sales", total_sales),
        ("Total Orders", total_orders),
        ("Overall ACOS", overall_acos),
        ("Overall ROAS", overall_roas)
    ]

    fmt_currency = workbook.add_format({'num_format': '$#,##0.00'})
    fmt_percent = workbook.add_format({'num_format': '0.00%'})
    
    ws_ov.write(1, 1, "TỔNG QUAN TÀI KHOẢN", workbook.add_format({'bold': True, 'size': 14}))
    for i, (k, v) in enumerate(kpi_data):
        row = i + 3
        ws_ov.write(row, 1, k, workbook.add_format({'bold': True, 'bg_color': '#f2f2f2', 'border': 1}))
        if 'Spend' in k or 'Sales' in k:
            ws_ov.write(row, 2, v, fmt_currency)
        elif 'ACOS' in k:
            ws_ov.write(row, 2, v, fmt_percent)
        else:
            ws_ov.write(row, 2, v)
