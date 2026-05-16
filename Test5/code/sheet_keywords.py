def build_keywords_sheet(writer, df_kw):
    cols_kw = ["Campaign ID", "Campaign Name (Informational only)", "Ad Group Name (Informational only)", 
               "Portfolio Name (Informational only)", "Keyword / Target", "Match Type", "State", "Bid", 
               "Lifecycle Phase", "Impressions", "Clicks", "Spend", "Sales", "Orders", 
               "Click-through Rate", "Conversion Rate", "ACOS", "CPC", "ROAS", "Health Tag", "Suggested Action"]
               
    df_kw_out = df_kw[[c for c in cols_kw if c in df_kw.columns]].sort_values("Spend", ascending=False)
    df_kw_out.to_excel(writer, sheet_name='🔑 KEYWORDS', index=False)
    
    ws = writer.sheets['🔑 KEYWORDS']
    ws.set_column('B:E', 25)
    ws.set_column('F:U', 12)
    ws.freeze_panes(1, 0)
