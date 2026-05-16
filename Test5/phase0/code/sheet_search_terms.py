def build_search_terms_sheet(writer, df_st):
    if df_st.empty:
        return
        
    cols_st = ["Campaign Name (Informational only)", "Ad Group Name (Informational only)", 
               "Keyword Text", "Match Type", "Customer Search Term", "Impressions", "Clicks", 
               "Spend", "Sales", "Orders", "Click-through Rate", "Conversion Rate", "ACOS", "CPC", "ROAS", "Suggestion"]
    df_st_out = df_st[[c for c in cols_st if c in df_st.columns]].sort_values("Spend", ascending=False)
    df_st_out.to_excel(writer, sheet_name='🔍 SEARCH TERMS', index=False)
    
    ws = writer.sheets['🔍 SEARCH TERMS']
    ws.set_column('B:E', 25)
    ws.set_column('F:U', 12)
    ws.freeze_panes(1, 0)
