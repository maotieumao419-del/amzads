def build_campaigns_sheet(writer, df_camp):
    cols_camp = ["Campaign ID", "Campaign Name", "Portfolio Name (Informational only)", "State", "Targeting Type", 
                 "Bidding Strategy", "Daily Budget", "Campaign Age", "Lifecycle Phase", 
                 "Impressions", "Clicks", "Spend", "Sales", "Orders", "Units", 
                 "Click-through Rate", "Conversion Rate", "ACOS", "CPC", "ROAS", "Health Tag"]
    
    df_camp_out = df_camp[[c for c in cols_camp if c in df_camp.columns]].sort_values("Spend", ascending=False)
    df_camp_out.to_excel(writer, sheet_name='🎯 CAMPAIGNS', index=False)
    
    ws = writer.sheets['🎯 CAMPAIGNS']
    ws.set_column('B:E', 25)
    ws.set_column('F:U', 12)
    ws.freeze_panes(1, 0)
    
    return df_camp_out
