import pandas as pd
import numpy as np
from datetime import datetime
from utils import determine_phase, get_campaign_health, get_keyword_suggestion, get_search_term_suggestion

def process_data(file_path, calendar):
    df_sp = pd.read_excel(file_path, sheet_name="Sponsored Products Campaigns", dtype=str)
    try:
        df_st = pd.read_excel(file_path, sheet_name="SP Search Term Report", dtype=str)
    except:
        df_st = pd.DataFrame()

    num_cols = ["Daily Budget", "Bid", "Impressions", "Clicks", "Spend", "Sales", "Orders", "Units", "Click-through Rate", "Conversion Rate", "ACOS", "CPC", "ROAS", "Percentage"]
    for col in num_cols:
        if col in df_sp.columns:
            df_sp[col] = pd.to_numeric(df_sp[col], errors='coerce').fillna(0)
        if not df_st.empty and col in df_st.columns:
            df_st[col] = pd.to_numeric(df_st[col], errors='coerce').fillna(0)

    today = datetime.now()

    df_camp = df_sp[df_sp["Entity"] == "Campaign"].copy()
    df_kw = df_sp[df_sp["Entity"].isin(["Keyword", "Product Targeting"])].copy()

    df_camp["Campaign Age"] = df_camp["Start Date"].apply(lambda x: (today - pd.to_datetime(str(x), format="%Y%m%d", errors='coerce')).days if pd.notna(x) else 0)
    df_camp["Lifecycle Phase"] = df_camp.apply(lambda row: determine_phase(row["Campaign Name"], row["Start Date"], today, calendar), axis=1)
    df_camp["Health Tag"] = df_camp.apply(lambda row: get_campaign_health(row["ACOS"], row["Orders"], row["Clicks"], row["Impressions"]), axis=1)

    camp_lifecycle = df_camp.set_index("Campaign ID")["Lifecycle Phase"].to_dict()
    df_kw["Lifecycle Phase"] = df_kw["Campaign ID"].map(camp_lifecycle).fillna("Unknown")
    df_kw["Keyword / Target"] = np.where(df_kw["Entity"] == "Keyword", df_kw["Keyword Text"], df_kw["Product Targeting Expression"])
    df_kw["Health Tag"] = df_kw.apply(lambda row: get_campaign_health(row["ACOS"], row["Orders"], row["Clicks"], row["Impressions"]), axis=1)
    df_kw["Suggested Action"] = df_kw.apply(lambda row: get_keyword_suggestion(row["Orders"], row["ACOS"], row["Clicks"]), axis=1)

    if not df_st.empty:
        df_st["Suggestion"] = df_st.apply(lambda row: get_search_term_suggestion(row["Orders"], row["ACOS"], row["Clicks"]), axis=1)

    return df_camp, df_kw, df_st, today
