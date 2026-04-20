import pandas as pd
import os
import re
from datetime import datetime

# -------------------------------------------------------
# NHIỆM VỤ: Bước 2 của pipeline A_create_campaign
# Đọc raw CSV → phân tích top KW → tạo bulksheet upload
# Input:  data/input/raw_sp_date.csv
# Output: data/output/upload_manual.xlsx
# -------------------------------------------------------

SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
TOP_N_KEYWORDS   = 10

AMAZON_TEMPLATE_COLUMNS = [
    'Product', 'Entity', 'Operation', 'Campaign ID', 'Ad Group ID', 'Portfolio ID',
    'Ad ID', 'Keyword ID', 'Product Targeting ID', 'Campaign Name', 'Ad Group Name',
    'Start Date', 'End Date', 'Targeting Type', 'State', 'Daily Budget', 'SKU',
    'Ad Group Default Bid', 'Bid', 'Keyword Text', 'Native Language Keyword',
    'Native Language Locale', 'Match Type', 'Bidding Strategy', 'Placement',
    'Percentage', 'Product Targeting Expression', 'Audience ID', 'Shopper Cohort Percentage',
    'Shopper Cohort Type'
]


def generate_manual_campaign(top_kw_objects):
    """
    Generates a perfectly formatted Amazon Bulksheet for the top N winning keywords (Batch SKC).
    """
    rows = []
    today_str = datetime.now().strftime("%Y%m%d")

    for idx, kw_obj in enumerate(top_kw_objects, start=1):
        raw_kw        = kw_obj.get('Keyword Text', '')
        sanitized_kw  = re.sub(r'[^a-zA-Z0-9]', '_', raw_kw.lower()).strip('_')
        sanitized_kw  = re.sub(r'_+', '_', sanitized_kw)

        raw_campaign_id = f"TEST_SKC_{idx}_{sanitized_kw}"
        campaign_id     = re.sub(r'[^a-zA-Z0-9_\-\s]', '', raw_campaign_id)
        campaign_id     = ' '.join(campaign_id.split())
        campaign_name   = campaign_id[:128]

        daily_budget_val = 10.0
        daily_budget_str = '10'
        ad_group_id      = f"ADG_{idx}"

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Campaign', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Portfolio ID': kw_obj.get('Portfolio ID', ''),
                     'Campaign Name': campaign_name, 'Start Date': today_str,
                     'Targeting Type': 'MANUAL', 'State': 'enabled',
                     'Daily Budget': daily_budget_str, 'Bidding Strategy': 'Dynamic bids - down only'})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Bidding Adjustment', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'State': 'enabled',
                     'Placement': 'placement top', 'Percentage': '35'})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Ad Group', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Ad Group ID': ad_group_id,
                     'Ad Group Name': ad_group_id, 'State': 'enabled',
                     'Ad Group Default Bid': '0.50'})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Product Ad', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Ad Group ID': ad_group_id,
                     'State': 'enabled', 'SKU': kw_obj.get('SKU', '')})

        try:
            bid_val = float(kw_obj.get('Bid', '1.0'))
        except ValueError:
            bid_val = 1.0
        if bid_val < 0:       bid_val = abs(bid_val)
        if bid_val > daily_budget_val: bid_val = daily_budget_val
        final_bid_str = str(bid_val)

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Keyword', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Ad Group ID': ad_group_id,
                     'State': 'enabled', 'Bid': final_bid_str,
                     'Keyword Text': raw_kw, 'Match Type': 'exact'})

    df_upload = pd.DataFrame(rows, columns=AMAZON_TEMPLATE_COLUMNS).fillna("")

    output_filename = os.path.join(SCRIPT_DIR, 'data', 'output', 'upload_manual.xlsx')
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)

    while True:
        try:
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                df_upload.to_excel(writer, index=False, sheet_name='Sponsored Products Campaigns')
                workbook  = writer.book
                worksheet = writer.sheets['Sponsored Products Campaigns']
                text_fmt  = workbook.add_format({'num_format': '@'})
                for col_num in range(len(df_upload.columns)):
                    worksheet.set_column(col_num, col_num, 15, text_fmt)
            break
        except PermissionError:
            input(f"\n[!] File {output_filename} đang mở. Đóng Excel lại rồi nhấn Enter...")


# =============================================================================
# RULE ENGINE
# =============================================================================

def apply_business_rules(df_kw):
    cr_col  = next((c for c in df_kw.columns if c.lower() in ['conversion rate', '7 day conversion rate', 'cr']), None)
    clk_col = next((c for c in df_kw.columns if c.lower() in ['clicks', 'click']), None)

    if cr_col:
        cr_cleaned = (df_kw[cr_col].astype(str)
                      .str.replace('%', '', regex=False)
                      .str.replace(',', '.', regex=False)
                      .str.strip())
        df_kw['CR_num'] = pd.to_numeric(cr_cleaned, errors='coerce').astype('float64').fillna(0.0)
    else:
        df_kw['CR_num'] = 0.0

    if clk_col:
        clk_cleaned = df_kw[clk_col].astype(str).str.replace(',', '', regex=False).str.strip()
        df_kw['Clicks_num'] = pd.to_numeric(clk_cleaned, errors='coerce').fillna(0).astype(int)
    else:
        df_kw['Clicks_num'] = 0

    df_kw['Final_Score'] = df_kw['CR_num'].astype('float64')
    return df_kw


def get_top_n_keywords(df):
    entity_col = next((c for c in df.columns if c.lower() == 'entity'), None)
    if not entity_col:
        raise ValueError("Column 'Entity' not found in dataset!")

    df_kw = df[df[entity_col].astype(str).str.strip().str.lower() == 'keyword'].copy()
    if df_kw.empty:
        raise ValueError("No rows with Entity='Keyword' found. Aborting.")

    df_scored   = apply_business_rules(df_kw)
    df_filtered = df_scored[(df_scored['CR_num'] > 0) & (df_scored['Clicks_num'] > 0)]
    df_sorted   = df_filtered.sort_values(by=['CR_num', 'Clicks_num'], ascending=False)
    top_n_df    = df_sorted.head(TOP_N_KEYWORDS)

    kw_col = next((c for c in top_n_df.columns if 'keyword text' in c.lower()), 'Keyword Text')

    print(f"\n--- TOP {TOP_N_KEYWORDS} KEYWORDS (sorted by Final Score) ---")
    for _, row in top_n_df.iterrows():
        cr_pct = f"{round(float(row['CR_num']) * 100, 2)}%"
        print(f"  KW: {str(row.get(kw_col,'')):<40} | CR: {cr_pct:<8} | Clicks: {int(row['Clicks_num'])}")
    print("-" * 60 + "\n")

    return [row for _, row in top_n_df.iterrows()]


# =============================================================================
# MAIN
# =============================================================================

def main():
    input_file = os.path.join(SCRIPT_DIR, "data", "input", "raw_sp_date.csv")

    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Cannot find input file: {input_file}")

    print(f"Reading {input_file} for analysis...")
    df = pd.read_csv(input_file, dtype=str, keep_default_na=False)
    df.columns = df.columns.str.strip()

    entity_col = next((c for c in df.columns if c.lower() == 'entity'), 'Entity')
    df['entity_normalized'] = df[entity_col].astype(str).str.strip().str.lower()

    top_kw_rows = get_top_n_keywords(df)
    if not top_kw_rows:
        print("No valid keywords found. Exiting.")
        return

    kw_text_col   = next((c for c in df.columns if 'keyword text' in c.lower()), 'Keyword Text')
    match_col     = next((c for c in df.columns if 'match type'   in c.lower()), 'Match Type')
    bid_col       = next((c for c in df.columns if c.lower() == 'bid'),           'Bid')
    campaign_col  = next((c for c in df.columns if c.lower() == 'campaign id'),   'Campaign ID')
    ad_group_col  = next((c for c in df.columns if c.lower() == 'ad group id'),   'Ad Group ID')
    portfolio_col = next((c for c in df.columns if c.lower() == 'portfolio id'),  'Portfolio ID')
    sku_col       = next((c for c in df.columns if c.lower() == 'sku'),           'SKU')

    top_kw_objects = []

    print("\n" + "=" * 65)
    print("  [LINEAGE TRACING] REVERSE LOOKUP FOR PRODUCT SKUs")
    print("=" * 65)

    for row in top_kw_rows:
        kw_text          = str(row.get(kw_text_col, '')).strip()
        match_val        = str(row.get(match_col,   'exact')).strip().lower() or 'exact'
        bid_val          = str(row.get(bid_col,     '1.0')).strip() or '1.0'
        winning_campaign = str(row.get(campaign_col, '')).strip()
        winning_ad_group = str(row.get(ad_group_col, '')).strip()

        campaign_rows = df[
            (df['entity_normalized'] == 'campaign') &
            (df[campaign_col].astype(str).str.strip() == winning_campaign)
        ]
        portfolio_id = ""
        if not campaign_rows.empty:
            val = str(campaign_rows.iloc[0].get(portfolio_col, '')).strip()
            if val and val.lower() != 'nan':
                portfolio_id = val

        product_ad_rows = df[
            (df['entity_normalized'] == 'product ad') &
            (df[ad_group_col].astype(str).str.strip() == winning_ad_group)
        ]
        if product_ad_rows.empty:
            print(f"  [WARNING] Skipping '{kw_text}' — No Product Ad for Ad Group ID {winning_ad_group}")
            continue

        val = str(product_ad_rows.iloc[0].get(sku_col, '')).strip()
        if not val or val.lower() == 'nan':
            print(f"  [WARNING] Skipping '{kw_text}' — SKU is empty/missing")
            continue

        sku = val
        print(f"  [FOUND] Traced '{kw_text}' -> SKU: {sku} (Portfolio: {portfolio_id})")
        top_kw_objects.append({'Keyword Text': kw_text, 'Match Type': match_val,
                               'Bid': bid_val, 'SKU': sku, 'Portfolio ID': portfolio_id})

    print("=" * 65 + "\n")

    if not top_kw_objects:
        print("No valid keywords with SKUs to generate campaigns for. Exiting.")
        return

    output_path = os.path.join(SCRIPT_DIR, "data", "output", "upload_manual.xlsx")
    print(f"Generating batch SKC campaigns -> {output_path} ...")

    generate_manual_campaign(top_kw_objects)

    print(f"Done. File saved at: {output_path}")
    print(f"Total campaigns generated: {len(top_kw_objects)}")
    print(f"Total expected rows in excel: {len(top_kw_objects) * 5}")


if __name__ == "__main__":
    main()
