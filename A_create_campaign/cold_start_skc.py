import pandas as pd
import os
import re
from datetime import datetime

# -------------------------------------------------------
# NHIỆM VỤ: Tạo Single Keyword Campaign từ đầu (Cold Start)
# Nhập thủ công: SKU, Portfolio ID, Budget, Keyword list
# Output: data/output/upload_new_test.xlsx
# -------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

AMAZON_TEMPLATE_COLUMNS = [
    'Product', 'Entity', 'Operation', 'Campaign ID', 'Ad Group ID', 'Portfolio ID',
    'Ad ID', 'Keyword ID', 'Product Targeting ID', 'Campaign Name', 'Ad Group Name',
    'Start Date', 'End Date', 'Targeting Type', 'State', 'Daily Budget', 'SKU',
    'Ad Group Default Bid', 'Bid', 'Keyword Text', 'Native Language Keyword',
    'Native Language Locale', 'Match Type', 'Bidding Strategy', 'Placement',
    'Percentage', 'Product Targeting Expression', 'Audience ID', 'Shopper Cohort Percentage',
    'Shopper Cohort Type'
]

def main():
    print("=====================================================")
    print("   COLD START SINGLE KEYWORD CAMPAIGN GENERATOR")
    print("=====================================================")

    print("\nPlease provide the campaign criteria below:")

    sku = input("Target SKU: ").strip()
    while not sku:
        print("[!] Target SKU is required.")
        sku = input("Target SKU: ").strip()

    portfolio_id = input("Portfolio ID (Leave blank to skip): ").strip()

    while True:
        daily_budget_str = input("Daily Budget (e.g., 10): ").strip()
        if not daily_budget_str: daily_budget_str = "10"
        try:
            budget_val = float(daily_budget_str)
            if budget_val <= 0:
                print("[!] Daily Budget must be a positive number.")
                continue
            daily_budget = str(budget_val)
            break
        except ValueError:
            print("[!] Invalid input. Please enter a number for Daily Budget.")

    while True:
        default_bid_str = input("Default Keyword Bid (e.g., 0.75): ").strip()
        if not default_bid_str: default_bid_str = "0.75"
        try:
            bid_val = float(default_bid_str)
            if bid_val < 0:
                print("[!] Default Keyword Bid cannot be negative.")
                continue
            if bid_val > budget_val:
                print(f"[!] Keyword Bid ({bid_val}) cannot be higher than Daily Budget ({budget_val}).")
                continue
            default_bid = str(bid_val)
            break
        except ValueError:
            print("[!] Invalid input. Please enter a number for Default Keyword Bid.")

    keyword_list_str = input("Keyword List (comma-separated): ").strip()
    while not keyword_list_str:
        print("[!] A comma-separated list of keywords is required.")
        keyword_list_str = input("Keyword List (comma-separated): ").strip()

    keywords = [kw.strip() for kw in keyword_list_str.split(',') if kw.strip()]
    if not keywords:
        print("No valid keywords found. Exiting.")
        return

    rows = []
    today_str = datetime.now().strftime("%Y%m%d")

    for kw in keywords:
        sanitized_kw  = re.sub(r'[^a-zA-Z0-9]', '_', kw.lower()).strip('_')
        sanitized_kw  = re.sub(r'_+', '_', sanitized_kw)
        raw_campaign_id = f"TEST_NEW_{sku}_{sanitized_kw}"
        campaign_id   = re.sub(r'[^a-zA-Z0-9_\-\s]', '', raw_campaign_id)
        campaign_id   = ' '.join(campaign_id.split())
        campaign_name = campaign_id[:128]
        ad_group_id   = "ADG_01"

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Campaign', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Portfolio ID': portfolio_id,
                     'Campaign Name': campaign_name, 'Start Date': today_str,
                     'Targeting Type': 'Manual', 'State': 'Enabled',
                     'Daily Budget': daily_budget, 'Bidding Strategy': 'Dynamic bids - down only'})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Bidding Adjustment', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'State': 'Enabled',
                     'Placement': 'Placement Top', 'Percentage': '35'})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Ad Group', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Ad Group ID': ad_group_id,
                     'Ad Group Name': ad_group_id, 'State': 'Enabled',
                     'Ad Group Default Bid': default_bid})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Product Ad', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Ad Group ID': ad_group_id,
                     'State': 'Enabled', 'SKU': sku})

        rows.append({'Product': 'Sponsored Products', 'Entity': 'Keyword', 'Operation': 'Create',
                     'Campaign ID': campaign_id, 'Ad Group ID': ad_group_id,
                     'State': 'Enabled', 'Bid': default_bid,
                     'Keyword Text': kw, 'Match Type': 'exact'})

    df_upload       = pd.DataFrame(rows, columns=AMAZON_TEMPLATE_COLUMNS).fillna("")
    output_filename = os.path.join(SCRIPT_DIR, 'data', 'output', 'upload_new_test.xlsx')
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)

    print("\nGenerating batches...")
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

    print(f"\n[SUCCESS] Exported {len(rows)} rows successfully!")
    print(f"Total Cold Start Campaigns Generated: {len(keywords)}")
    print(f"File Saved: {output_filename}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user. Exiting.")
