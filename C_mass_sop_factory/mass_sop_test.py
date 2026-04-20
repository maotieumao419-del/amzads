import os
import re
import pandas as pd
from datetime import datetime
import warnings

# -------------------------------------------------------
# NHIỆM VỤ: Tạo hàng loạt campaign (V3 Enterprise)
# Input:  ../../RULE&TEMPLATE/PPC_NGUYÊN.xlsx (chọn sheet)
# Output: data/output/mass_sop_part_*.xlsx
# -------------------------------------------------------

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
RULE_TEMPLATE    = os.path.normpath(os.path.join(SCRIPT_DIR, '..', 'RULE&TEMPLATE'))
INPUT_FILE       = os.path.join(RULE_TEMPLATE, 'PPC_NGUYÊN.xlsx')
OUTPUT_DIR       = os.path.join(SCRIPT_DIR, 'data', 'output')
CAMPAIGNS_PER_CHUNK = 1428

AMAZON_TEMPLATE_COLUMNS = [
    "Product", "Entity", "Operation", "Campaign Id", "Ad Group Id", "Portfolio Id",
    "Ad Id", "Keyword Id", "Product Targeting Id", "Campaign Name", "Ad Group Name",
    "Start Date", "End Date", "Targeting Type", "State", "Daily Budget", "SKU", "ASIN",
    "Eligibility Status", "Reasons for Ineligibility", "Ad Group Default Bid",
    "Bid", "Keyword Text", "Match Type", "Bidding Strategy", "Placement", "Percentage",
    "Product Targeting Expression"
]

def main():
    print(f"--- MASS SOP FACTORY V3 (FINAL ENTERPRISE) ---")

    if not os.path.exists(INPUT_FILE):
        print(f"Error: Input file not found at '{INPUT_FILE}'")
        print(f"Please ensure PPC_NGUYÊN.xlsx is in the RULE&TEMPLATE folder.")
        return

    print(f"Found input file at: {INPUT_FILE}")

    try:
        xl = pd.ExcelFile(INPUT_FILE)
    except Exception as e:
        print(f"Error reading {INPUT_FILE}: {e}")
        return

    sheets = xl.sheet_names
    print("\n--- AVAILABLE SHEETS ---")
    for i, sheet in enumerate(sheets):
        print(f"[{i + 1}] {sheet}")

    while True:
        try:
            selection   = input(f"\nEnter the number of the sheet to process (1-{len(sheets)}): ").strip()
            sheet_index = int(selection) - 1
            if 0 <= sheet_index < len(sheets):
                selected_sheet = sheets[sheet_index]
                break
            else:
                print("Invalid number. Please try again.")
        except ValueError:
            print("Please enter a valid integer.")

    print("\n--- GLOBAL SETTINGS ---")
    base_prefix = input("Enter Base Campaign Prefix (e.g., DESK-DECOR_BUTTERFLYMOM_KT): ").strip()
    if not base_prefix:
        print("CRITICAL: Prefix is required! Exiting.")
        return

    portfolio_id = input("Portfolio ID (Leave blank for none): ").strip()

    date_suffix = input("Date Suffix [YYYYMMDD] (Leave blank for today's date): ").strip()
    if not date_suffix:
        date_suffix = datetime.now().strftime("%Y%m%d")

    default_budget_input = input("Default Daily Budget (Leave blank for 5): ").strip()
    try:
        default_budget = float(default_budget_input) if default_budget_input else 5.0
    except ValueError:
        default_budget = 5.0

    # Auto-extract SKU from prefix
    parts       = base_prefix.split('_')
    current_sku = None
    for i, part in enumerate(parts):
        if part.upper() in ['KT', 'PT']:
            if i > 0: current_sku = parts[i-1]
            break
    if not current_sku and len(parts) >= 2:
        current_sku = parts[-2] if parts[-1].upper() in ['KT', 'PT'] else parts[-1]
    if not current_sku:
        current_sku = input("Could not auto-detect SKU from prefix. Enter Target SKU manually: ").strip()

    print(f"\n[Settings Applied]\n - Prefix: {base_prefix}\n - SKU: {current_sku}\n - Portfolio: '{portfolio_id}'\n - Date Suffix: {date_suffix}")
    print(f"-> Loading data from sheet: '{selected_sheet}'...")

    df_raw = pd.read_excel(INPUT_FILE, sheet_name=selected_sheet)

    col_note    = next((c for c in df_raw.columns if "GHI CHÚ" in str(c).upper() or "NOTE" in str(c).upper()), None)
    col_keyword = next((c for c in df_raw.columns if "TARGET" in str(c).upper() or "KEYWORD TEXT" in str(c).upper() or "KEYWORD" in str(c).upper() or "KW" in str(c).upper()), None)

    successful_campaigns = 0
    skipped_rows         = 0
    chunk_counter        = 1
    current_chunk_rows   = []

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    def export_chunk(rows, chunk_index):
        if not rows: return
        df_out   = pd.DataFrame(rows, columns=AMAZON_TEMPLATE_COLUMNS).fillna("")
        out_file = os.path.join(OUTPUT_DIR, f"mass_sop_part_{chunk_index}.xlsx")
        with pd.ExcelWriter(out_file, engine='xlsxwriter') as writer:
            df_out.to_excel(writer, index=False, sheet_name='Sponsored Products Campaigns')
            workbook  = writer.book
            worksheet = writer.sheets['Sponsored Products Campaigns']
            string_fmt = workbook.add_format({'num_format': '@'})
            for col_num in range(len(AMAZON_TEMPLATE_COLUMNS)):
                worksheet.set_column(col_num, col_num, 15, string_fmt)
        print(f"   [!] Exported block {chunk_index} -> {len(rows)//7} Campaigns saved to '{out_file}'")

    for idx, row in df_raw.iterrows():
        kw_val = row['Target'] if 'Target' in df_raw.columns else (row[col_keyword] if col_keyword else None)

        if pd.isna(kw_val) or str(kw_val).strip() == "":
            skipped_rows += 1
            continue

        kw_str   = str(kw_val).strip()
        note_val = row[col_note] if col_note else ""
        note_str = str(note_val).strip()

        parsed_match_type = 'exact'
        n_lower = note_str.lower()
        if 'exact'  in n_lower: parsed_match_type = 'exact'
        elif 'broad'  in n_lower: parsed_match_type = 'broad'
        elif 'phrase' in n_lower: parsed_match_type = 'phrase'

        base_bid = 0.50
        floats   = [float(x) for x in re.findall(r'\d+\.\d+', note_str)]
        if floats:
            base_bid = floats[0]
        else:
            b_match = re.search(r'(?i)bid\s*(\d+(?:\.\d+)?)', note_str)
            if b_match:
                base_bid = float(b_match.group(1))
            else:
                raw_nums = re.findall(r'\d+(?:\.\d+)?', note_str)
                base_bid = float(raw_nums[0]) if raw_nums and float(raw_nums[0]) < 10 else 0.50

        placement_pct = None
        pct_match     = re.search(r'(\d+)\s*(?:TRP|T|R|P|%)', note_str, re.IGNORECASE)
        if pct_match:
            placement_pct = int(pct_match.group(1))
        else:
            print(f"Row {idx+2} (KW: {kw_str}): Skipping - No Placement % found in notes.")
            skipped_rows += 1
            continue

        match_type_str   = parsed_match_type.lower()
        campaign_string  = f"{base_prefix}_{match_type_str}_{placement_pct}TRP_{date_suffix}_{kw_str}"
        ad_group_name    = kw_str

        def base_row():
            r = {col: "" for col in AMAZON_TEMPLATE_COLUMNS}
            r["Product"]     = "Sponsored Products"
            r["Operation"]   = "Create"
            r["Campaign Id"] = campaign_string
            r["State"]       = "enabled"
            return r

        r1 = base_row(); r1["Entity"] = "Campaign"; r1["Campaign Name"] = campaign_string
        r1["Portfolio Id"] = portfolio_id; r1["Targeting Type"] = "MANUAL"
        r1["Daily Budget"] = default_budget; r1["Bidding Strategy"] = "Dynamic bids - down only"

        r2 = base_row(); r2["Entity"] = "Bidding Adjustment"; r2["Placement"] = "placement top";    r2["Percentage"] = placement_pct
        r3 = base_row(); r3["Entity"] = "Bidding Adjustment"; r3["Placement"] = "placementProductPage"; r3["Percentage"] = placement_pct
        r4 = base_row(); r4["Entity"] = "Bidding Adjustment"; r4["Placement"] = "placementRestOfSearch"; r4["Percentage"] = placement_pct

        r5 = base_row(); r5["Entity"] = "Ad Group"; r5["Ad Group Id"] = campaign_string
        r5["Ad Group Name"] = ad_group_name; r5["Ad Group Default Bid"] = base_bid

        r6 = base_row(); r6["Entity"] = "Product Ad"; r6["Ad Group Id"] = campaign_string; r6["SKU"] = current_sku

        r7 = base_row(); r7["Entity"] = "Keyword"; r7["Ad Group Id"] = campaign_string
        r7["Keyword Text"] = kw_str; r7["Match Type"] = parsed_match_type; r7["Bid"] = base_bid

        current_chunk_rows.extend([r1, r2, r3, r4, r5, r6, r7])
        successful_campaigns += 1

        if successful_campaigns > 0 and successful_campaigns % CAMPAIGNS_PER_CHUNK == 0:
            export_chunk(current_chunk_rows, chunk_counter)
            current_chunk_rows = []
            chunk_counter += 1

    if current_chunk_rows:
        export_chunk(current_chunk_rows, chunk_counter)

    print("\n" + "=" * 40)
    print("      FINAL SUCCESS SUMMARY")
    print("=" * 40)
    print(f"Successful SKC Campaigns: {successful_campaigns}")
    print(f"Total Generated Rows: {successful_campaigns * 7}")
    print("=" * 40)

if __name__ == "__main__":
    main()
