import pandas as pd
import os
import sys
import datetime
from tabulate import tabulate

# -------------------------------------------------------
# NHIỆM VỤ: Update Sponsored Product campaigns qua CLI
# Input:  data/input/report.xlsx  (hoặc raw_sp_date.csv)
# Output: data/output/upload_update_manual.xlsx
#         data/output/history_log.txt
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

def get_float_input(prompt_msg):
    """Vòng lặp lấy input số thực từ người dùng (Strict Validation)."""
    while True:
        val = input(prompt_msg).strip()
        try:
            return str(float(val))
        except ValueError:
            print("❌ Lỗi định dạng! Vui lòng nhập số.")

def get_menu_selection(prompt_title, options):
    """Hiển thị menu và trả về string tương ứng với lựa chọn theo index."""
    while True:
        print(f"\n{prompt_title}")
        for key, val in options.items():
            print(f"[{key}] {val}")
        user_input = input("👉 Chọn số: ").strip()
        try:
            choice = int(user_input)
            if choice in options:
                return options[choice]
            else:
                print("❌ Lựa chọn không hợp lệ, vui lòng chọn lại số có trong danh sách.")
        except ValueError:
            print("❌ Lựa chọn không hợp lệ, vui lòng chọn lại số có trong danh sách.")

def main():
    print("=" * 60)
    print(" AMAZON BULKSHEET UPDATE - INTERACTIVE CLI OMEGA ")
    print("=" * 60)

    file_path = os.path.join(SCRIPT_DIR, "data", "input", "report.xlsx")
    csv_path  = os.path.join(SCRIPT_DIR, "data", "input", "raw_sp_date.csv")

    print("Đang tải dữ liệu...")
    if os.path.exists(file_path):
        df = pd.read_excel(file_path, sheet_name='Sponsored Products Campaigns', dtype=str)
    elif os.path.exists(csv_path):
        df = pd.read_csv(csv_path, dtype=str)
    else:
        print("❌ Lỗi: Không tìm thấy file dữ liệu (report.xlsx hoặc raw_sp_date.csv) trong data/input/")
        return

    df.fillna("", inplace=True)
    if 'Entity' in df.columns:
        df['Entity'] = df['Entity'].str.strip().str.lower()

    queued_operations = []

    while True:
        keyword_text = input("👉 Nhập Keyword Text cần tìm: ").strip()
        match_type   = get_menu_selection("👉 Chọn loại đối sánh (Match Type):", {1: 'exact', 2: 'phrase', 3: 'broad'})

        mask = (df['Entity'] == 'keyword') & (df['Keyword Text'].str.strip().str.lower() == keyword_text.lower())
        if match_type:
            mask = mask & (df['Match Type'].str.strip().str.lower() == match_type)

        matches = df[mask]
        if matches.empty:
            print("❌ Không tìm thấy keyword nào phù hợp.")
            continue

        table_data    = []
        match_records = []
        stt = 1

        print("\nĐang xử lý Reverse SKU Lookup...")
        for idx, row in matches.iterrows():
            ad_group_id = row.get('Ad Group ID', '')
            sku = "UNKNOWN"
            if ad_group_id:
                product_ads = df[(df['Entity'] == 'product ad') & (df['Ad Group ID'] == ad_group_id)]
                if not product_ads.empty:
                    found_sku = product_ads.iloc[0].get('SKU', '').strip()
                    if found_sku: sku = found_sku

            table_data.append([stt, sku, row.get('Campaign Name', ''), row.get('Ad Group Name', ''),
                               row.get('Bid', ''), row.get('Keyword Text', ''), row.get('Match Type', ''),
                               row.get('Campaign ID', ''), ad_group_id, row.get('Keyword ID', '')])
            match_records.append({'sku': sku, 'row': row})
            stt += 1

        headers = ["STT", "SKU", "Campaign Name", "Ad Group Name", "Bid", "Keyword",
                   "Match Type", "Campaign ID", "Ad Group ID", "Keyword ID"]
        print("\n" + "=" * 100)
        print(" KẾT QUẢ TÌM KIẾM KEYWORD ")
        print("=" * 100)
        print(tabulate(table_data, headers=headers, tablefmt="fancy_grid"))

        while True:
            user_input = input("\n👉 Nhập STT để thao tác | 'S' để tìm keyword khác | '0' để Dừng & Xuất file: ").strip()
            if user_input.lower() == 's':
                selected_stt = 's'
                break
            try:
                selected_stt = int(user_input)
                if selected_stt == 0:
                    break
                if 1 <= selected_stt <= len(table_data):
                    break
                else:
                    print("❌ Lỗi: STT không hợp lệ.")
            except ValueError:
                print("❌ Lỗi định dạng! Vui lòng nhập số, 'S' hoặc '0'.")

        if selected_stt == 's': continue
        if selected_stt == 0:   break

        selected_record    = match_records[selected_stt - 1]
        selected_row       = selected_record['row']
        selected_sku       = selected_record['sku']
        locked_campaign_id = selected_row.get('Campaign ID', '')
        locked_ad_group_id = selected_row.get('Ad Group ID', '')
        locked_keyword_id  = selected_row.get('Keyword ID', '')
        locked_campaign_name = selected_row.get('Campaign Name', '').strip()

        if not locked_campaign_name and locked_campaign_id:
            c_row = df[(df['Entity'] == 'campaign') & (df['Campaign ID'] == locked_campaign_id)]
            if not c_row.empty:
                locked_campaign_name = c_row.iloc[0].get('Campaign Name', '').strip()

        locked_keyword_text = selected_row.get('Keyword Text', '')
        print(f"\n✅ Đã khóa mục tiêu: Keyword '{locked_keyword_text}' (SKU: {selected_sku}) thuộc Campaign: {locked_campaign_name}")

        available_actions = [1, 2, 3, 4, 5, 6, 7, 8, 9]
        action_mapping = {
            1: "Change Campaign Name",
            2: "Update Campaign End Date, Budget, and Bidding Strategy",
            3: "Archive a Campaign",
            4: "Change Ad Group Name and Default Bid",
            5: "Archive an Ad Group",
            6: "Create a new Ad Group in an existing Campaign",
            7: "Pause a Keyword",
            8: "Archive a Keyword",
            9: "Add a new Keyword to an existing Campaign"
        }

        while True:
            if len(available_actions) == 0:
                print("\n🎉 Tất cả các loại cập nhật đã được thực hiện!")
                break

            print("\n" + "=" * 60)
            print(" DANH SÁCH THAO TÁC (CẬP NHẬT) ")
            print("=" * 60)
            print("0: Quay lại menu tìm kiếm từ khóa")
            for action_id in available_actions:
                print(f"{action_id}: {action_mapping[action_id]}")
            print("=" * 60)

            user_input = input("👉 Chọn các loại Update (nhiều số cách nhau bởi dấu phẩy) hoặc 0 để quay lại: ").strip()
            if not user_input: continue

            try:
                chosen_actions = [int(x.strip()) for x in user_input.split(',')]
            except ValueError:
                print("❌ Lỗi định dạng! Vui lòng nhập các số cách nhau bởi dấu phẩy.")
                continue

            if 0 in chosen_actions: break

            for current_action in chosen_actions:
                if current_action not in available_actions:
                    print(f"⚠️ Hành động số {current_action} đã được thực hiện hoặc không hợp lệ, bỏ qua!")
                    continue

                print(f"\n--- THỰC THI HÀNH ĐỘNG SỐ {current_action} ---")
                row_data = {col: "" for col in AMAZON_TEMPLATE_COLUMNS}
                row_data['Product'] = "Sponsored Products"

                if current_action == 1:
                    row_data['Entity'] = "Campaign"; row_data['Operation'] = "Update"
                    row_data['Campaign ID'] = locked_campaign_id
                    row_data['Campaign Name'] = input("Enter New Campaign Name: ").strip()
                    row_data['State'] = get_menu_selection("👉 Chọn trạng thái (State):", {1: 'enabled', 2: 'paused', 3: 'archived'})

                elif current_action == 2:
                    row_data['Entity'] = "Campaign"; row_data['Operation'] = "Update"
                    row_data['Campaign ID'] = locked_campaign_id
                    row_data['End Date'] = input("Enter New End Date (YYYYMMDD): ").strip()
                    row_data['Daily Budget'] = get_float_input("Enter New Daily Budget (e.g., 20.0): ")
                    row_data['Bidding Strategy'] = get_menu_selection("👉 Chọn Chiến lược giá thầu:", {1: 'Dynamic bids - down only', 2: 'Dynamic bids - up and down', 3: 'Fixed bid'})

                elif current_action == 3:
                    row_data['Entity'] = "Campaign"; row_data['Operation'] = "Archive"
                    row_data['Campaign ID'] = locked_campaign_id

                elif current_action == 4:
                    row_data['Entity'] = "Ad group"; row_data['Operation'] = "Update"
                    row_data['Campaign ID'] = locked_campaign_id; row_data['Ad Group ID'] = locked_ad_group_id
                    row_data['Ad Group Name'] = input("Enter New Ad Group Name: ").strip()
                    row_data['Ad Group Default Bid'] = get_float_input("Enter New Default Bid (e.g., 1.5): ")

                elif current_action == 5:
                    row_data['Entity'] = "Ad group"; row_data['Operation'] = "Archive"
                    row_data['Campaign ID'] = locked_campaign_id; row_data['Ad Group ID'] = locked_ad_group_id

                elif current_action == 6:
                    row_data['Entity'] = "Ad group"; row_data['Operation'] = "Create"
                    row_data['Campaign ID'] = locked_campaign_id
                    new_ag_name = input("Enter New Ad Group Name: ").strip()
                    row_data['Ad Group ID'] = new_ag_name; row_data['Ad Group Name'] = new_ag_name
                    row_data['Ad Group Default Bid'] = get_float_input("Enter Default Bid (e.g., 1.0): ")
                    row_data['State'] = "enabled"

                elif current_action == 7:
                    row_data['Entity'] = "Keyword"; row_data['Operation'] = "Update"
                    row_data['Campaign ID'] = locked_campaign_id; row_data['Ad Group ID'] = locked_ad_group_id
                    row_data['Keyword ID'] = locked_keyword_id; row_data['State'] = "paused"

                elif current_action == 8:
                    row_data['Entity'] = "Keyword"; row_data['Operation'] = "Archive"
                    row_data['Campaign ID'] = locked_campaign_id; row_data['Ad Group ID'] = locked_ad_group_id
                    row_data['Keyword ID'] = locked_keyword_id

                elif current_action == 9:
                    row_data['Entity'] = "Keyword"; row_data['Operation'] = "Create"
                    row_data['Campaign ID'] = locked_campaign_id; row_data['Ad Group ID'] = locked_ad_group_id
                    row_data['Keyword Text'] = input("Enter New Keyword Text: ").strip()
                    row_data['Match Type']   = get_menu_selection("👉 Chọn loại đối sánh:", {1: 'exact', 2: 'phrase', 3: 'broad'})
                    bid_input = input("Enter Keyword Bid (press Enter to inherit): ").strip()
                    if bid_input:
                        try:
                            row_data['Bid'] = str(float(bid_input))
                        except ValueError:
                            print("⚠️  Bid không hợp lệ, bỏ qua — sẽ kế thừa từ Ad Group Default Bid.")
                    row_data['State'] = "enabled"

                ENTITY_FIELDS = {
                    'campaign':  ['Product', 'Entity', 'Operation', 'Campaign ID', 'Campaign Name',
                                  'Start Date', 'End Date', 'Targeting Type', 'State', 'Daily Budget',
                                  'Bidding Strategy', 'Placement', 'Percentage'],
                    'ad group':  ['Product', 'Entity', 'Operation', 'Campaign ID', 'Ad Group ID',
                                  'Ad Group Name', 'State', 'Ad Group Default Bid'],
                    'keyword':   ['Product', 'Entity', 'Operation', 'Campaign ID', 'Ad Group ID',
                                  'Keyword ID', 'State', 'Keyword Text', 'Match Type', 'Bid',
                                  'Native Language Keyword', 'Native Language Locale'],
                }
                entity_type    = row_data['Entity'].lower()
                allowed_fields = ENTITY_FIELDS.get(entity_type, list(AMAZON_TEMPLATE_COLUMNS))
                clean_row      = {col: "" for col in AMAZON_TEMPLATE_COLUMNS}
                for field in allowed_fields:
                    if field in row_data:
                        clean_row[field] = row_data[field]

                matched_row = None
                for existing_row in queued_operations:
                    if existing_row['Entity'].lower() == entity_type and \
                       existing_row['Operation'].lower() == clean_row['Operation'].lower():
                        if entity_type == 'campaign' and existing_row['Campaign ID'] == clean_row['Campaign ID']:
                            matched_row = existing_row; break
                        elif entity_type == 'ad group' and \
                             existing_row['Campaign ID'] == clean_row['Campaign ID'] and \
                             existing_row['Ad Group ID'] == clean_row['Ad Group ID']:
                            matched_row = existing_row; break
                        elif entity_type == 'keyword' and existing_row['Keyword ID'] == clean_row['Keyword ID']:
                            matched_row = existing_row; break

                if matched_row is not None:
                    for key in allowed_fields:
                        val = clean_row.get(key, "")
                        if str(val).strip() != "":
                            matched_row[key] = val
                else:
                    queued_operations.append(clean_row)

                if current_action not in [6, 9]:
                    available_actions.remove(current_action)
                print("✅ Đã lưu hành động!")

    # ---------- EXPORT ----------
    if not queued_operations:
        print("\n❌ Không có hành động nào được chọn. Thoát chương trình.")
        return

    print(f"\n✅ Đã chọn thực thi {len(queued_operations)} hành động. Đang tiến hành xuất file...")

    df_out          = pd.DataFrame(queued_operations, columns=AMAZON_TEMPLATE_COLUMNS)
    output_filename = os.path.join(SCRIPT_DIR, "data", "output", "upload_update_manual.xlsx")
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)

    print(f"\n[INFO] Đang tạo file Excel: {output_filename}...")
    while True:
        try:
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, index=False, sheet_name='Sponsored Products Campaigns')
                workbook  = writer.book
                worksheet = writer.sheets['Sponsored Products Campaigns']
                text_fmt  = workbook.add_format({'num_format': '@'})
                for col_num in range(len(AMAZON_TEMPLATE_COLUMNS)):
                    worksheet.set_column(col_num, col_num, 18, text_fmt)
            break
        except PermissionError:
            print(f"\n❌ File '{output_filename}' đang mở. Vui lòng đóng Excel lại.")
            input("👉 Đã đóng xong? Nhấn [Enter] để thử lại...")

    log_path = os.path.join(SCRIPT_DIR, "data", "output", "history_log.txt")
    now_str  = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(log_path, 'a', encoding='utf-8') as f:
            for op_row in queued_operations:
                f.write(f"[{now_str}] - SUCCESS: Applied [{op_row['Operation']} {op_row['Entity']}] to Campaign '{locked_campaign_name}'.\n")
        print("\n🎉 [THÀNH CÔNG] File sẵn sàng upload lên Amazon Seller Central.")
        print(f"📝 Đã ghi lịch sử vào: {log_path}")
    except Exception as e:
        print(f"\n✅ Đã tạo file excel, nhưng lỗi khi ghi log: {e}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Đã nhận lệnh ngắt (Ctrl+C). Thoát chương trình an toàn.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Lỗi hệ thống: {e}")
        sys.exit(1)
