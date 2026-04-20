import pandas as pd
import json
import os

# -------------------------------------------------------
# NHIỆM VỤ: Rule Engine lõi — đọc rules từ rules.json
#           và áp dụng vào file bulk Amazon
#
# Mặc định:
#   Rules:  F_core_engine/rules.json  (cùng thư mục)
#   Input:  nhập đường dẫn file bulk lúc chạy
#   Output: Final_Upload_Ready.xlsx (cùng thư mục với file bulk)
#
# Để dùng file rules khác: nhập đường dẫn khi được hỏi
# -------------------------------------------------------

SCRIPT_DIR        = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RULES_PATH = os.path.join(SCRIPT_DIR, "rules.json")


def load_rules(json_path):
    """
    Parses the external JSON file containing business logic.
    Handles list of rules directly or a parent 'rules' key.
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if isinstance(data, list):
        return data
    elif isinstance(data, dict) and 'rules' in data:
        return data['rules']
    else:
        return [data]


def evaluate_metric_condition(row_data, condition):
    """
    Evaluates a single metric condition against a row.
    Converts values to float, treating NaN/blank as 0.0.
    """
    col      = condition.get('column')
    operator = condition.get('operator')
    val      = condition.get('value')

    if col not in row_data:
        return False

    try:
        row_val_str = str(row_data[col]).strip() if pd.notna(row_data[col]) else ""
        if row_val_str.lower() == 'nan': row_val_str = ""
        row_val = float(row_val_str) if row_val_str != "" else 0.0
    except (ValueError, TypeError):
        row_val = 0.0

    try:
        target_val = float(val)
    except (ValueError, TypeError):
        target_val = 0.0

    if operator == '==': return row_val == target_val
    if operator == '>=': return row_val >= target_val
    if operator == '<=': return row_val <= target_val
    if operator == '>':  return row_val > target_val
    if operator == '<':  return row_val < target_val
    return False


def process_data(df, rules):
    """
    Applies the JSON rules to the DataFrame row by row.
    Returns a new DataFrame containing ONLY the modified rows.
    """
    modified_indices = set()

    for index, row in df.iterrows():
        current_row_data = row.to_dict()
        row_modified     = False

        for rule in rules:
            # Bỏ qua key _comment nếu có
            conds   = rule.get('conditions', {})
            actions = rule.get('actions', {})

            # --- Condition 1: Entity ---
            target_entity = conds.get('target_entity', '')
            entity_val    = str(current_row_data.get('Entity', '')).strip()
            if entity_val != target_entity:
                continue

            # --- Condition 2: Campaign Name ---
            campaign_contains = conds.get('campaign_contains', '')
            if isinstance(campaign_contains, str) and campaign_contains != "":
                camp_name = str(current_row_data.get('Campaign Name', '')).lower()
                camp_id   = str(current_row_data.get('Campaign Id',   '')).lower()
                if camp_name == 'nan': camp_name = ""
                if camp_id   == 'nan': camp_id   = ""
                if campaign_contains.lower() not in camp_name and \
                   campaign_contains.lower() not in camp_id:
                    continue

            # --- Condition 3: Match Type ---
            match_type_contains = conds.get('match_type_contains', '')
            if isinstance(match_type_contains, str) and match_type_contains != "":
                match_type = str(current_row_data.get('Match Type', '')).lower()
                if match_type == 'nan': match_type = ""
                if match_type_contains.lower() not in match_type:
                    continue

            # --- Condition 4: Placement ---
            placement_type = conds.get('placement_type', '')
            if isinstance(placement_type, str) and placement_type != "":
                placement   = str(current_row_data.get('Placement', '')).lower()
                bidding_adj = str(current_row_data.get('Bidding Adjustment Placement', '')).lower()
                if placement   == 'nan': placement   = ""
                if bidding_adj == 'nan': bidding_adj = ""
                if placement_type.lower() not in placement and \
                   placement_type.lower() not in bidding_adj:
                    continue

            # --- Condition 5: Metric Conditions ---
            metric_conditions = conds.get('metric_conditions', [])
            metrics_matched   = True
            for mc in metric_conditions:
                if not evaluate_metric_condition(current_row_data, mc):
                    metrics_matched = False
                    break
            if not metrics_matched:
                continue

            # --- EXECUTE ACTION ---
            action_type   = actions.get('action_type')
            target_column = actions.get('target_column')
            action_val    = actions.get('value')

            if not target_column:
                continue

            if action_type == 'set_value':
                current_row_data[target_column] = action_val
                row_modified = True

            elif action_type == 'increase_value':
                try:
                    curr_str = str(current_row_data.get(target_column, "")).strip()
                    if curr_str.lower() == 'nan': curr_str = ""
                    current_val = float(curr_str) if curr_str != "" else 0.0
                except (ValueError, TypeError):
                    current_val = 0.0
                try:
                    add_val = float(action_val)
                except (ValueError, TypeError):
                    add_val = 0.0

                new_val = round(current_val + add_val, 2)
                current_row_data[target_column] = str(new_val)
                row_modified = True

        if row_modified:
            current_row_data['Operation'] = 'Update'
            for col, val in current_row_data.items():
                if col not in df.columns:
                    df[col] = None
                df.at[index, col] = val
            modified_indices.add(index)

    return df.loc[list(modified_indices)].copy()


def main():
    print("=" * 55)
    print(" Amazon Ads Automation — Core Engine (JSON-driven)")
    print("=" * 55)

    # ── 1. Chọn file bulk Amazon ──────────────────────────────
    excel_path = input("\nNhập đường dẫn file bulk Amazon (.xlsx): ").strip().strip('"').strip("'")
    if not os.path.exists(excel_path):
        print(f"❌ Không tìm thấy file: '{excel_path}'")
        return

    # ── 2. Chọn file rules JSON ───────────────────────────────
    print(f"\nFile rules mặc định: {DEFAULT_RULES_PATH}")
    json_input = input("Nhập đường dẫn rules.json khác (Ấn Enter để dùng mặc định): ").strip().strip('"').strip("'")

    json_path = json_input if json_input else DEFAULT_RULES_PATH

    if not os.path.exists(json_path):
        print(f"❌ Không tìm thấy file rules: '{json_path}'")
        return

    # ── 3. Load Rules ─────────────────────────────────────────
    print(f"\n[1/3] Đang tải Business Rules từ: {json_path}")
    try:
        rules = load_rules(json_path)
        # Lọc bỏ các entry chỉ có _comment
        rules = [r for r in rules if 'conditions' in r and 'actions' in r]
        print(f"  -> Đã tải {len(rules)} rule(s) thành công.")
    except Exception as e:
        print(f"  -> ❌ Lỗi đọc JSON: {e}")
        return

    # ── 4. Load Bulk Data ────────────────────────────────────
    print(f"\n[2/3] Đang đọc file bulk Amazon...")
    try:
        df = pd.read_excel(excel_path, sheet_name='Sponsored Products Campaigns', dtype=str)
        print(f"  -> Đã đọc sheet 'Sponsored Products Campaigns' — {len(df)} dòng.")
    except Exception as e:
        print(f"  -> ❌ Lỗi đọc Excel: {e}")
        return

    # ── 5. Áp dụng Rules ────────────────────────────────────
    print(f"\n[3/3] Đang đánh giá và áp dụng {len(rules)} rule(s)...")
    final_df = process_data(df, rules)
    print(f"  -> Kết quả: {len(final_df)} dòng bị kích hoạt bởi rules.")

    if final_df.empty:
        print("\n⚠️  Không có rule nào khớp. Không xuất file.")
        return

    # ── 6. Xuất file output ──────────────────────────────────
    output_dir      = os.path.dirname(excel_path)
    output_path     = os.path.join(output_dir, "Final_Upload_Ready.xlsx")

    print(f"\n=> Đang đóng gói và xuất file: '{output_path}'...")
    try:
        final_df = final_df.fillna("").replace("nan", "")

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Sponsored Products Campaigns')
            workbook  = writer.book
            worksheet = writer.sheets['Sponsored Products Campaigns']
            text_fmt  = workbook.add_format({'num_format': '@'})
            for col_num in range(len(final_df.columns)):
                worksheet.set_column(col_num, col_num, 18, text_fmt)

        print(f"✅ Xuất thành công! File sẵn sàng upload: {output_path}")
        print(f"   Tổng số dòng thay đổi: {len(final_df)}")
    except PermissionError:
        print(f"❌ Lỗi cấp quyền: File '{output_path}' đang mở trong Excel. Vui lòng đóng lại.")
    except Exception as e:
        print(f"❌ Lỗi xuất file: {e}")

    print("\n" + "=" * 55)
    print(" HOÀN THÀNH")
    print("=" * 55)


if __name__ == "__main__":
    main()
