import pandas as pd
import os
import sys

# -------------------------------------------------------
# NHIỆM VỤ: Tự động update hàng loạt theo Portfolio (Multi-Account)
# Input:  data/input 1/*.xlsx (mỗi file = 1 tài khoản Amazon)
# Output: data/output/Auto_Upload_<keyword>_<account>.xlsx
# -------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_col(cols, possible_names):
    for n in possible_names:
        if n in cols:
            return n
    return None

def main():
    print("=" * 70)
    print(" 🚀 TOTAL AUTOMATION: DYNAMIC PORTFOLIO UPDATER (MULTI-ACCOUNT) 🚀")
    print("=" * 70)
    print("Hệ thống xử lý độc lập từng luồng file đầu vào để tách biệt tài khoản.\n")

    input_dir  = os.path.join(SCRIPT_DIR, "data", "input 1")
    output_dir = os.path.join(SCRIPT_DIR, "data", "output")

    if not os.path.exists(input_dir):
        print(f"❌ Lỗi: Không tìm thấy thư mục {input_dir}")
        return

    files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
    if not files:
        print(f"❌ Lỗi: Không tìm thấy file gốc excel (.xlsx) nào trong {input_dir}")
        return

    print("[BƯỚC 1] THIẾT LẬP THUẬT TOÁN")
    target_keyword = input("👉 Nhập từ khóa Portfolio cần tìm (Ví dụ: Graduation, Valentine...): ").strip()
    if not target_keyword:
        print("❌ Bạn chưa nhập từ khóa. Hủy thao tác.")
        return

    user_input = input("👉 Nhập ngưỡng Impression chặn cho chiến dịch Exact (Ấn Enter để chọn mức 100): ").strip()
    try:
        imp_threshold = float(user_input) if user_input else 100.0
    except ValueError:
        print("⚠️ Nhập sai định dạng, tự động dùng ngưỡng 100.")
        imp_threshold = 100.0

    safe_keyword = "".join([c if c.isalnum() else "_" for c in target_keyword]).strip("_")

    print(f"\n[BƯỚC 2] BẮT ĐẦU QUÉT {len(files)} FILE TÀI KHOẢN RIÊNG BIỆT...\n")

    for f in files:
        print("+" + "-" * 80 + "+")
        print(f"|  ĐANG XỬ LÝ FILE TÀI KHOẢN: {f}")
        print("+" + "-" * 80 + "+")

        filepath  = os.path.join(input_dir, f)
        base_name = os.path.splitext(f)[0]

        try:
            xl          = pd.ExcelFile(filepath)
            sheet_names = xl.sheet_names
            found_ids   = set()

            if 'Portfolios' in sheet_names:
                df_port    = xl.parse('Portfolios', dtype=str)
                p_name_col = get_col(df_port.columns, ['Portfolio Name'])
                p_id_col   = get_col(df_port.columns, ['Portfolio Id', 'Portfolio ID'])
                if p_name_col and p_id_col:
                    grads = df_port[df_port[p_name_col].astype(str).str.contains(target_keyword, case=False, na=False)]
                    found_ids.update(grads[p_id_col].dropna().tolist())

            for sheet in ['Sponsored Products Campaigns', 'Sponsored Brands Campaigns', 'Sponsored Display Campaigns']:
                if sheet in sheet_names:
                    df_camp    = xl.parse(sheet, dtype=str)
                    entity_col = get_col(df_camp.columns, ['Entity'])
                    p_name_col = get_col(df_camp.columns, ['Portfolio Name', 'Portfolio Name (Informational only)'])
                    p_id_col   = get_col(df_camp.columns, ['Portfolio Id', 'Portfolio ID'])
                    if entity_col and p_name_col and p_id_col:
                        grads = df_camp[(df_camp[entity_col].astype(str).str.lower() == 'portfolio') &
                                        (df_camp[p_name_col].astype(str).str.contains(target_keyword, case=False, na=False))]
                        found_ids.update(grads[p_id_col].dropna().tolist())

            print(f"  -> 🎯 Đã nhận diện {len(found_ids)} Portfolio(s) chứa lệnh điều chỉnh.")
            if len(found_ids) == 0:
                print(f"  -> ⏭️ Bỏ qua File này vì không có Portfolio liên quan.")
                continue

            matching_campaigns_dfs = []
            for sheet in ['Sponsored Products Campaigns', 'Sponsored Brands Campaigns', 'Sponsored Display Campaigns']:
                if sheet in sheet_names:
                    df_sheet   = xl.parse(sheet, dtype=str)
                    entity_col = get_col(df_sheet.columns, ['Entity'])
                    p_id_col   = get_col(df_sheet.columns, ['Portfolio Id', 'Portfolio ID'])
                    if entity_col and p_id_col:
                        camp_rows    = df_sheet[(df_sheet[entity_col].astype(str).str.contains('Campaign', case=False, na=False)) &
                                                (df_sheet[p_id_col].isin(found_ids))]
                        camp_id_col  = get_col(df_sheet.columns, ['Campaign Id', 'Campaign ID'])
                        camp_name_col = get_col(df_sheet.columns, ['Campaign Name', 'Campaign Name (Informational only)'])
                        camp_ids  = set(camp_rows[camp_id_col].dropna())  if camp_id_col  else set()
                        camp_names = set(camp_rows[camp_name_col].dropna()) if camp_name_col else set()
                        mask = pd.Series(False, index=df_sheet.index)
                        if len(camp_ids) > 0 and camp_id_col:
                            mask = mask | df_sheet[camp_id_col].isin(camp_ids)
                        if len(camp_names) > 0 and camp_name_col:
                            mask = mask | df_sheet[camp_name_col].isin(camp_names)
                        matched_rows = df_sheet[mask]
                        if not matched_rows.empty:
                            matching_campaigns_dfs.append(matched_rows)

            if not matching_campaigns_dfs:
                print(f"  -> ⏭️ Portfolio tồn tại nhưng không có dữ liệu để quét. Bỏ qua.")
                continue

            combined_df = pd.concat(matching_campaigns_dfs, ignore_index=True)
            combined_df.fillna("", inplace=True)

            print(f"  -> ⚙️ Đang áp dụng Luật (Broad/Phrase/Exact) riêng cho tài khoản '{base_name}'...")
            updates = []

            cols          = combined_df.columns.tolist()
            camp_name_col = 'Campaign Name' if 'Campaign Name' in cols else ('Campaign Name (Informational only)' if 'Campaign Name (Informational only)' in cols else None)
            camp_id_col   = 'Campaign ID'   if 'Campaign ID'   in cols else ('Campaign Id' if 'Campaign Id' in cols else None)

            camp_id_to_name = {}
            if camp_name_col and camp_id_col:
                for idx2, row2 in combined_df.iterrows():
                    c_name = str(row2.get(camp_name_col, '')).strip()
                    c_id   = str(row2.get(camp_id_col, '')).strip()
                    if c_name and c_name.lower() != 'nan' and c_id and c_id.lower() != 'nan':
                        camp_id_to_name[c_id] = c_name

            for index, row in combined_df.iterrows():
                entity         = str(row.get('Entity', '')).strip().lower()
                c_id           = str(row.get(camp_id_col, '')).strip() if camp_id_col else ""
                camp_name_raw  = str(row.get(camp_name_col, '')).strip() if camp_name_col else ""
                if not camp_name_raw or camp_name_raw.lower() == 'nan':
                    camp_name_raw = camp_id_to_name.get(c_id, "")
                camp_name = camp_name_raw.lower()
                modified  = False
                new_row   = row.copy()

                if 'broad' in camp_name:
                    if entity == 'bidding adjustment':
                        placement = str(row.get('Placement', '')).lower()
                        if 'product page' in placement:
                            new_row['Percentage'] = "0"; modified = True
                elif 'phrase' in camp_name:
                    if entity == 'bidding adjustment':
                        placement = str(row.get('Placement', '')).lower()
                        if 'product page' in placement:
                            new_row['Percentage'] = "10"; modified = True
                elif 'exact' in camp_name or 'ex' in camp_name:
                    if entity == 'keyword':
                        imp_str = str(row.get('Impressions', '')).strip()
                        try:
                            imp_val = float(imp_str) if imp_str else 0.0
                        except ValueError:
                            imp_val = 0.0
                        if imp_val < imp_threshold:
                            bid_str = str(row.get('Bid', '')).strip()
                            if bid_str:
                                try:
                                    new_row['Bid'] = str(round(float(bid_str) + 0.1, 2)); modified = True
                                except ValueError:
                                    pass

                if modified:
                    new_row['Operation'] = "Update"
                    updates.append(new_row)

            if not updates:
                print(f"  -> 🚫 Không có dòng nào vi phạm ngưỡng logic. Bỏ qua tài khoản này.")
                continue

            df_out          = pd.DataFrame(updates)
            os.makedirs(output_dir, exist_ok=True)
            output_filename = os.path.join(output_dir, f"Auto_Upload_{safe_keyword}_{base_name}.xlsx")

            print(f"  -> 📦 Đang đóng gói {len(updates)} chỉnh sửa ra file Excel...")
            try:
                with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                    df_out.to_excel(writer, index=False, sheet_name='Sponsored Products Campaigns')
                    workbook  = writer.book
                    worksheet = writer.sheets['Sponsored Products Campaigns']
                    text_fmt  = workbook.add_format({'num_format': '@'})
                    for col_num in range(len(df_out.columns)):
                        worksheet.set_column(col_num, col_num, 18, text_fmt)
                print(f"  -> ✅ File sẵn sàng Upload: {output_filename}")
            except PermissionError:
                print(f"  -> ❌ Lỗi Cấp Quyền: Không thể ghi đè '{output_filename}'. Vui lòng đóng Excel lại.")

        except Exception as e:
            print(f"  -> ❌ LỖI VĂNG QUÁ TRÌNH TẠI FILE NÀY: {e}")

    print("\n" + "=" * 70)
    print(" 🎉 HOÀN THÀNH - TOÀN BỘ CÁC TÀI KHOẢN ĐÃ ĐƯỢC XỬ LÝ XUYÊN SUỐT 🎉")
    print("=" * 70)

if __name__ == "__main__":
    main()
