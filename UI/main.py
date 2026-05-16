# main.py

import pandas as pd
import glob
import os
from ui_builder import build_interactive_dashboard

def main():
    print("=== BẮT ĐẦU PIPELINE ETL & DASHBOARD GENERATOR ===")
    
    # Định nghĩa thư mục Input chứa các file cập nhật mới nhất
    # Trỏ về folder D_auto_bulk_updater/data/final_xlsx/
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(BASE_DIR, "D_auto_bulk_updater", "data", "final_xlsx")
    
    # 1. Pipeline ETL: Tìm các file khớp pattern "PPC_*_UPDATED"
    pattern = os.path.join(data_dir, "PPC_*_UPDATED*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        print(f"❌ Lỗi: Không tìm thấy file nào khớp pattern tại {pattern}")
        print("Vui lòng đảm bảo các file PPC_..._UPDATED.xlsx đã có trong D:/data/final_xlsx/")
        return
        
    all_data = []
    # Loại bỏ các sheet không chứa dữ liệu SKU thực tế
    exclude_sheets = ['Listing', 'Portfolio ID', 'Sheet1', 'Sheet2'] 
    
    print(f"Tìm thấy {len(files)} file cần xử lý...")
    
    for file_path in files:
        print(f"Đang trích xuất: {os.path.basename(file_path)}")
        try:
            xl = pd.ExcelFile(file_path, engine='openpyxl')
            for sheet_name in xl.sheet_names:
                if sheet_name in exclude_sheets:
                    continue
                    
                # Đọc tạm để tìm header row
                temp_df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
                header_row_idx = None
                
                # Tự động dò tìm dòng chứa 'Campaign Name' hoặc 'STT'
                for idx, row in temp_df.iterrows():
                    row_str = ' '.join([str(val) for val in row.values if pd.notna(val)])
                    if 'Campaign Name' in row_str or 'STT' in row_str:
                        header_row_idx = idx
                        break
                
                if header_row_idx is None:
                    print(f"Bỏ qua sheet {sheet_name}: Không tìm thấy dòng Header.")
                    continue
                
                # Đọc lại với header chính xác
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=header_row_idx)
                
                # Xóa khoảng trắng thừa trong tên cột
                df.columns = df.columns.str.strip()
                
                # LỌC FILE RỖNG
                if 'Spend' not in df.columns or 'Sales' not in df.columns:
                    print(f"Bỏ qua sheet {sheet_name}: Thiếu cột Spend hoặc Sales.")
                    continue
                
                # Ép kiểu dữ liệu (to_numeric) và fill NaN
                numeric_cols = ['Impressions', 'Clicks', 'Spend', 'Sales', 'Orders']
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # Thêm cột SKU vào đầu tiên
                df.insert(0, 'SKU', sheet_name)
                    
                all_data.append(df)
        except Exception as e:
            print(f"❌ Lỗi khi đọc file {file_path}: {e}")
            
    if not all_data:
        print("❌ Không có dữ liệu hợp lệ nào được trích xuất.")
        return
        
    # Gộp toàn bộ dữ liệu thành 1 Master DataFrame (Data Layer)
    df_merged = pd.concat(all_data, ignore_index=True)
    
    # Thực hiện reset_index sau khi concat
    df_merged.reset_index(drop=True, inplace=True)
    
    # Loại bỏ các cột 'Unnamed' hoặc cột chỉ chứa giá trị rỗng
    df_merged.dropna(how='all', axis=1, inplace=True)
    df_merged = df_merged.loc[:, ~df_merged.columns.str.contains('^Unnamed', case=False, na=False)]
    
    # Đảm bảo cột SKU luôn nằm ở vị trí đầu tiên (index 0)
    cols = ['SKU'] + [c for c in df_merged.columns if c != 'SKU']
    df_merged = df_merged[cols]
    
    print(f"✅ Đã gộp thành công {len(df_merged)} dòng dữ liệu từ tất cả các SKU.")
    
    # 2. Xây dựng Dashboard
    # Ghi vào thư mục UI/data/output
    ui_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(ui_dir, "data", "output")
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
        
    out_file = os.path.join(out_dir, "AMZ_Interactive_Dashboard.xlsx")
    print(f"Đang sinh file Dashboard Client-Side Rendering...")
    
    writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
    workbook = writer.book
    
    build_interactive_dashboard(workbook, writer, df_merged)
    
    writer.close()
    
    print(f"=== HOÀN TẤT ===")
    print(f"✅ File Output đã sẵn sàng: {out_file}")
    print(f"Mở file lên và sử dụng Menu Dropdown ở ô B2/B3 để tương tác phân tích dữ liệu.")

if __name__ == "__main__":
    main()
