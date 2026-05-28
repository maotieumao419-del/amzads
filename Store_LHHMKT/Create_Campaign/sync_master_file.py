import os
import shutil
import pandas as pd
import glob

# ==============================================================================
# ==============================================================================
# CẤU HÌNH ĐƯỜNG DẪN TỆP TIN
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_1_DIR = os.path.join(BASE_DIR, 'data', 'input 1')
INPUT_2_DIR = os.path.join(BASE_DIR, 'data', 'input 2')

OUTPUT_FILE = os.path.join(INPUT_2_DIR, 'PPC_PROCESSED_OUTPUT.xlsx')

def main():
    print(f"\n[INFO] === BẮT ĐẦU ĐỒNG BỘ DỮ LIỆU VÀO INTERNAL FILE ===")
    
    # ── Tìm file PPC_*.xlsx gốc trong input 1 ──────────
    internal_files = glob.glob(os.path.join(INPUT_1_DIR, "PPC_*.xlsx"))
    # Bỏ qua file rác của Excel và file Output
    internal_files = [f for f in internal_files if not os.path.basename(f).startswith("~$") and os.path.basename(f) != "PPC_PROCESSED_OUTPUT.xlsx" and not os.path.basename(f).endswith("_UPDATED.xlsx")]
    
    if not internal_files:
        print(f"[ERROR] Không tìm thấy file Internal (PPC_*.xlsx) tại:\n'{INPUT_1_DIR}'.")
        return
        
    if len(internal_files) > 1:
        internal_files.sort(key=os.path.getmtime, reverse=True)
        print(f"  [WARNING] Nhiều file PPC_*.xlsx → chọn mới nhất: {os.path.basename(internal_files[0])}")
        
    INTERNAL_FILE = internal_files[0]
    
    # Tạo tên file UPDATED dựa trên tên file gốc và lưu vào cùng thư mục với file gốc
    base_name, ext = os.path.splitext(os.path.basename(INTERNAL_FILE))
    UPDATED_INTERNAL_FILE = os.path.join(INPUT_1_DIR, f"{base_name}_UPDATED{ext}")
    
    # Kiểm tra sự tồn tại của file Output
    if not os.path.exists(OUTPUT_FILE):
        print(f"[ERROR] Không tìm thấy file Output tại:\n'{OUTPUT_FILE}'.")
        return

    # -------------------------------------------------------------------------
    # BƯỚC 1: Đọc dữ liệu Output
    # -------------------------------------------------------------------------
    print(f"\n[1/4] Đọc dữ liệu từ file Output...")
    processed_skus = []
    output_sheets_data = {}
    
    try:
        output_excel = pd.ExcelFile(OUTPUT_FILE)
        processed_skus = output_excel.sheet_names
        print(f"      -> Tìm thấy {len(processed_skus)} SKU đã xử lý: {', '.join(processed_skus)}")
        
        # Load tất cả các sheet vào dictionary (bộ nhớ)
        for sheet in processed_skus:
            output_sheets_data[sheet] = pd.read_excel(OUTPUT_FILE, sheet_name=sheet)
            
    except Exception as e:
        print(f"[ERROR] Lỗi khi đọc file Output: {e}")
        return

    # -------------------------------------------------------------------------
    # BƯỚC 2: Cập nhật Sheet 'Listing' trong Internal File
    # -------------------------------------------------------------------------
    print(f"\n[2/4] Cập nhật trạng thái trong sheet 'Listing'...")
    try:
        listing_df = pd.read_excel(INTERNAL_FILE, sheet_name='Listing')
        
        # Làm sạch cột SKU (loại bỏ \n, khoảng trắng thừa)
        if 'SKU' in listing_df.columns:
            listing_df['SKU'] = listing_df['SKU'].astype(str).str.strip()
            
            # Hỗ trợ cả hai tên cột: 'Status' hoặc 'Trạng thái'
            col_status = None
            if 'Status' in listing_df.columns:
                col_status = 'Status'
            elif 'Trạng thái' in listing_df.columns:
                col_status = 'Trạng thái'
            elif 'Trạng Thái' in listing_df.columns:
                col_status = 'Trạng Thái'

            if col_status:
                # Lọc ra: Dòng có SKU nằm trong list processed_skus VÀ trạng thái hiện tại là 'Công việc Mới' hoặc rỗng
                mask = (listing_df['SKU'].isin(processed_skus)) & (listing_df[col_status].isin(['Công việc Mới', 'New', '', 'Chưa tạo']))
                num_updated = mask.sum()
                
                # Cập nhật cột trạng thái
                listing_df.loc[mask, col_status] = 'Chờ Upload'
                print(f"      -> Đã cập nhật thành 'Chờ Upload' cho {num_updated} dòng SKU trong sheet Listing.")
            else:
                print("[WARNING] Không tìm thấy cột 'Status' hoặc 'Trạng thái' trong sheet 'Listing'. Bỏ qua cập nhật Listing.")
        else:
            print("[WARNING] Không tìm thấy cột 'SKU' trong sheet 'Listing'. Bỏ qua cập nhật.")
            
    except Exception as e:
        print(f"[ERROR] Lỗi khi đọc/xử lý sheet 'Listing' từ Internal file: {e}")
        return

    # -------------------------------------------------------------------------
    # BƯỚC 3 & 4: Tạo File Mới và Ghi đè (Replace) Các Sheet
    # -------------------------------------------------------------------------
    print(f"\n[3/4] Tạo bản sao Internal file thành '{os.path.basename(UPDATED_INTERNAL_FILE)}'...")
    try:
        # Copy file nguyên bản sang một file mới để thao tác bổ sung / sửa đổi
        shutil.copy2(INTERNAL_FILE, UPDATED_INTERNAL_FILE)
    except Exception as e:
        print(f"[ERROR] Lỗi khi copy file Internal: {e}")
        return

    print(f"\n[4/4] Ghi đè (Replace) các sheet vào file mới...")
    try:
        # Sử dụng pd.ExcelWriter với engine='openpyxl', mode='a', if_sheet_exists='replace'
        with pd.ExcelWriter(UPDATED_INTERNAL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            
            # 1. Ghi lại sheet Listing (đã cập nhật) vào file mới
            listing_df.to_excel(writer, sheet_name='Listing', index=False)
            
            # 2. Cập nhật các sheet SKU trong file Internal
            # Đọc file internal để lấy dữ liệu gốc của các sheet SKU thay vì lấy từ file output 
            # (vì file output đã bị dãn dòng và mất các dòng active cũ)
            internal_excel = pd.ExcelFile(INTERNAL_FILE)
            for sheet_name in processed_skus:
                if sheet_name in internal_excel.sheet_names:
                    df_sku = pd.read_excel(INTERNAL_FILE, sheet_name=sheet_name)
                    
                    # Cập nhật hoặc tạo mới cột trạng thái trong sheet SKU
                    col_ts = None
                    for col in ["Status", "Trạng thái", "Trạng Thái"]:
                        if col in df_sku.columns:
                            col_ts = col
                            break
                    
                    if not col_ts:
                        col_ts = "Status"
                        df_sku[col_ts] = "" # Tạo cột mới nếu chưa có
                    
                    # Fill NaN bằng chuỗi rỗng để so sánh
                    df_sku[col_ts] = df_sku[col_ts].fillna("").astype(str)
                    
                    # Chuyển các dòng rỗng hoặc 'Chưa tạo' thành 'Active'
                    mask = df_sku[col_ts].str.strip().str.lower().isin(["", "chưa tạo", "nan"])
                    df_sku.loc[mask, col_ts] = "Active"
                    
                    df_sku.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"      -> Đã cập nhật trạng thái 'Active' cho sheet SKU: {sheet_name}")
                else:
                    print(f"      -> [WARNING] Không tìm thấy sheet {sheet_name} trong file Internal gốc để cập nhật.")
                
        print(f"      -> Quá trình cập nhật trạng thái hoàn tất.")
        print(f"\n[SUCCESS] HOÀN TẤT ĐỒNG BỘ!")
        print(f"          File kết quả: {UPDATED_INTERNAL_FILE}")
        print(f"          Bạn hãy kiểm tra file UPDATED này trước khi sử dụng để thay thế file Internal gốc.")
        
    except PermissionError:
        print(f"[ERROR] Không thể thay đổi file: {UPDATED_INTERNAL_FILE}\n" 
              f"        -> Vui lòng đóng file Excel nếu bạn đang mở nó ở cửa sổ khác trước khi chạy script.")
    except Exception as e:
        print(f"[ERROR] Lỗi trong quá trình xuất file Excel: {e}")

if __name__ == "__main__":
    main()
