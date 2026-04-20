import os
import shutil
import pandas as pd

# ==============================================================================
# CẤU HÌNH ĐƯỜNG DẪN TỆP TIN
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'data', 'input')

MASTER_FILE = os.path.join(INPUT_DIR, 'PPC_NGUYEN.xlsx')
OUTPUT_FILE = os.path.join(INPUT_DIR, 'PPC_PROCESSED_OUTPUT.xlsx')
UPDATED_MASTER_FILE = os.path.join(INPUT_DIR, 'PPC_NGUYEN_UPDATED.xlsx')

def main():
    print(f"\n[INFO] === BẮT ĐẦU ĐỒNG BỘ DỮ LIỆU VÀO MASTER FILE ===")
    
    # Kiểm tra sự tồn tại của các file cần thiết
    if not os.path.exists(MASTER_FILE):
        print(f"[ERROR] Không tìm thấy file Master tại:\n'{MASTER_FILE}'.")
        return
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
    # BƯỚC 2: Cập nhật Sheet 'Listing' trong Master File
    # -------------------------------------------------------------------------
    print(f"\n[2/4] Cập nhật trạng thái trong sheet 'Listing'...")
    try:
        listing_df = pd.read_excel(MASTER_FILE, sheet_name='Listing')
        
        # Làm sạch cột SKU (loại bỏ \n, khoảng trắng thừa)
        if 'SKU' in listing_df.columns:
            listing_df['SKU'] = listing_df['SKU'].astype(str).str.strip()
            
            # Cập nhật trạng thái
            if 'Trạng thái' in listing_df.columns:
                # Lọc ra: Dòng có SKU nằm trong list processed_skus VÀ trạng thái hiện tại là 'Công việc Mới'
                mask = (listing_df['SKU'].isin(processed_skus)) & (listing_df['Trạng thái'] == 'Công việc Mới')
                num_updated = mask.sum()
                
                # Cập nhật cột trạng thái
                listing_df.loc[mask, 'Trạng thái'] = 'Chờ Upload'
                print(f"      -> Đã cập nhật thành 'Chờ Upload' cho {num_updated} dòng SKU.")
            else:
                print("[WARNING] Không tìm thấy cột 'Trạng thái' trong sheet 'Listing'. Bỏ qua cập nhật.")
        else:
            print("[WARNING] Không tìm thấy cột 'SKU' trong sheet 'Listing'. Bỏ qua cập nhật.")
            
    except Exception as e:
        print(f"[ERROR] Lỗi khi đọc/xử lý sheet 'Listing' từ Master file: {e}")
        return

    # -------------------------------------------------------------------------
    # BƯỚC 3 & 4: Tạo File Mới và Ghi đè (Replace) Các Sheet
    # -------------------------------------------------------------------------
    print(f"\n[3/4] Tạo bản sao Master file thành '{os.path.basename(UPDATED_MASTER_FILE)}'...")
    try:
        # Copy file nguyên bản sang một file mới để thao tác bổ sung / sửa đổi
        shutil.copy2(MASTER_FILE, UPDATED_MASTER_FILE)
    except Exception as e:
        print(f"[ERROR] Lỗi khi copy file Master: {e}")
        return

    print(f"\n[4/4] Ghi đè (Replace) các sheet vào file mới...")
    try:
        # Sử dụng pd.ExcelWriter với engine='openpyxl', mode='a', if_sheet_exists='replace'
        # Thiết lập này cho phép chúng ta thay thế/thêm sheet vào một file Excel ĐÃ CÓ (UPDATED_MASTER_FILE)
        # mà không làm ảnh hưởng (xoá) các sheet khác không bị gọi đến.
        with pd.ExcelWriter(UPDATED_MASTER_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            
            # 1. Ghi lại sheet Listing (đã cập nhật) vào file mới
            listing_df.to_excel(writer, sheet_name='Listing', index=False)
            
            # 2. Ghi đè các sheet SKU
            for sheet_name, df in output_sheets_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        print(f"      -> Quá trình ghi đè hoàn tất.")
        print(f"\n[SUCCESS] HOÀN TẤT ĐỒNG BỘ!")
        print(f"          File kết quả: {UPDATED_MASTER_FILE}")
        print(f"          Bạn hãy kiểm tra file UPDATED này trước khi sử dụng để thay thế file Master gốc.")
        
    except PermissionError:
        print(f"[ERROR] Không thể thay đổi file: {UPDATED_MASTER_FILE}\n" 
              f"        -> Vui lòng đóng file Excel nếu bạn đang mở nó ở cửa sổ khác trước khi chạy script.")
    except Exception as e:
        print(f"[ERROR] Lỗi trong quá trình xuất file Excel: {e}")

if __name__ == "__main__":
    main()
