import pandas as pd
import sys

def compare_excels(file1, file2):
    print(f"File 1: {file1}")
    print(f"File 2: {file2}")
    
    try:
        xl1 = pd.ExcelFile(file1)
        xl2 = pd.ExcelFile(file2)
    except Exception as e:
        print(f"Error loading files: {e}")
        return

    print(f"\nSheet names in File 1: {xl1.sheet_names}")
    print(f"Sheet names in File 2: {xl2.sheet_names}")
    
    # Compare sheets
    common_sheets = set(xl1.sheet_names).intersection(set(xl2.sheet_names))
    
    for sheet in common_sheets:
        print(f"\n--- Comparing Sheet: '{sheet}' ---")
        df1 = xl1.parse(sheet)
        df2 = xl2.parse(sheet)
        
        print(f"File 1 shape: {df1.shape}")
        print(f"File 2 shape: {df2.shape}")
        
        if list(df1.columns) != list(df2.columns):
            print("Columns differ!")
            print(f"File 1 columns: {list(df1.columns)}")
            print(f"File 2 columns: {list(df2.columns)}")
            
            # Find differences in columns
            s1 = set(df1.columns)
            s2 = set(df2.columns)
            print(f"Columns only in File 1: {s1 - s2}")
            print(f"Columns only in File 2: {s2 - s1}")
        else:
            print("Columns are identical.")
            # Compare data if columns are same
            if df1.equals(df2):
                print("Data is identical.")
            else:
                print("Data differs.")

if __name__ == "__main__":
    f1 = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\C_mass_sop_factory\data\output\FATHERDAY_DADGIFTFROMDAUGHTER.xlsx"
    f2 = r"c:\Users\nnh16\ads-trading-system\TEST\G_reset_system\C_mass_sop_factory\data\output\Bulk_Create_FATHERDAY_DADGIFTFROMDAUGHTER.xlsx"
    compare_excels(f1, f2)
