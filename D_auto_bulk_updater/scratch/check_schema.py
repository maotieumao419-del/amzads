import pandas as pd
import json
import glob
import os

folder = r"C:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater\data\output\update"
files = glob.glob(os.path.join(folder, "*.xlsx"))
if files:
    df = pd.read_excel(files[0], engine='openpyxl')
    print("Columns in", os.path.basename(files[0]), ":", df.columns.tolist())
    print("First few Campaign Names:", df['Campaign Name'].head().tolist())
