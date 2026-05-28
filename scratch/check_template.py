import pandas as pd

path = r"c:\Users\nnh16\ads-trading-system\TEST\RULE&TEMPLATE\AmazonAdvertisingBulksheetSellerTemplate (3).xlsx"
xl = pd.ExcelFile(path, engine='openpyxl')
sheet_name = 'Sponsored Products Campaigns'
df = pd.read_excel(xl, sheet_name=sheet_name)
print("Columns for 'Sponsored Products Campaigns':", list(df.columns))
print("First row:", df.head(1).to_dict(orient='records'))
