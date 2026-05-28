import pandas as pd
path = r"c:\Users\nnh16\ads-trading-system\TEST\RULE&TEMPLATE\AmazonAdvertisingBulksheetSellerTemplate (3).xlsx"
df = pd.read_excel(path, sheet_name='Sponsored Products Campaigns')
print(list(df.columns))
