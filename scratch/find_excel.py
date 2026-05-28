import glob
import os

root = r"c:\Users\nnh16\ads-trading-system\TEST"
files = glob.glob(os.path.join(root, "**", "*.xlsx"), recursive=True)
for f in files:
    print(f)
