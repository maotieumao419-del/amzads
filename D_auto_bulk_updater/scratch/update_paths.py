import os
import glob
import re

base_path = r'c:\Users\nnh16\ads-trading-system\TEST\D_auto_bulk_updater'
files_to_update = glob.glob(os.path.join(base_path, '*.py'))

replacements = {
    r'\"data\",\s*\"input 1\"': r'"data", "raw_xlsx"',
    r'\'data\',\s*\'input 1\'': r"'data', 'raw_xlsx'",
    r'\"data\",\s*\"input 2\"': r'"data", "working_json"',
    r'\'data\',\s*\'input 2\'': r"'data', 'working_json'",
    r'\"data\",\s*\"input 3\"': r'"data", "working_json"',
    r'\'data\',\s*\'input 3\'': r"'data', 'working_json'",
    r'\"data\",\s*\"input 4\"': r'"data", "final_xlsx"',
    r'\'data\',\s*\'input 4\'': r"'data', 'final_xlsx'",
    r'\"data\",\s*\"output\"': r'"data", "final_xlsx"',
    r'\'data\',\s*\'output\'': r"'data', 'final_xlsx'",
}

for fp in files_to_update:
    if os.path.basename(fp) == 'main_pipeline.py':
        continue
    with open(fp, 'r', encoding='utf-8') as f:
        content = f.read()
    
    new_content = content
    for pattern, repl in replacements.items():
        new_content = re.sub(pattern, repl, new_content)
    
    if new_content != content:
        with open(fp, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f'Updated {os.path.basename(fp)}')
