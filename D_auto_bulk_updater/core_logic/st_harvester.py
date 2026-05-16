"""
st_harvester.py  —  Search Term Harvester Engine
================================================
Role  : Reads Search Term Reports (ST_Report_*.json), evaluates search terms
        based on performance, and generates Amazon Upload bulksheets to isolate
        or scale the search terms.

Input  : data/working_json/ST_Report_*.json
Output : data/final_xlsx/Amazon_Upload_Ready_ST_Harvest_[Date].xlsx
"""

import os
import sys
import json
import glob
import logging
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE_DIR, 'io_handlers'))
from excel_json_handler import json_to_xlsx, AMAZON_BULK_COLUMNS

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

JSON_DIR  = os.path.join(BASE_DIR, 'data', 'working_json')
FINAL_DIR = os.path.join(BASE_DIR, 'data', 'final_xlsx')
os.makedirs(FINAL_DIR, exist_ok=True)

_GARBAGE_RE = re.compile(r'[^\d.]')

def clean_metric(value, type_func=float):
    """Clean metric values by stripping non-numeric characters."""
    try:
        if value is None or str(value).strip() in ('', '-', '--', 'N/A', 'n/a'):
            return type_func(0)
        cleaned = _GARBAGE_RE.sub('', str(value))
        if not cleaned:
            return type_func(0)
        return type_func(float(cleaned))
    except (ValueError, TypeError):
        return type_func(0)

def extract_st_metrics(row: dict) -> dict:
    """
    TASK 1: Extract and clean Search Term metrics & identifiers.
    Returns safe typed metrics.
    """
    return {
        'campaign_name': str(row.get('Campaign Name', '')).strip(),
        'ad_group_name': str(row.get('Ad Group Name', '')).strip(),
        'customer_search_term': str(row.get('Customer Search Term', '')).strip().lower(),
        'match_type': str(row.get('Match Type', '')).strip().lower(),
        'impressions': clean_metric(row.get('Impressions', 0), int),
        'clicks': clean_metric(row.get('Clicks', 0), int),
        'spend': clean_metric(row.get('Spend', 0), float),
        'sales': clean_metric(row.get('Sales', 0), float),
        'orders': clean_metric(row.get('Orders', 0), int),
        'cpc': clean_metric(row.get('CPC', 0), float),
    }

def evaluate_search_term(row_data: dict, metrics: dict) -> list:
    """
    TASK 2: Evaluate Search Term into Bleeder or Winner branches.
    Returns a list of dicts representing bulk actions.
    """
    actions = []
    
    orders = metrics['orders']
    clicks = metrics['clicks']
    match_type = metrics['match_type']
    cpc = metrics['cpc']
    
    st = metrics['customer_search_term']
    camp = metrics['campaign_name']
    ag = metrics['ad_group_name']
    
    if not st or not camp:
        return actions

    # Branch 1: BLEEDER TERM (Negative Exact)
    is_phrase_bleeder = (match_type == 'phrase' and clicks > 10)
    # Handle both broad and auto equivalent match types usually seen in reports
    is_broad_bleeder  = ((match_type == 'broad' or match_type == '- (targeting)') and clicks > 15)
    
    if orders == 0 and (is_phrase_bleeder or is_broad_bleeder):
        actions.append({
            'type': 'bleeder',
            'row': {
                'Product': 'Sponsored Products',
                'Entity': 'Negative Keyword',
                'Operation': 'Create',
                'Campaign Name': camp,
                'Ad Group Name': ag,
                'Keyword Text': st,
                'Match Type': 'Negative Exact',
                'State': 'Enabled'
            }
        })
        
    # Branch 2: WINNER TERM (Harvesting & Isolation)
    elif orders >= 1:
        # Row 1: Negate in Source
        actions.append({
            'type': 'winner_negate',
            'row': {
                'Product': 'Sponsored Products',
                'Entity': 'Negative Keyword',
                'Operation': 'Create',
                'Campaign Name': camp,
                'Ad Group Name': ag,
                'Keyword Text': st,
                'Match Type': 'Negative Exact',
                'State': 'Enabled'
            }
        })
        
        # Row 2: Scale (Create Exact)
        scaled_camp = f"{camp}_Exact"
        scaled_ag   = f"{ag}_Exact" if ag else ""
        # Base Bid = CPC hiện tại * 1.2
        bid = round(cpc * 1.2, 2) if cpc > 0 else 0.50 
        
        actions.append({
            'type': 'winner_scale',
            'row': {
                'Product': 'Sponsored Products',
                'Entity': 'Keyword',
                'Operation': 'Create',
                'Campaign Name': scaled_camp,
                'Ad Group Name': scaled_ag,
                'Keyword Text': st,
                'Match Type': 'Exact',
                'Bid': str(bid),
                'State': 'Enabled'
            }
        })

    return actions

def main():
    print("--- SEARCH TERM HARVESTER ---")
    st_files = glob.glob(os.path.join(JSON_DIR, 'ST_Report_*.json'))
    
    if not st_files:
        logging.warning("No Search Term reports found (ST_Report_*.json) in data/working_json/.")
        return
        
    all_upload_rows = []
    bleeder_count = 0
    winner_count = 0
    
    for file_path in st_files:
        logging.info(f"Processing: {os.path.basename(file_path)}")
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            # Normalize data into a flat list of rows
            rows = []
            if isinstance(data, dict):
                for sheet, sheet_data in data.items():
                    if isinstance(sheet_data, list):
                        rows.extend(sheet_data)
                    elif isinstance(sheet_data, dict) and 'rows' in sheet_data:
                        rows.extend(sheet_data['rows'])
            elif isinstance(data, list):
                rows = data
                
            for row in rows:
                metrics = extract_st_metrics(row)
                acts = evaluate_search_term(row, metrics)
                
                for act in acts:
                    if act['type'] == 'bleeder':
                        bleeder_count += 1
                    elif act['type'] == 'winner_scale':
                        winner_count += 1
                    
                    # Fill standard columns
                    upload_row = {col: '' for col in AMAZON_BULK_COLUMNS}
                    upload_row.update(act['row'])
                    all_upload_rows.append(upload_row)
                    
        except Exception as e:
            logging.error(f"Error processing {file_path}: {e}")
            
    logging.info(f"Harvesting complete. Found {bleeder_count} Bleeder terms and {winner_count} Winner terms.")
    
    # TASK 3: OUTPUT GENERATION
    if all_upload_rows:
        today = datetime.now().strftime('%d%m%Y')
        out_name = f"Amazon_Upload_Ready_ST_Harvest_{today}"
        json_out = os.path.join(JSON_DIR, f"{out_name}.json")
        xlsx_out = os.path.join(FINAL_DIR, f"{out_name}.xlsx")
        
        with open(json_out, 'w', encoding='utf-8') as f:
            json.dump(all_upload_rows, f, ensure_ascii=False, indent=2)
            
        if json_to_xlsx(json_out, xlsx_out):
            logging.info(f"Successfully generated Excel: {os.path.basename(xlsx_out)}")
    else:
        logging.info("No actionable search terms found. No output generated.")

if __name__ == '__main__':
    main()
