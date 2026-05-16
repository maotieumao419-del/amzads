import logging

def analyze_search_term(row, negative_click_threshold):
    """
    So khớp Search Term.
    Nếu Search Term == Keyword -> Trả về lệnh Pause Keyword.
    Nếu Search Term != Keyword và Search Term Clicks > threshold -> Trả về lệnh Negative Exact Search Term.
    """
    keyword_text = str(row.get('Keyword text', row.get('Keyword Text', ''))).strip().lower()
    search_term = str(row.get('Search Term', '')).strip().lower()
    
    # Lấy click của dòng hiện tại (vì file input được merge giữa STR và Bulk, 
    # mặc định row này đang chứa chỉ số của Search Term).
    st_clicks = row.get('Clicks', 0)
    try:
        st_clicks = int(float(st_clicks))
    except (ValueError, TypeError):
        st_clicks = 0

    if not search_term or search_term == "nan" or search_term == "":
        return None
        
    if not keyword_text or keyword_text == "nan":
        return None

    if search_term == keyword_text:
        logging.info(f"Search Term [{search_term}] TRÙNG khớp Keyword. Hành động: pause_keyword")
        return {
            "action": "pause_keyword",
            "target_entity": "Keyword",
            "log_msg": f"Paused: Search Term trùng hoàn toàn với Keyword '{keyword_text}'"
        }
    else:
        if st_clicks > negative_click_threshold:
            logging.info(f"Search Term [{search_term}] KHÁC Keyword và Clicks={st_clicks} > {negative_click_threshold}. Hành động: negative_exact")
            return {
                "action": "negative_exact",
                "target_entity": "Campaign Negative Keyword",
                "keyword_text": search_term,
                "match_type": "Negative Exact",
                "log_msg": f"Negative Exact: Search Term '{search_term}' cắn {st_clicks} clicks."
            }
        else:
            return {
                "action": "keep_bid",
                "target_entity": "Keyword",
                "log_msg": f"Ignored: Search Term '{search_term}' khác keyword nhưng Clicks ({st_clicks}) chưa vượt ngưỡng."
            }

def extract_to_exact(row):
    """
    Trích xuất Search Term đã ra đơn để chuẩn bị chạy Exact.
    """
    search_term = str(row.get('Search Term', '')).strip().lower()
    if not search_term or search_term == "nan" or search_term == "":
        return None
        
    return {
        "action": "extract_to_exact",
        "search_term": search_term,
        "log_msg": f"Extracted Search Term '{search_term}' to Exact Match due to Conversion."
    }
