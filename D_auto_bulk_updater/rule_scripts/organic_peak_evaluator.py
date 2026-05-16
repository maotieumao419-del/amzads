"""
organic_peak_evaluator.py  —  Rule Script
==========================================
Rule   : Organic_Peak_Evaluator
Trigger: organic_ratio >= 0.7 AND tacos <= 0.1
Action : Update -> set_bid = current_bid * 0.85
"""

def execute(rule_config: dict, metrics: dict, match_type: str) -> list:
    """
    Parameters
    ----------
    rule_config : full rule object from rules.json
    metrics     : full_metrics dict (expected to contain 'total_sales' and 'total_orders')
    match_type  : e.g. 'Exact', 'Broad', 'Phrase'

    Returns list with one action dict if triggered, else [].
    """
    # ── Fallback an toàn ──────────────────────────────────────────────────
    # Nếu không có dữ liệu Business Report (total_sales, total_orders)
    if 'total_sales' not in metrics or 'total_orders' not in metrics:
        return []

    total_sales = float(metrics.get('total_sales', 0.0))
    total_orders = int(metrics.get('total_orders', 0))
    
    # Nếu doanh thu tổng = 0 hoặc đơn tổng = 0 thì không thể đạt Peak
    if total_sales <= 0 or total_orders <= 0:
        return []

    orders = int(metrics.get('orders', 0))
    spend = float(metrics.get('spend', 0.0))
    current_bid = float(metrics.get('current_bid', 0.0))

    # ── Logic tính toán ───────────────────────────────────────────────────
    organic_orders = total_orders - orders
    # Cẩn thận trường hợp đơn Ads cập nhật chậm hơn tổng đơn gây ra số âm
    organic_orders = max(organic_orders, 0)
    
    organic_ratio = organic_orders / total_orders
    tacos = spend / total_sales

    # ── Thresholds (Từ rules.json hoặc mặc định) ──────────────────────────
    conditions = rule_config.get('conditions', {})
    min_organic_ratio = float(conditions.get('min_organic_ratio', 0.7))
    max_tacos = float(conditions.get('max_tacos', 0.1))

    # ── Rule kích hoạt ────────────────────────────────────────────────────
    if organic_ratio >= min_organic_ratio and tacos <= max_tacos:
        target_entity = rule_config.get('target_entity', 'Keyword')
        
        # Action Payload: Giảm 15% bid
        new_bid = round(current_bid * 0.85, 2)
        
        # Xây dựng log template
        log_template = rule_config.get(
            'log_template', 
            "Organic Peak (Milking): Organic Ratio {organic_ratio_pct}%, TACoS {tacos_pct}%. Giảm 15% bid (${current_bid} -> ${new_bid})."
        )
        
        try:
            log_msg = log_template.format(
                organic_orders=organic_orders,
                total_orders=total_orders,
                organic_ratio_pct=round(organic_ratio * 100, 2),
                tacos_pct=round(tacos * 100, 2),
                current_bid=current_bid,
                new_bid=new_bid,
                match_type=match_type
            )
        except KeyError:
            log_msg = log_template

        payload = {
            'target_entity': target_entity,
            'action': 'Update',
            'set_bid': new_bid,
            'log_msg': log_msg
        }
        return [payload]

    return []
