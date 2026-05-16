"""
traffic_starved_winner.py  —  Rule Script
==========================================
Rule   : Traffic_Starved_Winner
Trigger: impressions <= 800  AND  clicks >= 4  AND  orders >= 2
Routing: flat (all match types)
Action : Bidding Adjustment — boost Top-of-Search placement by 20%
         to pull more impressions for a keyword with excellent CVR
"""


def execute(rule_config: dict, metrics: dict, match_type: str) -> list:
    """
    Parameters
    ----------
    rule_config : full rule object from rules.json
    metrics     : full_metrics dict (11 keys)
    match_type  : e.g. 'Exact', 'Broad', 'Phrase'

    Returns list with one action dict if triggered, else [].
    """
    conditions      = rule_config.get('conditions', {})
    impressions_max = conditions.get('impressions_max', float('inf'))
    clicks_min      = conditions.get('clicks_min',      0)
    orders_min      = conditions.get('orders_min',      0)

    impressions = metrics.get('impressions', 0)
    clicks      = metrics.get('clicks',      0)
    orders      = metrics.get('orders',      0)

    # ── Condition check ───────────────────────────────────────────────────
    if impressions > impressions_max:
        return []
    if clicks < clicks_min:
        return []
    if orders < orders_min:
        return []

    # ── Build payload ─────────────────────────────────────────────────────
    try:
        log_msg = rule_config.get('log_template', '').format(
            impressions=impressions,
            clicks=clicks,
            orders=orders,
            units=metrics.get('units', 0),
            spend=metrics.get('spend', 0),
            sales=metrics.get('sales', 0),
            acos=metrics.get('acos', 0),
            cpc=metrics.get('cpc', 0),
            ctr=metrics.get('ctr', 0),
            cvr=metrics.get('cvr', 0),
            roas=metrics.get('roas', 0),
            match_type=match_type,
        )
    except KeyError:
        log_msg = rule_config.get('log_template', '')

    try:
        pct = int(float(rule_config.get('set_percentage', 20)))
    except (ValueError, TypeError):
        pct = 20

    return [{
        'action':         rule_config.get('action', 'Update'),
        'target_entity':  rule_config.get('target_entity', 'Bidding Adjustment'),
        'placement_type': rule_config.get('placement_type', 'Placement Top'),
        'set_percentage': pct,
        'log_msg':        log_msg,
    }]
