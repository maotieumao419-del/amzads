"""
boost_low_impression_exact.py  —  Rule Script
==============================================
Rule   : Boost_Low_Impression_Exact
Trigger: impressions <= 99  AND  clicks <= 0
Routing: match_type_logic → Exact only
Action : Keyword bid increase (+0.1)
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
    conditions    = rule_config.get('conditions', {})
    impressions_max = conditions.get('impressions_max', float('inf'))
    clicks_max      = conditions.get('clicks_max',      float('inf'))

    impressions = metrics.get('impressions', 0)
    clicks      = metrics.get('clicks',      0)

    # ── Condition check ───────────────────────────────────────────────────
    if impressions > impressions_max:
        return []
    if clicks > clicks_max:
        return []

    # ── Match Type Routing ────────────────────────────────────────────────
    mt_logic = rule_config.get('match_type_logic', {})
    logic    = next(
        (v for k, v in mt_logic.items() if k.lower() == match_type.lower()),
        None
    )
    if logic is None:
        return []

    # ── Build payload ─────────────────────────────────────────────────────
    try:
        log_msg = rule_config.get('log_template', '').format(
            impressions=impressions, clicks=clicks,
            orders=metrics.get('orders', 0),
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

    return [{
        'action':        logic.get('action', 'Update'),
        'target_entity': logic.get('target_entity', 'Keyword'),
        'adjust_bid_by': logic.get('adjust_bid_by', 0.1),
        'log_msg':       log_msg,
    }]
