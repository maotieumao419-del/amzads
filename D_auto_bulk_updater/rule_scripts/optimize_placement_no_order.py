"""
optimize_placement_no_order.py  —  Rule Script (Strategy Module)
=================================================================
Rule   : Optimize_Placement_No_Order
Source : rules.json → "rule_name": "Optimize_Placement_No_Order"

Trigger condition (all must be true):
  impressions >= impressions_min  AND
  clicks      >= clicks_min       AND
  orders      <= orders_max  (= 0)

Action:
  Route by match_type via match_type_logic.
  Supported match types: Broad → set_percentage 0,
                         Phrase → set_percentage 10.
  Returns a Bidding Adjustment row for the matched type,
  or an empty list if the match_type is not listed.

Contract with run_rule_engine:
  execute() must return List[dict].  An empty list means "no action".
  None / non-list returns are silently ignored by the engine.
"""


def execute(rule_config: dict, metrics: dict, match_type: str) -> list:
    """
    Evaluate Optimize_Placement_No_Order against live metrics.

    Parameters
    ----------
    rule_config : dict
        The full rule object from rules.json for this rule.
    metrics : dict
        The full_metrics payload produced by extract_full_metrics().
        Keys: impressions, clicks, orders, units, spend, sales,
              ctr, cvr, acos, cpc, roas.
    match_type : str
        Match type string parsed from the 'Ghi chú' column, e.g. 'Broad'.

    Returns
    -------
    list
        One-element list with the action dict if triggered, else [].
    """
    # ── Step 1: Read condition thresholds from config ─────────────────────
    conditions    = rule_config.get('conditions', {})
    impressions_min = conditions.get('impressions_min', 0)
    clicks_min      = conditions.get('clicks_min',      0)
    orders_max      = conditions.get('orders_max',      0)

    # ── Step 2: Volume trigger check ──────────────────────────────────────
    impressions = metrics.get('impressions', 0)
    clicks      = metrics.get('clicks',      0)
    orders      = metrics.get('orders',      0)

    if impressions < impressions_min:
        return []
    if clicks < clicks_min:
        return []
    if orders > orders_max:
        return []

    # ── Step 3: Match Type Routing (case-insensitive) ─────────────────────
    mt_logic = rule_config.get('match_type_logic', {})
    logic    = next(
        (v for k, v in mt_logic.items() if k.lower() == match_type.lower()),
        None
    )
    if logic is None:
        # This match type is not in match_type_logic → no action
        return []

    # ── Step 4: Build output payload ──────────────────────────────────────
    try:
        log_msg = rule_config.get('log_template', '').format(
            impressions=impressions,
            clicks=clicks,
            orders=orders,
            match_type=match_type,
            units=metrics.get('units', 0),
            spend=metrics.get('spend',  0),
            sales=metrics.get('sales',  0),
            acos=metrics.get('acos',   0),
            cpc=metrics.get('cpc',    0),
            ctr=metrics.get('ctr',    0),
            cvr=metrics.get('cvr',    0),
            roas=metrics.get('roas',   0),
        )
    except KeyError:
        log_msg = rule_config.get('log_template', '')

    return [{
        'action':        logic.get('action', 'Update'),
        'target_entity': 'Bidding Adjustment',
        'placement_type': logic.get('placement_type', ''),
        'set_percentage': logic.get('set_percentage', 0),
        'log_msg':        log_msg,
    }]
