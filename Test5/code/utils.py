import pandas as pd

def safe_div(a, b):
    return a / b if b and b > 0 else 0

def determine_phase(campaign_name, start_date_str, today, calendar):
    name = str(campaign_name).upper()
    is_season = False
    for season in calendar.get("seasons", []):
        if season.get("keyword") and season["keyword"].upper() in name:
            is_season = True
            break
    
    if not isinstance(start_date_str, str) and pd.isna(start_date_str):
        age_days = 0
    else:
        try:
            start_date = pd.to_datetime(str(start_date_str).strip(), format="%Y%m%d", errors='coerce')
            if pd.isna(start_date):
                age_days = 0
            else:
                age_days = (today - start_date).days
        except Exception:
            age_days = 0

    if is_season:
        return "Seasonal (Peak/Pre/Post)"
    else:
        if age_days < 30: return "Launch"
        elif age_days < 90: return "Growth"
        else: return "Mature/Dormant"

def get_campaign_health(acos, orders, clicks, impressions):
    if acos <= 0.25 and orders >= 3: return "⭐ Star"
    if acos > 0.40 and orders >= 1: return "🔴 Bleeder"
    if 0.25 < acos <= 0.40 and orders >= 1: return "⚠️ Watch"
    if orders == 0 and clicks >= 15: return "💀 Dead"
    if impressions < 100: return "😴 Sleep"
    return "🆕 New"

def get_keyword_suggestion(orders, acos, clicks):
    if orders >= 2 and acos <= 0.30: return "📈 Increase Bid"
    if orders == 0 and clicks >= 15: return "⏸️ Pause"
    if orders >= 1 and acos > 0.40: return "📉 Decrease Bid"
    return "👀 Monitor"

def get_search_term_suggestion(orders, acos, clicks):
    if orders >= 2 and acos <= 0.30: return "➕ Add as Exact"
    if orders == 0 and clicks >= 10: return "➖ Negative"
    if orders >= 1 and acos > 0.40: return "⚠️ Reduce/Monitor"
    return "👀 Monitor"
