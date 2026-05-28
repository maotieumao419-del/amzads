# config.py

# ─── KPI Thresholds ────────────────────────────────────────────────────────
TARGET_ACOS  = 0.30   # Ngưỡng ACOS mục tiêu (30%)
ALERT_CTR    = 0.003  # CTR dưới ngưỡng này = cảnh báo đỏ
ALERT_CVR    = 0.05   # CVR dưới ngưỡng này = cảnh báo vàng

# ─── Health Tag Thresholds ──────────────────────────────────────────────────
HEALTH_ACOS_STAR     = 0.25   # ACOS ≤ 25% → ⭐ Star
HEALTH_ACOS_BLEEDER  = 0.40   # ACOS > 40% → 🔴 Bleeder
HEALTH_ACOS_WATCH    = 0.40   # 25% < ACOS ≤ 40% → ⚠️ Watch
HEALTH_MIN_ORDERS    = 1      # Tối thiểu orders để tính ACOS
HEALTH_DEAD_CLICKS   = 15     # Clicks >= X mà 0 order → 💀 Dead
HEALTH_SLEEP_IMPR    = 100    # Impressions < X → 😴 Sleep

# ─── Dashboard Config ───────────────────────────────────────────────────────
# Số kỳ ngày gần nhất lấy để vẽ sparkline (mỗi file UPDATED có N date blocks)
SPARKLINE_PERIODS = 6

# ─── Input / Output Paths ───────────────────────────────────────────────────
# Relative từ D_auto_bulk_updater/
INPUT_SUBPATH  = "data/final_xlsx"
INPUT_PATTERN  = "PPC_*_UPDATED*.xlsx"
OUTPUT_SUBPATH = "data/output"
OUTPUT_FILE    = "AMZ_Interactive_Dashboard.xlsx"

# ─── Sheet Tab Names ────────────────────────────────────────────────────────
SHEET_OVERVIEW   = "📊 OVERVIEW"
SHEET_CAMPAIGNS  = "🎯 CAMPAIGNS"
SHEET_KEYWORDS   = "🔑 KEYWORDS"
SHEET_DEEP_DIVE  = "🔍 DEEP DIVE"
SHEET_RAW_KW     = "RAW_KW_DATA"
SHEET_TS_DATA    = "TS_DATA"

# ─── Base Columns from PPC_UPDATED.xlsx ─────────────────────────────────────
BASE_COLS = ['STT', 'Campaign Name', 'Loại Campaign', 'Target', 'Ghi chú', 'Status']
METRIC_COLS = [
    'Impressions', 'Clicks', 'Click-through Rate', 'Spend', 'Sales',
    'Orders', 'Units', 'Conversion Rate', 'ACOS', 'CPC', 'ROAS', 'Placement Note'
]
