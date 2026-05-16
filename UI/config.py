# config.py

# Các mốc chẩn đoán phễu PPC
TARGET_ACOS = 0.30
ALERT_CTR = 0.003
ALERT_CVR = 0.05

# Ánh xạ tọa độ cột của file Excel gộp (RAW_SKU_DATA)
# Lưu ý: Cột A là SKU (do logic gộp file tự thêm vào), 
# các chữ cái dưới đây ánh xạ đến vị trí cột trong sheet RAW.
SCHEMA_MAPPING = {
    'SKU': 'A',
    'Campaign': 'B', 
    'Impressions': 'H',
    'Clicks': 'I',
    'CTR': 'J',
    'Spend': 'K',
    'Sales': 'L',
    'Orders': 'M',
    'CVR': 'O',
    'ACOS': 'P'
}
