import os
import sys
import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import subprocess

# Thêm thư mục UI vào path để import các module sẵn có
UI_DIR = os.path.dirname(os.path.abspath(__file__))
if UI_DIR not in sys.path:
    sys.path.insert(0, UI_DIR)

from ui_data_prep import (
    find_latest_updated_file,
    load_ppc_updated,
    aggregate_campaigns,
    aggregate_campaign_ts,
)
from config import INPUT_SUBPATH, INPUT_PATTERN

app = FastAPI(title="Amazon Ads Trading API", version="1.0.0")

# Cấu hình CORS để Frontend (Vite/React) ở cổng khác có thể gọi được API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Trong thực tế nên giới hạn cụ thể (vd: http://localhost:5173)
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/dashboard")
def get_dashboard_data():
    """
    API lấy dữ liệu Dashboard mới nhất.
    Sử dụng lại đúng logic của file main.py để đảm bảo dữ liệu khớp 100% với file Excel.
    """
    BASE_DIR = os.path.dirname(UI_DIR)   # TEST/
    data_dir = os.path.join(BASE_DIR, "D_auto_bulk_updater", INPUT_SUBPATH)

    if not os.path.exists(data_dir):
        data_dir = os.path.join(UI_DIR, "data", "input")

    file_path = find_latest_updated_file(data_dir, INPUT_PATTERN)

    if not file_path:
        raise HTTPException(status_code=404, detail="Không tìm thấy file nguồn (PPC_*_UPDATED.xlsx)")

    try:
        # Tái sử dụng ETL
        etl = load_ppc_updated(file_path)
        df_kw = etl['df_kw']
        df_kw_ts = etl['df_kw_ts']
        date_blocks = etl['date_blocks']
        skus = etl['skus']

        if df_kw.empty:
            raise HTTPException(status_code=400, detail="File không có dữ liệu hợp lệ")

        # Tổng hợp Campaign
        df_camp = aggregate_campaigns(df_kw)

        # Tổng hợp Time Series cho Campaign
        df_camp_ts = pd.DataFrame()
        if date_blocks and not df_kw_ts.empty:
            df_camp_ts = aggregate_campaign_ts(df_kw_ts, date_blocks)

        # Chuyển đổi DataFrame sang list of dicts (JSON serializable)
        # Điền NaN thành None hoặc 0 để JSON hợp lệ
        df_kw = df_kw.fillna(0)
        df_camp = df_camp.fillna(0)
        df_camp_ts = df_camp_ts.fillna(0)

        # Chuẩn bị dữ liệu trả về
        return {
            "source_file": os.path.basename(file_path),
            "date_blocks": date_blocks,
            "skus": skus,
            "summary": {
                "total_spend": float(df_camp['Spend'].sum()),
                "total_sales": float(df_camp['Sales'].sum()),
                "total_orders": int(df_camp['Orders'].sum()),
                "total_clicks": int(df_camp['Clicks'].sum()),
                "total_impressions": int(df_camp['Impressions'].sum()),
                "overall_acos": float(df_camp['Spend'].sum() / df_camp['Sales'].sum()) if df_camp['Sales'].sum() > 0 else 0.0,
            },
            "campaigns": df_camp.to_dict(orient="records"),
            "campaign_time_series": df_camp_ts.reset_index().to_dict(orient="records") if not df_camp_ts.empty else [],
            # Tránh gửi quá nhiều keyword nếu file quá lớn, ở đây gửi toàn bộ nhưng có thể limit
            "keywords": df_kw.to_dict(orient="records")
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/run-pipeline")
def run_pipeline():
    """
    API endpoint để kích hoạt chạy toàn bộ pipeline backend.
    """
    BASE_DIR = os.path.dirname(UI_DIR)   # TEST/
    script_path = os.path.join(BASE_DIR, "D_auto_bulk_updater", "main_pipeline.py")
    
    if not os.path.exists(script_path):
        raise HTTPException(status_code=404, detail=f"Không tìm thấy script {script_path}")
        
    try:
        result = subprocess.run(
            [sys.executable, script_path], # Dùng sys.executable để lấy đúng python hiện hành
            cwd=BASE_DIR,
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            raise HTTPException(
                status_code=500, 
                detail=f"Lỗi khi chạy script: {result.stderr or result.stdout}"
            )
            
        return {"status": "success", "message": "Pipeline đã chạy thành công!"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    print("Khởi động Backend Server...")
    print("Truy cập API tại: http://localhost:8000/api/dashboard")
    uvicorn.run("serve_webapp:app", host="0.0.0.0", port=8000, reload=True)
