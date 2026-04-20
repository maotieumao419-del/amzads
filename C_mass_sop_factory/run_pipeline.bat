@echo off
title C_mass_sop_factory Pipeline
chcp 65001 >nul

:: Chuyển đến thư mục chứa file bat
cd /d "%~dp0"

echo ======================================================================
echo BẮT ĐẦU CHẠY PIPELINE (FOLDER C)
echo ======================================================================
echo.

echo [0/4] Kiem tra va cai dat thu vien can thiet...
pip install pandas openpyxl >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [CẢNH BÁO] Khong the cai dat thu vien tu dong. Vui long kiem tra pip.
)

echo.
echo [1/4] Đang chay excel_new_camp.py (Xu ly file nguyen va cross-join)...
python excel_new_camp.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LỖI] excel_new_camp.py gap van de. Dung pipeline!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/4] Đang chay mass_sop_factory.py (Tao file Bulk cho Amazon)...
python mass_sop_factory.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LỖI] mass_sop_factory.py gap van de. Dung pipeline!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/4] Đang chay amz_schema_validator.py (Kiem tra form Bulk Amazon)...
python amz_schema_validator.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [CẢNH BÁO] Qua trinh cham diem Schema gap van de!
)

echo.
echo [4/4] Đang chay sync_master_file.py (Dong bo trang thai ve Master file)...
python sync_master_file.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LỖI] sync_master_file.py gap van de. Dung pipeline!
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ======================================================================
echo PIPELINE HOÀN TẤT!
echo 1. Kiem tra thu muc data/output de nhan cac file Bulk_*.xlsx
echo 2. Kiem tra file PPC_NGUYEN_UPDATED.xlsx de xem ket qua dong bo
echo ======================================================================
pause
