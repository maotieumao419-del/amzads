@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ===================================================
echo  LUONG A: TAO CAMPAIGN MOI - TOP KEYWORD WINNER
echo ===================================================
echo.

echo [1/3] Kich hoat moi truong ao...
call ..\..\venv\Scripts\activate

echo.
echo [2/3] Dang chay Buoc 1 - Chuyen doi Excel sang CSV...
python excel_to_csv_test.py
IF %ERRORLEVEL% NEQ 0 (
    echo LOI: Buoc 1 that bai. Kiem tra file report.xlsx trong data\input\
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [3/3] Dang chay Buoc 2 - Phan tich KW va tao Campaign...
python main_test.py
IF %ERRORLEVEL% NEQ 0 (
    echo LOI: Buoc 2 that bai.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ===================================================
echo  THANH CONG! File da san sang tai: data\output\upload_manual.xlsx
echo ===================================================
call deactivate
pause
