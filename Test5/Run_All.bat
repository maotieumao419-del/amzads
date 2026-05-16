@echo off
chcp 65001 >nul
echo ==========================================
echo    AMAZON ADS DASHBOARD PIPELINE (TEST 5)
echo ==========================================
echo.

echo [1/2] Dang phan tich ten file va thiet lap thoi gian...
f:\Minhpython\venv\Scripts\python.exe f:\Minhpython\Test5\phase_bo_sung\code\parse_filename.py
if %errorlevel% neq 0 (
    echo [Loi] Giai doan phan tich file that bai!
    pause
    exit /b %errorlevel%
)
echo.

echo [2/2] Dang tao Dashboard va nap lich su vao Database...
f:\Minhpython\venv\Scripts\python.exe f:\Minhpython\Test5\phase0\code\main.py
if %errorlevel% neq 0 (
    echo [Loi] Giai doan tao Dashboard that bai!
    pause
    exit /b %errorlevel%
)
echo.

echo ==========================================
echo    HOAN THANH! Kiem tra thu muc phase0/output
echo ==========================================
pause
