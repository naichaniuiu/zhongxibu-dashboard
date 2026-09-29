@echo off
chcp 936 > nul

REM ============================================================
REM  Zhongxibu Dashboard Q2 - One-Click Update + Auto Push to GitHub
REM  Double-click to: Read Excel -> Generate HTML -> Push to Pages
REM  URL: https://naichaniuiu.github.io/zhongxibu-dashboard/
REM ============================================================

setlocal

REM ==================== CONFIG ====================
REM Data file (modify if Excel is in another location)
set "EXCEL=%USERPROFILE%\Desktop\新建文件夹\业绩 欠款看板 Q2.xlsx"

REM Python interpreter (WorkBuddy managed)
set "PYTHON=C:\Users\%USERNAME%\.workbuddy\binaries\python\versions\3.13.12\python.exe"
REM =================================================

REM Detect paths relative to this .bat
set "SELF_DIR=%~dp0"
set "PROJECT_DIR=%SELF_DIR%.."
set "LOG=%SELF_DIR%update_log.txt"

echo === Dashboard Q2 Update: %date% %time% === > "%LOG%"

echo ============================================================
echo   Zhongxibu Q2 Dashboard One-Click Update
echo   Time: %date% %time%
echo ============================================================
echo.

REM ---- Step 1: Check Excel ----
if not exist "%EXCEL%" (
    echo [ERROR] Excel not found: "%EXCEL%" >> "%LOG%"
    echo [1/3] FAILED: Excel file not found!
    echo   Expected: %EXCEL%
    echo   Please put it at the path above or modify CONFIG in this .bat.
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)
echo [1/3] Excel found: %EXCEL%
echo [1/4] Excel found. >> "%LOG%"

REM ---- Step 2: Build data ----
echo.
echo [2/3] Building data from Excel...
echo [2/4] Building data... >> "%LOG%"
"%PYTHON%" "%PROJECT_DIR%\build_data.py" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] build_data.py failed. >> "%LOG%"
    echo [2/3] FAILED: build_data.py error. See update_log.txt
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)
echo [2/3] Done.

REM ---- Step 3: Inject into HTML + push to GitHub ----
echo.
echo [3/3] Generating HTML and pushing to GitHub...
echo [3/4] Inject + push... >> "%LOG%"
"%PYTHON%" "%PROJECT_DIR%\inject.py" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] inject.py failed. >> "%LOG%"
    echo [3/3] FAILED: inject.py error. See update_log.txt
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)

cd /d "%SELF_DIR%"
git add index.html >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] git add failed. >> "%LOG%"
    echo [3/3] FAILED: git add.
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)

REM Check if there are staged changes
git diff --cached --quiet
if not errorlevel 1 (
    echo [INFO] No changes to commit. Skipping push. >> "%LOG%"
    echo [3/3] No changes to commit. Data may be same as last update.
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 0
)

git -c user.email=deploy@local -c user.name=Deployer commit -m "Update Q2 dashboard %date%" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] git commit failed. >> "%LOG%"
    echo [3/3] FAILED: git commit.
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)

git push origin main >> "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] git push failed. >> "%LOG%"
    echo [3/3] FAILED: git push.
    echo   Possible causes: network, expired token, or repo permissions.
    echo   See update_log.txt for details.
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)

echo === Update completed: %date% %time% === >> "%LOG%"
echo [3/3] Pushed successfully!
echo.
echo ============================================================
echo   DONE! Dashboard updated.
echo.
echo   Managers can view at:
echo   https://naichaniuiu.github.io/zhongxibu-dashboard/
echo.
echo   GitHub Pages may take 1-2 minutes to refresh.
echo ============================================================
echo.
start "" "https://naichaniuiu.github.io/zhongxibu-dashboard/"
echo Press any key to close...
pause >nul
exit /b 0
