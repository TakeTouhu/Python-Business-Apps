@echo off
rem Launch business_utility.py and keep the window open so errors are visible
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (
    py business_utility.py
) else (
    python business_utility.py
)
echo.
echo ---- Exit code: %errorlevel% ----
pause
