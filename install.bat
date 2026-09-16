@echo off
echo Installing Python dependencies...
pip install -r requirements.txt
echo.
echo Installing Playwright Chromium browser...
playwright install chromium
echo.
echo Done!
echo.
echo Next steps:
echo   1. Fill in PBI_REPORT_URL in config.py (paste your report URL)
echo   2. Run: python setup_credentials.py
echo   3. Run: python setup_task_scheduler.ps1  (as Administrator, for scheduling)
echo   4. Test: python main.py
pause
