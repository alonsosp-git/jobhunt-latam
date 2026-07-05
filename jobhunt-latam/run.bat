@echo off
REM JobHunt LatAm - Windows launcher (double-click me)
cd /d "%~dp0"
echo Installing/updating dependencies (first run only)...
python -m pip install -r requirements.txt
echo.
echo Starting JobHunt LatAm... open http://127.0.0.1:5000 in your browser.
python app.py
pause
