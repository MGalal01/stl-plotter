@echo off
title STL Plotter
cd /d "%~dp0"
set STL_PLOTTER_SAVE_SETTINGS=1
echo Starting STL Plotter... your browser will open in a few seconds.
echo Keep this window open while you use the app. Close it to stop the app.
echo.
start "" /b cmd /c "timeout /t 4 /nobreak >nul & start http://localhost:8765"
python -m streamlit run app.py --server.port 8765 --server.address localhost --server.headless true
echo.
echo The app has stopped. If you saw an error above, run "Install requirements.bat" first.
pause
