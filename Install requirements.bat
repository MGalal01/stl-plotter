@echo off
title STL Plotter - install requirements
cd /d "%~dp0"
echo Installing the Python packages STL Plotter needs (one time only)...
echo.
python -m pip install -r requirements.txt
echo.
echo Done. You can now double-click "Run STL Plotter.bat".
pause
