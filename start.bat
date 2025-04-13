@echo off
chcp 65001 > nul
title Mic Control
echo Starting Mic Control...
.venv\Scripts\python.exe mic_control.py
if errorlevel 1 (
    echo.
    echo Program failed to start. Check the error messages above.
    echo Press any key to exit...
    pause > nul
) 