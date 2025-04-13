@echo off
chcp 65001 > nul
start /min .venv\Scripts\python.exe mic_control.py
if errorlevel 1 (
    echo Program failed to start. Check the error messages above.
    pause > nul
) 