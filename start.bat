@echo off
chcp 65001 > nul
echo Запуск программы Mic Control...
.venv\Scripts\python.exe mic_control.py
pause 