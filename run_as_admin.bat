@echo off
:: Запуск Python-скрипта с правами администратора через PowerShell
set SCRIPT=mic_control.py
set ARGS=
:: Если надо передавать аргументы, добавь их в ARGS

powershell -Command "Start-Process '.venv\Scripts\python.exe' -ArgumentList '%SCRIPT% %ARGS%' -Verb RunAs" 