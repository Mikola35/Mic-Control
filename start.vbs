Set WShell = CreateObject("WScript.Shell")
WShell.Run "cmd /c .venv\Scripts\python.exe mic_control.py", 0, False 