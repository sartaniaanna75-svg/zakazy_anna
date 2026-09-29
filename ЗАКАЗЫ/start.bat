@echo off
cd /d "%~dp0"
"C:\Users\1\AppData\Local\Programs\Python\Python312\python.exe" "%~dp0src\build_html.py"
if errorlevel 1 (
  echo Не удалось собрать страницу.
  pause
  exit /b 1
)
start "" "http://127.0.0.1:8765/zakazy.html"
"C:\Users\1\AppData\Local\Programs\Python\Python312\python.exe" -m http.server 8765 --bind 127.0.0.1
pause
