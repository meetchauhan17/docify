@echo off
echo.
echo  =============================================
echo    Docify AI - Starting Server...
echo  =============================================
echo.
echo  [*] Using venv: C:\Meet\python\venv
echo  [*] Server will start at: http://localhost:8000
echo.

rem Kill any existing process on port 8000
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 8000 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"

C:\Meet\python\venv\Scripts\python.exe -m uvicorn main:app --reload --port 8000
pause
