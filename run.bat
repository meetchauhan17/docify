@echo off
echo.
echo  =============================================
echo    Docify AI - Starting Server...
echo  =============================================
echo.
echo  [*] Using venv: C:\Meet\python\venv
echo  [*] Server will start at: http://localhost:8000
echo.
C:\Meet\python\venv\Scripts\python.exe -m uvicorn main:app --reload --port 8000
pause
