@echo off
title FNCL Group Manulife Investment Report Generator Runner
echo ==========================================================
echo       FNCL Group Manulife Investment Report Generator
echo ==========================================================
echo.
echo [1/2] Opening dashboard in your default browser...
start http://127.0.0.1:8080
echo.
echo [2/2] Launching backend server...
echo Local Access Link: http://127.0.0.1:8080
echo.
echo If sharing with another laptop on the same Wi-Fi, use one of these:
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /i "ipv4"') do (
    echo   http:%%A:8080
)
echo.
echo (Keep this window open. Press Ctrl+C to stop the server)
echo.
cd backend
venv\Scripts\python.exe -m uvicorn app:app --host 0.0.0.0 --port 8080
pause

