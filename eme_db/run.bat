@echo off
REM ============================================
REM  EME Experiment Data Management System
REM  Launch script - double click to start
REM  Re-running while service is up RESTARTS the
REM  service so code updates always take effect,
REM  then opens the browser.
REM ============================================

cd /d "%~dp0"

set "PY=%LocalAppData%\Programs\Python\Python311\python.exe"
if not exist "%PY%" (
    echo [ERROR] Python 3.11 not found: %PY%
    pause
    exit /b 1
)

set "PORT=8501"

REM If a service is already running on the port, stop it first.
REM (Otherwise code updates would never be loaded - the old process
REM  keeps serving the old version, and re-running this script would
REM  just open a browser to the stale app.)
"%PY%" -c "import socket,sys; s=socket.socket(); s.settimeout(0.5); sys.exit(0 if s.connect_ex(('127.0.0.1',%PORT%))==0 else 1)" >nul 2>&1
if not errorlevel 1 (
    echo Service already running - restarting to load latest code...
    for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":%PORT% .*LISTENING"') do taskkill /F /PID %%P >nul 2>&1
    timeout /t 2 /nobreak >nul
)

echo ============================================
echo  EME Experiment Data Management System
echo  Starting server (first load takes a while)...
echo ============================================
echo.

start "" /b "%PY%" -m streamlit run app.py --server.headless=true --server.port=%PORT% --browser.gatherUsageStats=false

REM Wait until port is ready (max 60s), then open browser once
set /a TRIES=0
:WAITLOOP
timeout /t 1 /nobreak >nul
"%PY%" -c "import socket,sys; s=socket.socket(); s.settimeout(1); sys.exit(0 if s.connect_ex(('127.0.0.1',%PORT%))==0 else 1)" >nul 2>&1
if errorlevel 1 (
    set /a TRIES+=1
    <nul set /p =.
    if %TRIES% LSS 60 goto WAITLOOP
    echo.
    echo [WARN] Server not ready in 60s. Open http://localhost:%PORT% manually later.
    goto END
)
echo.
echo Server ready - opening browser...
start "" http://localhost:%PORT%
echo.
echo ============================================
echo  Service is running at http://localhost:%PORT%
echo  Keep this window open. Press any key to STOP
echo  the service (or just close this window).
echo ============================================
pause >nul
REM Stop the service when user presses a key
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":%PORT% .*LISTENING"') do taskkill /F /PID %%P >nul 2>&1

:END
