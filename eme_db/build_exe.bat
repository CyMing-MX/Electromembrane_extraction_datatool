@echo off
REM ============================================
REM  EME - one-click rebuild script (for developer)
REM  Usage: edit code, then double-click this file.
REM  Output: dist\EME_Data\  (distribute the whole folder)
REM ============================================

cd /d "%~dp0"

set "PY=%LocalAppData%\Programs\Python\Python311\python.exe"
if not exist "%PY%" (
    echo [ERROR] Python 3.11 not found: %PY%
    pause
    exit /b 1
)

echo Building EME_Data.exe (takes a few minutes)...
"%PY%" -m PyInstaller --noconfirm --name EME_Data --onedir --console --icon "..\emedatatool.ico" ^
  --exclude-module PyQt5 --exclude-module PySide2 --exclude-module PySide6 --exclude-module tkinter --exclude-module IPython --exclude-module torch --exclude-module torchvision --exclude-module torchaudio --exclude-module scipy ^
  launcher.py ^
  --add-data "app.py;eme_db" --add-data "config.py;eme_db" --add-data "parser.py;eme_db" ^
  --add-data "calculator.py;eme_db" --add-data "models.py;eme_db" --add-data "excel_builder.py;eme_db" ^
  --add-data "dashboard.py;eme_db" --add-data "reference.py;eme_db" --add-data "naming.py;eme_db" ^
  --add-data "settings.py;eme_db" --add-data "current.py;eme_db" ^
  --collect-all streamlit --collect-all plotly --collect-all altair --collect-all pyarrow --collect-all charset_normalizer --collect-all rfc3987_syntax --collect-all jsonschema_specifications --hidden-import simpleeval
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

REM Attach user-data template (analyte/SLM/photos/sample experiments)
mkdir dist\EME_Data\data 2>nul
xcopy /e /i /y "..\data\experiments" "dist\EME_Data\data\experiments" >nul 2>&1
xcopy /e /i /y "..\data\photo" "dist\EME_Data\data\photo" >nul 2>&1
xcopy /e /i /y "..\data\current" "dist\EME_Data\data\current" >nul 2>&1
copy /y "..\data\analyte.xlsx" "dist\EME_Data\data\" >nul 2>&1
copy /y "..\data\SLM.xlsx" "dist\EME_Data\data\" >nul 2>&1

echo.
echo ============================================
echo  Build OK:  dist\EME_Data.exe  (+ dist\data folder)
echo ============================================
pause
