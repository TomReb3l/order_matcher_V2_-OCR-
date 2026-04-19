@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

set "APP_NAME=OrderMatcher-Lite"
set "PYTHON=.venv_lite\Scripts\python.exe"

if not exist "%PYTHON%" (
    py -3 -m venv .venv_lite
)

"%PYTHON%" -m pip install --upgrade pip setuptools wheel
"%PYTHON%" -m pip install -r requirements-lite.txt
"%PYTHON%" -m pip install pyinstaller customtkinter pandas

> build_config.py (
    echo APP_VARIANT = "LITE"
    echo OCR_ENABLED = False
)

set "EXTRA_ARGS="
if exist "app_icon.ico" set "EXTRA_ARGS=!EXTRA_ARGS! --add-data ""app_icon.ico;."""
if exist "assets_icon.png" set "EXTRA_ARGS=!EXTRA_ARGS! --add-data ""assets_icon.png;."""

if exist "build\%APP_NAME%" rmdir /s /q "build\%APP_NAME%"
if exist "dist\%APP_NAME%" rmdir /s /q "dist\%APP_NAME%"
del /q "%APP_NAME%.spec" 2>nul

echo [INFO] Γίνεται build της Lite έκδοσης...
"%PYTHON%" -m PyInstaller --noconfirm --clean --windowed --name "%APP_NAME%" !EXTRA_ARGS! ^
  --hidden-import customtkinter ^
  --hidden-import tkinter ^
  --hidden-import PIL ^
  --hidden-import pandas ^
  --collect-all customtkinter ^
  --collect-all pandas ^
  --exclude-module pytesseract ^
  --exclude-module pymupdf ^
  --exclude-module fitz ^
  app.py
set "BUILD_EXIT=%ERRORLEVEL%"

> build_config.py (
    echo APP_VARIANT = "LITE"
    echo OCR_ENABLED = False
)

if not "%BUILD_EXIT%"=="0" (
    echo.
    echo [ERROR] Το Lite build απέτυχε.
    endlocal
    pause
    exit /b %BUILD_EXIT%
)

echo.
echo [DONE] Ολοκληρώθηκε το Lite build.
echo Άνοιξε τον φάκελο dist\%APP_NAME%\
echo [INFO] Το build διατηρεί τυχόν υπάρχον dist\OrderMatcher-OCR\ ανέπαφο.
endlocal
pause
