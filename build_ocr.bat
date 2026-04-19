@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

if not exist "third_party\tesseract\tesseract.exe" (
    echo [ERROR] Δεν βρέθηκε third_party\tesseract\tesseract.exe
    pause
    exit /b 1
)

if not exist "third_party\tesseract\tessdata\ell.traineddata" (
    echo [ERROR] Δεν βρέθηκε third_party\tesseract\tessdata\ell.traineddata
    pause
    exit /b 1
)

if not exist "third_party\tesseract\tessdata\eng.traineddata" (
    echo [ERROR] Δεν βρέθηκε third_party\tesseract\tessdata\eng.traineddata
    pause
    exit /b 1
)

if not exist "third_party\tesseract\tessdata\osd.traineddata" (
    echo [WARN] Δεν βρέθηκε third_party\tesseract\tessdata\osd.traineddata
)

set "PYTHON=.venv_ocr\Scripts\python.exe"

if not exist "%PYTHON%" (
    py -3 -m venv .venv_ocr
)

"%PYTHON%" -m pip install --upgrade pip setuptools wheel
"%PYTHON%" -m pip install -r requirements-ocr.txt
"%PYTHON%" -m pip install pyinstaller

> build_config.py (
    echo APP_VARIANT = "OCR"
    echo OCR_ENABLED = True
)

set "EXTRA_ARGS="
if exist "app_icon.ico" set "EXTRA_ARGS=!EXTRA_ARGS! --add-data app_icon.ico;."
if exist "assets_icon.png" set "EXTRA_ARGS=!EXTRA_ARGS! --add-data assets_icon.png;."
set "EXTRA_ARGS=!EXTRA_ARGS! --add-binary third_party\tesseract\tesseract.exe;third_party\tesseract"

for %%F in ("third_party\tesseract\*.dll") do (
    set "EXTRA_ARGS=!EXTRA_ARGS! --add-binary %%~fF;third_party\tesseract"
)

for %%F in ("third_party\tesseract\tessdata\*.traineddata") do (
    set "EXTRA_ARGS=!EXTRA_ARGS! --add-data %%~fF;third_party\tesseract\tessdata"
)

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q OrderMatcher-OCR.spec 2>nul

echo [INFO] Γίνεται build της OCR έκδοσης...
"%PYTHON%" -m PyInstaller --noconfirm --clean --windowed --name OrderMatcher-OCR !EXTRA_ARGS! ^
  --hidden-import customtkinter ^
  --hidden-import tkinter ^
  --hidden-import PIL ^
  --hidden-import pandas ^
  --hidden-import pytesseract ^
  --hidden-import pymupdf ^
  --hidden-import fitz ^
  --collect-all customtkinter ^
  --collect-all pandas ^
  --collect-all pymupdf ^
  app.py
set "BUILD_EXIT=%ERRORLEVEL%"

> build_config.py (
    echo APP_VARIANT = "LITE"
    echo OCR_ENABLED = False
)

if not "%BUILD_EXIT%"=="0" (
    echo.
    echo [ERROR] Το OCR build απέτυχε.
    endlocal
    pause
    exit /b %BUILD_EXIT%
)

echo.
echo [DONE] Ολοκληρώθηκε το OCR build.
echo Άνοιξε τον φάκελο dist\OrderMatcher-OCR\
endlocal
pause
