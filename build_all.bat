@echo off
setlocal EnableExtensions EnableDelayedExpansion
title OrderMatcher - Build All

cd /d "%~dp0"

echo.
echo [1/7] Checking Inno Setup...
set "ISCC="
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"

if not defined ISCC (
    echo ERROR: Inno Setup 6 not found.
    echo Install Inno Setup 6 or edit the ISCC path inside this script.
    exit /b 1
)

echo Found ISCC: "%ISCC%"

echo.
echo [2/7] Cleaning old portable and installer outputs...
if exist "portable" rmdir /s /q "portable"
if exist "installer_output" rmdir /s /q "installer_output"
mkdir "portable" >nul 2>&1
mkdir "installer_output" >nul 2>&1

echo.
echo [3/7] Building Lite...
call "%~dp0build_lite.bat"
if errorlevel 1 (
    echo ERROR: Lite build failed.
    exit /b 1
)
if not exist "dist\OrderMatcher-Lite" (
    echo ERROR: dist\OrderMatcher-Lite was not created.
    exit /b 1
)

echo.
echo [4/7] Building OCR...
call "%~dp0build_ocr.bat"
if errorlevel 1 (
    echo ERROR: OCR build failed.
    exit /b 1
)
if not exist "dist\OrderMatcher-OCR" (
    echo ERROR: dist\OrderMatcher-OCR was not created.
    exit /b 1
)

echo.
echo [5/7] Creating portable ZIP packages...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "if (Test-Path 'portable\OrderMatcher-Lite-Portable.zip') { Remove-Item 'portable\OrderMatcher-Lite-Portable.zip' -Force }; " ^
    "if (Test-Path 'portable\OrderMatcher-OCR-Portable.zip') { Remove-Item 'portable\OrderMatcher-OCR-Portable.zip' -Force }; " ^
    "Compress-Archive -Path 'dist\OrderMatcher-Lite' -DestinationPath 'portable\OrderMatcher-Lite-Portable.zip' -Force; " ^
    "Compress-Archive -Path 'dist\OrderMatcher-OCR' -DestinationPath 'portable\OrderMatcher-OCR-Portable.zip' -Force"
if errorlevel 1 (
    echo ERROR: Portable ZIP creation failed.
    exit /b 1
)

echo.
echo [6/7] Compiling installer for Lite...
"%ISCC%" "%~dp0installer_lite.iss"
if errorlevel 1 (
    echo ERROR: Lite installer compile failed.
    exit /b 1
)

echo.
echo [7/7] Compiling installer for OCR...
"%ISCC%" "%~dp0installer_ocr.iss"
if errorlevel 1 (
    echo ERROR: OCR installer compile failed.
    exit /b 1
)

echo.
echo DONE.
echo.
echo Outputs:
echo   Portable:
echo     portable\OrderMatcher-Lite-Portable.zip
echo     portable\OrderMatcher-OCR-Portable.zip
echo.
echo   Installers:
echo     installer_output\OrderMatcher-Lite-Setup.exe
echo     installer_output\OrderMatcher-OCR-Setup.exe
echo.
pause
exit /b 0
