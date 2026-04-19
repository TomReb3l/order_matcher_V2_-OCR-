@echo off
setlocal EnableExtensions

cd /d "%~dp0"

if not exist portable mkdir portable

echo [INFO] Καθαρισμός παλιών portable zip...
del /q "portable\OrderMatcher-Lite-Portable.zip" 2>nul
del /q "portable\OrderMatcher-OCR-Portable.zip" 2>nul

set "FAILED=0"

call :zip_build "OrderMatcher-Lite"
if errorlevel 1 set "FAILED=1"

call :zip_build "OrderMatcher-OCR"
if errorlevel 1 set "FAILED=1"

echo.
if "%FAILED%"=="1" (
    echo [ERROR] Κάποιο portable zip απέτυχε.
    pause
    exit /b 1
)

echo [DONE] Ολοκληρώθηκε η δημιουργία portable zip.
pause
exit /b 0

:zip_build
set "NAME=%~1"
set "SOURCE=dist\%NAME%"
set "ZIP=portable\%NAME%-Portable.zip"

if not exist "%SOURCE%" (
    echo [WARN] Δεν βρέθηκε %SOURCE%
    exit /b 0
)

echo [INFO] Δημιουργία %NAME%-Portable.zip ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -LiteralPath '%SOURCE%' -DestinationPath '%ZIP%' -Force"

if errorlevel 1 (
    echo [ERROR] Απέτυχε η δημιουργία του %ZIP%
    exit /b 1
)

exit /b 0
