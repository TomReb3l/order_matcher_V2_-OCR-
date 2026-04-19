@echo off
setlocal EnableExtensions

cd /d "%~dp0"

if not exist portable mkdir portable

echo [INFO] Καθαρισμός παλιών portable zip...
del /q portable\OrderMatcher-Lite-Portable.zip 2>nul
del /q portable\OrderMatcher-OCR-Portable.zip 2>nul

if exist dist\OrderMatcher-Lite (
    echo [INFO] Δημιουργία OrderMatcher-Lite-Portable.zip ...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\OrderMatcher-Lite\*' -DestinationPath 'portable\OrderMatcher-Lite-Portable.zip' -Force"
) else (
    echo [WARN] Δεν βρέθηκε dist\OrderMatcher-Lite
)

if exist dist\OrderMatcher-OCR (
    echo [INFO] Δημιουργία OrderMatcher-OCR-Portable.zip ...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'dist\OrderMatcher-OCR\*' -DestinationPath 'portable\OrderMatcher-OCR-Portable.zip' -Force"
) else (
    echo [WARN] Δεν βρέθηκε dist\OrderMatcher-OCR
)

echo.
echo [DONE] Ολοκληρώθηκε η δημιουργία portable zip.
pause
