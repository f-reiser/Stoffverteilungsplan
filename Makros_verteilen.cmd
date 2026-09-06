@echo off
rem  Stoffverteilungsplan - Makros in alle Mappen dieses Ordners uebertragen.
rem  Doppelklick genuegt.
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Makros_verteilen.ps1"
echo.
pause
