@echo off
REM Startet den BMP-Scanner (Honeywell im USB-Serial-Modus) minimiert.
REM Diese Datei und BMP-Scanner.ps1 muessen im selben Ordner liegen.
start "BMP-Scanner" /min powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0BMP-Scanner.ps1"
