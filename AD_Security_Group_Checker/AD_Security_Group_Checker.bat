@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "AD_Security_Group_Checker.ps1"
pause
