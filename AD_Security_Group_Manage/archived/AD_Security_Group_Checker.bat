@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "AD_Security_Group_Checker_v2.5.ps1"
pause
