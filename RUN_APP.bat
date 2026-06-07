@echo off
title Master Sheet Finder Pro
cd /d "%~dp0"
py -3 app.pyw
if errorlevel 1 python app.pyw
pause
