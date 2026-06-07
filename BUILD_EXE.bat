@echo off
setlocal
title Build Master Sheet Finder Pro EXE
cd /d "%~dp0"

echo.
echo Building Master Sheet Finder Pro as a Windows EXE...
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PY_CMD=py -3
) else (
    set PY_CMD=python
)

if not exist .venv (
    %PY_CMD% -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

if exist app.ico (
    pyinstaller --noconfirm --onefile --windowed --name MasterSheetFinderPro --icon app.ico app.pyw
) else (
    pyinstaller --noconfirm --onefile --windowed --name MasterSheetFinderPro app.pyw
)

echo.
echo Done.
echo EXE location: %cd%\dist\MasterSheetFinderPro.exe
echo.
pause
