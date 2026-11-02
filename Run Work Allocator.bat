@echo off
cd /d "%~dp0"

rem Use python, or the py launcher if python isn't on PATH
set "PY=python"
set "PYW=pythonw"
python --version >nul 2>nul
if errorlevel 1 (
    set "PY=py"
    set "PYW=pyw"
    py --version >nul 2>nul
    if errorlevel 1 (
        echo Python 3 was not found.
        echo Install it from https://www.python.org/downloads/ and run this file again.
        pause
        exit /b 1
    )
)

rem Install the required packages on first use
%PY% -c "import openpyxl, reportlab" >nul 2>nul
if errorlevel 1 (
    echo Installing required packages...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo The packages could not be installed. Check the internet connection and try again.
        pause
        exit /b 1
    )
)

start "" %PYW% "%~dp0Work_Allocator.py"
