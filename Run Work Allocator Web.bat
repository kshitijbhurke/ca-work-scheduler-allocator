@echo off
cd /d "%~dp0"

rem Use python, or the py launcher if python isn't on PATH
set "PY=python"
python --version >nul 2>nul
if errorlevel 1 (
    set "PY=py"
    py --version >nul 2>nul
    if errorlevel 1 (
        echo Python 3 was not found.
        echo Install it from https://www.python.org/downloads/ and run this file again.
        pause
        exit /b 1
    )
)

rem Install the required packages on first use
%PY% -c "import openpyxl, reportlab, streamlit" >nul 2>nul
if errorlevel 1 (
    echo Installing required packages. This can take a few minutes...
    %PY% -m pip install -r requirements-web.txt
    if errorlevel 1 (
        echo The packages could not be installed. Check the internet connection and try again.
        pause
        exit /b 1
    )
)

echo The app opens in your browser. Keep this window open while you use it.
%PY% -m streamlit run "web\Work_Allocator_Web.py" --server.port 8502
pause
