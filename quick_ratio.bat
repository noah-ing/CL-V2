@echo off
setlocal enabledelayedexpansion

echo.
echo ============================================
echo   Quick Call Ratio Calculator
echo ============================================
echo.

:: Find Vitelity CDR file
if "%~1"=="" (
    :: Look in current directory and subdirectories
    set "CDR_FILE="
    for /r %%f in (http-*.csv) do (
        set "CDR_FILE=%%f"
        goto :found
    )
    echo ERROR: No CDR file found (http-*.csv)
    echo.
    echo Usage: quick_ratio.bat [cdr_file.csv]
    echo.
    pause
    exit /b 1
) else (
    set "CDR_FILE=%~1"
)

:found
echo Using: !CDR_FILE!
echo.

cd /d "%~dp0"
python call_ratio.py "!CDR_FILE!"

echo.
pause
