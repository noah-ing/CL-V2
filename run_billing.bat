@echo off
setlocal enabledelayedexpansion

echo.
echo ============================================================
echo   CirrusLine Telecom Billing Report Generator
echo   v1.1 - December 2024
echo ============================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

:: Set the reports folder (where input files are located)
if "%~1"=="" (
    set "REPORTS_FOLDER=%~dp0Working Reports"
    echo No folder specified, looking in: !REPORTS_FOLDER!
) else (
    set "REPORTS_FOLDER=%~1"
    echo Using folder: !REPORTS_FOLDER!
)

:: Check if folder exists
if not exist "!REPORTS_FOLDER!" (
    echo.
    echo ERROR: Folder not found: !REPORTS_FOLDER!
    echo.
    echo Usage: run_billing.bat [path_to_reports_folder]
    echo Example: run_billing.bat "C:\Users\Dean\Desktop\Working Reports 2025_11"
    echo.
    pause
    exit /b 1
)

:: Create output directory
set "OUTPUT_DIR=%~dp0reports"
if not exist "!OUTPUT_DIR!" mkdir "!OUTPUT_DIR!"

echo.
echo ============================================================
echo   Checking for ZIP files to extract...
echo ============================================================
echo.

:: Extract any ZIP files found in the reports folder
:: This handles SkySwitch reports that come as ZIP files
set "ZIP_COUNT=0"
for %%f in ("!REPORTS_FOLDER!\*.zip") do (
    echo   Extracting: %%~nxf
    powershell -Command "Expand-Archive -Path '%%f' -DestinationPath '!REPORTS_FOLDER!' -Force" 2>nul
    if !errorlevel! equ 0 (
        set /a ZIP_COUNT+=1
    ) else (
        echo   Warning: Could not extract %%~nxf
    )
)

if !ZIP_COUNT! gtr 0 (
    echo.
    echo   Extracted !ZIP_COUNT! ZIP file(s)
)

echo.
echo Looking for input files...
echo.

:: Find Vitelity CDR (http-*.csv)
set "VITELITY_CDR="
for %%f in ("!REPORTS_FOLDER!\http-*.csv") do (
    set "VITELITY_CDR=%%f"
    echo   Found Vitelity CDR: %%~nxf
)

:: Find phone numbers file (phonenumbers__*.csv)
set "PHONE_NUMBERS="
for %%f in ("!REPORTS_FOLDER!\phonenumbers__*.csv") do (
    set "PHONE_NUMBERS=%%f"
    echo   Found Phone Numbers: %%~nxf
)

:: Find SMS file (syneteks-*.csv)
set "SMS_FILE="
for %%f in ("!REPORTS_FOLDER!\syneteks-*.csv") do (
    set "SMS_FILE=%%f"
    echo   Found SMS File: %%~nxf
)

:: Find Domain Statistics (Domain-Statistics-*.xlsx)
set "DOMAIN_STATS="
for %%f in ("!REPORTS_FOLDER!\Domain-Statistics-*.xlsx") do (
    set "DOMAIN_STATS=%%f"
    echo   Found Domain Stats: %%~nxf
)

:: Find Master Excel with SkySwitch CDR (CDR SS records-*.xlsx)
set "MASTER_XLSX="
for %%f in ("!REPORTS_FOLDER!\CDR SS records-*.xlsx") do (
    set "MASTER_XLSX=%%f"
    echo   Found Master Excel: %%~nxf
)

echo.

:: Validate required files
if "!VITELITY_CDR!"=="" (
    echo ERROR: No Vitelity CDR file found (http-*.csv)
    pause
    exit /b 1
)

if "!PHONE_NUMBERS!"=="" (
    echo ERROR: No phone numbers file found (phonenumbers__*.csv)
    pause
    exit /b 1
)

echo ============================================================
echo   Running Billing Reports...
echo ============================================================
echo.

:: Change to script directory
cd /d "%~dp0"

:: Run Python with all arguments (Python handles empty strings gracefully)
python billing_reports.py "!VITELITY_CDR!" "!PHONE_NUMBERS!" "!OUTPUT_DIR!" "!SMS_FILE!" "!DOMAIN_STATS!" "!MASTER_XLSX!"

:: Check if Python ran successfully
if errorlevel 1 (
    echo.
    echo ERROR: Python script failed. See error message above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Reports saved to: !OUTPUT_DIR!
echo ============================================================
echo.
echo Generated files:
dir /b "!OUTPUT_DIR!\*.csv" 2>nul
echo.

pause
