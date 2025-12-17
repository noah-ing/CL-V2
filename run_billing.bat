@echo off
setlocal enabledelayedexpansion

echo.
echo ============================================================
echo   CirrusLine Telecom Billing Report Generator
echo   v1.0 - December 2024
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

:: Build the command
set "CMD=python billing_reports.py"
set "CMD=!CMD! "!VITELITY_CDR!""
set "CMD=!CMD! "!PHONE_NUMBERS!""
set "CMD=!CMD! "!OUTPUT_DIR!""

if not "!SMS_FILE!"=="" (
    set "CMD=!CMD! "!SMS_FILE!""
) else (
    set "CMD=!CMD! """
)

if not "!DOMAIN_STATS!"=="" (
    set "CMD=!CMD! "!DOMAIN_STATS!""
) else (
    set "CMD=!CMD! """
)

if not "!MASTER_XLSX!"=="" (
    set "CMD=!CMD! "!MASTER_XLSX!""
)

:: Run the command
cd /d "%~dp0"
!CMD!

echo.
echo ============================================================
echo   Reports saved to: !OUTPUT_DIR!
echo ============================================================
echo.
echo Generated files:
dir /b "!OUTPUT_DIR!\*.csv" 2>nul
echo.

pause
