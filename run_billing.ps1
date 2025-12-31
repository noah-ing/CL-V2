# CirrusLine Telecom Billing Report Generator
# PowerShell version - more reliable with OneDrive paths

Write-Host ""
Write-Host "============================================================"
Write-Host "  CirrusLine Telecom Billing Report Generator"
Write-Host "  v2.0 - December 2025"
Write-Host "============================================================"
Write-Host ""

# Check Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion"
} catch {
    Write-Host "ERROR: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python from https://www.python.org/downloads/"
    Read-Host "Press Enter to exit"
    exit 1
}

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Ask for reports folder
Write-Host ""
Write-Host "Where are your report files located?"
Write-Host "  1. Working Reports folder (in this directory)"
Write-Host "  2. Enter a custom path"
Write-Host ""
$choice = Read-Host "Enter 1 or 2"

if ($choice -eq "2") {
    $reportsFolder = Read-Host "Enter the full path to your reports folder"
} else {
    $reportsFolder = Join-Path $scriptDir "Working Reports"
}

# Check folder exists
if (-not (Test-Path $reportsFolder)) {
    Write-Host "ERROR: Folder not found: $reportsFolder" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Using folder: $reportsFolder"
Write-Host ""

# Create output directory
$outputDir = Join-Path $scriptDir "reports"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Extract ZIP files
Write-Host "============================================================"
Write-Host "  Checking for ZIP files to extract..."
Write-Host "============================================================"
Write-Host ""

$zipFiles = Get-ChildItem -Path $reportsFolder -Filter "*.zip" -ErrorAction SilentlyContinue
$zipCount = 0
foreach ($zip in $zipFiles) {
    Write-Host "  Extracting: $($zip.Name)"
    try {
        Expand-Archive -Path $zip.FullName -DestinationPath $reportsFolder -Force
        $zipCount++
    } catch {
        Write-Host "  Warning: Could not extract $($zip.Name)" -ForegroundColor Yellow
    }
}

if ($zipCount -gt 0) {
    Write-Host ""
    Write-Host "  Extracted $zipCount ZIP file(s)"
}

Write-Host ""
Write-Host "Looking for input files..."
Write-Host ""

# Find required files
$vitelityCdr = Get-ChildItem -Path $reportsFolder -Filter "http-*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
$phoneNumbers = Get-ChildItem -Path $reportsFolder -Filter "phonenumbers__*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
$smsFile = Get-ChildItem -Path $reportsFolder -Filter "syneteks-*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
$domainStats = Get-ChildItem -Path $reportsFolder -Filter "Domain-Statistics-*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1
$masterXlsx = Get-ChildItem -Path $reportsFolder -Filter "CDR SS records-*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1

# Display found files
if ($vitelityCdr) { Write-Host "  Found Vitelity CDR: $($vitelityCdr.Name)" }
if ($phoneNumbers) { Write-Host "  Found Phone Numbers: $($phoneNumbers.Name)" }
if ($smsFile) { Write-Host "  Found SMS File: $($smsFile.Name)" }
if ($domainStats) { Write-Host "  Found Domain Stats: $($domainStats.Name)" }
if ($masterXlsx) { Write-Host "  Found Master Excel: $($masterXlsx.Name)" }

Write-Host ""

# Validate required files
if (-not $vitelityCdr) {
    Write-Host "ERROR: No Vitelity CDR file found (http-*.csv)" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not $phoneNumbers) {
    Write-Host "ERROR: No phone numbers file found (phonenumbers__*.csv)" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "============================================================"
Write-Host "  Running Billing Reports..."
Write-Host "============================================================"
Write-Host ""

# Build arguments for Python
$pythonScript = Join-Path $scriptDir "billing_reports.py"

# Build argument list (use empty string for missing optional files)
$smsArg = if ($smsFile) { $smsFile.FullName } else { "" }
$domainArg = if ($domainStats) { $domainStats.FullName } else { "" }
$masterArg = if ($masterXlsx) { $masterXlsx.FullName } else { "" }

Write-Host "Running: python billing_reports.py ..."
Write-Host ""

# Run Python directly with arguments (handles spaces in paths correctly)
& python $pythonScript $vitelityCdr.FullName $phoneNumbers.FullName $outputDir $smsArg $domainArg $masterArg

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: Python script failed with exit code $LASTEXITCODE" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "============================================================"
Write-Host "  Reports saved to: $outputDir"
Write-Host "============================================================"
Write-Host ""
Write-Host "Generated files:"
Get-ChildItem -Path $outputDir -Filter "*.csv" | ForEach-Object { Write-Host "  $($_.Name)" }
Write-Host ""

Read-Host "Press Enter to exit"
