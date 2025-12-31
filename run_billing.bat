@echo off
:: Launch PowerShell script (more reliable with OneDrive paths)
powershell -ExecutionPolicy Bypass -File "%~dp0run_billing.ps1"
