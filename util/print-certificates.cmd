@echo off
REM Helper to run PowerShell script with relaxed policy for this invocation only.
setlocal
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0print-certificates.ps1" %*
endlocal
