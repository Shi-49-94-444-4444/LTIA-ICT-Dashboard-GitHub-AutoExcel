@echo off
title LTIA Dashboard - Excel Auto Sync to GitHub
cd /d "%~dp0"

where git >nul 2>&1
if errorlevel 1 if not exist "C:\Program Files\Git\cmd\git.exe" if not exist "%LOCALAPPDATA%\Programs\Git\cmd\git.exe" (
  echo [ERROR] Git for Windows was not found.
  echo Install Git for Windows first, then close and reopen this BAT.
  pause
  exit /b 1
)

if not exist "ICT-RFA-RFI-lists-New.xlsx" (
  echo [ERROR] ICT-RFA-RFI-lists-New.xlsx was not found in this folder.
  pause
  exit /b 1
)

if not exist "github_excel_sync.ps1" (
  echo [ERROR] github_excel_sync.ps1 is missing.
  pause
  exit /b 1
)

echo.
echo =============================================
echo LTIA Dashboard - Excel Auto Sync to GitHub
echo =============================================
echo Excel changes will be pushed automatically to:
echo   branch: dashboard-data
echo.
echo Keep this window open while working on the Excel file.
echo Close this window to stop the watcher.
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0github_excel_sync.ps1"
pause
