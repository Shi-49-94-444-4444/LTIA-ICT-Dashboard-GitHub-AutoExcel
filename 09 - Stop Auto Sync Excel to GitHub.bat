@echo off
title LTIA Dashboard - Stop Excel Auto Sync
powershell.exe -NoProfile -Command "$p=Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*github_excel_sync.ps1*' }; if($p){$p | ForEach-Object { Stop-Process -Id $_.ProcessId -Force; Write-Host ('Stopped PID '+$_.ProcessId)} } else { Write-Host 'No LTIA Excel sync process found.' }"
pause
