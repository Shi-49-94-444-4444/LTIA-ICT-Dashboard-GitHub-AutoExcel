LTIA DASHBOARD - GITHUB AUTO EXCEL

Architecture
------------
- main = dashboard code + GitHub Pages workflow
- dashboard-data = Excel data only
- local BAT/PowerShell watcher = pushes the newest workbook to dashboard-data
- browser = polls dashboard-data every 15 seconds

Important: static.yml only deploys on pushes to main. Therefore Excel pushes to dashboard-data do not rebuild GitHub Pages.

Normal daily workflow
---------------------
1. Start 08 - Start Auto Sync Excel to GitHub.bat.
2. Edit ICT-RFA-RFI-lists-New.xlsx.
3. Save the workbook.
4. Wait for the sync window to report that the data branch was updated.
5. Leave the dashboard page open; it checks for a new workbook every 15 seconds.

If the dashboard still shows old data
-------------------------------------
- Check that dashboard-data/ICT-RFA-RFI-lists-New.xlsx has a new commit on GitHub.
- Open browser DevTools > Console and look for Excel loading errors.
- Hard refresh the page once (Ctrl+F5) to rule out an old page instance.

The dashboard no longer falls back to an Excel copy on main, so a broken dashboard-data source produces a visible Excel loading error instead of silently showing stale data.
