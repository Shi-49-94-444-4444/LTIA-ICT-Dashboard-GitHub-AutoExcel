LTIA DASHBOARD - GITHUB AUTO EXCEL MODE

What changed
-------------
1. index.html no longer depends on the old data.js publisher output.
2. The dashboard reads ICT-RFA-RFI-lists-New.xlsx directly in the browser.
3. SheetJS 0.20.3 is loaded from the official SheetJS CDN.
4. On GitHub Pages, the dashboard reads the Excel file from the dashboard-data branch.
5. The Excel file is polled every 15 seconds by the dashboard, so a newly synced workbook appears without redeploying the page.
6. 08 - Start Auto Sync Excel to GitHub.bat watches the Excel file and automatically commits/pushes it to dashboard-data.

One-time setup
--------------
A. Put this folder in your Git repository and configure the origin remote.
B. Push the dashboard code (index.html, dashboard-config.js, excel-data-loader.js, github_excel_sync.ps1, and the BAT files) once to the Pages branch.
C. Enable GitHub Pages from that Pages branch.
D. Start: 08 - Start Auto Sync Excel to GitHub.bat
E. On first run, the script creates/updates the dashboard-data branch and uploads the Excel file there.

After that
----------
Edit and save ICT-RFA-RFI-lists-New.xlsx normally.
The watcher detects the change, waits until the file is stable, then commits/pushes the newest workbook to dashboard-data.
The live dashboard reads that branch directly and refreshes automatically.

Important
---------
- Git credentials must already work for the repository's origin remote.
- The GitHub repository must be publicly readable for the browser to fetch the raw Excel file this way.
- If the site uses a custom domain instead of *.github.io, set dataUrl in dashboard-config.js to the raw Excel URL.
- The old LTIA Dashboard Auto Publisher.exe and Xserver BAT files are retained as legacy files. They are not required for this GitHub auto-Excel flow.
