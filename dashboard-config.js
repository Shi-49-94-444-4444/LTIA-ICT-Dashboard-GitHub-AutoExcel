/*
 * LTIA Dashboard data source configuration.
 *
 * The default GitHub Pages setup is automatic:
 *   - dashboard code lives on the normal Pages branch (usually main)
 *   - the Excel file is auto-synced to the dashboard-data branch
 *   - the page reads the Excel file directly from that data branch
 *   - Excel changes therefore do NOT require a new Pages deployment
 *
 * For a custom domain or another host, set dataUrl to the raw Excel URL.
 */
window.LTIA_DASHBOARD_CONFIG = {
  excelFile: 'ICT-RFA-RFI-lists-New.xlsx',
  dataBranch: 'dashboard-data',
  fallbackBranch: 'main',
  dataUrl: '',
  refreshMs: 15000
};
