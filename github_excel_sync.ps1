param(
  [int]$CheckSeconds = 5,
  [int]$StableSeconds = 3,
  [string]$DataBranch = 'dashboard-data',
  [string]$ExcelFile = 'ICT-RFA-RFI-lists-New.xlsx'
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExcelPath = Join-Path $ProjectRoot $ExcelFile
$SyncRoot = Join-Path $env:LOCALAPPDATA 'LTIA Dashboard\Git Data Sync'
$SyncRepo = Join-Path $SyncRoot 'repo'

function Git([string[]]$Args) {
  & git @Args
  if ($LASTEXITCODE -ne 0) { throw "git $($Args -join ' ') failed with exit code $LASTEXITCODE" }
}

function GitText([string[]]$Args) {
  $result = & git @Args 2>&1
  if ($LASTEXITCODE -ne 0) { throw "git $($Args -join ' ') failed: $($result -join ' ')" }
  return ($result -join "`n").Trim()
}

function EnsureSyncRepo {
  $remote = GitText @('-C', $ProjectRoot, 'remote', 'get-url', 'origin')
  if ([string]::IsNullOrWhiteSpace($remote)) { throw 'No origin remote is configured in this project.' }

  New-Item -ItemType Directory -Force -Path $SyncRoot | Out-Null

  if (-not (Test-Path (Join-Path $SyncRepo '.git'))) {
    if (Test-Path $SyncRepo) { Remove-Item -LiteralPath $SyncRepo -Recurse -Force }
    Git @('clone', $remote, $SyncRepo)
  }

  Git @('-C', $SyncRepo, 'fetch', 'origin')

  $remoteBranch = & git -C $SyncRepo ls-remote --heads origin $DataBranch 2>$null
  $current = GitText @('-C', $SyncRepo, 'rev-parse', '--abbrev-ref', 'HEAD')

  if ($remoteBranch) {
    if ($current -ne $DataBranch) {
      & git -C $SyncRepo checkout $DataBranch 2>$null | Out-Null
      if ($LASTEXITCODE -ne 0) { Git @('-C', $SyncRepo, 'checkout', '-B', $DataBranch, "origin/$DataBranch") }
    }
    Git @('-C', $SyncRepo, 'reset', '--hard', "origin/$DataBranch")
    Git @('-C', $SyncRepo, 'clean', '-fd')
  }
  else {
    if ($current -ne $DataBranch) {
      Git @('-C', $SyncRepo, 'checkout', '--orphan', $DataBranch)
    }
    & git -C $SyncRepo rm -rf . 2>$null | Out-Null
    Git @('-C', $SyncRepo, 'clean', '-fd')
  }

  $rootName = (& git -C $ProjectRoot config user.name 2>$null).Trim()
  $rootEmail = (& git -C $ProjectRoot config user.email 2>$null).Trim()
  if ([string]::IsNullOrWhiteSpace($rootName)) { $rootName = 'LTIA Dashboard Auto Sync' }
  if ([string]::IsNullOrWhiteSpace($rootEmail)) { $rootEmail = 'ltia-dashboard-sync@users.noreply.github.com' }
  & git -C $SyncRepo config user.name $rootName
  & git -C $SyncRepo config user.email $rootEmail

  return $SyncRepo
}

function SyncExcel {
  if (-not (Test-Path $ExcelPath)) { throw "Excel file not found: $ExcelPath" }

  $syncRepo = EnsureSyncRepo
  $targetExcel = Join-Path $syncRepo $ExcelFile
  $targetDir = Split-Path -Parent $targetExcel
  New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
  Copy-Item -LiteralPath $ExcelPath -Destination $targetExcel -Force

  $source = Get-Item -LiteralPath $ExcelPath
  $stamp = $source.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
  $meta = Join-Path $syncRepo 'Data Last Updated.txt'
  @(
    "Excel modified: $stamp"
    "Synced to GitHub branch: $DataBranch"
    "Synced at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
  ) | Set-Content -LiteralPath $meta -Encoding UTF8

  Git @('-C', $syncRepo, 'add', '-f', '--', $ExcelFile, 'Data Last Updated.txt')
  & git -C $syncRepo diff --cached --quiet
  if ($LASTEXITCODE -eq 0) {
    return $false
  }

  $message = "Auto-sync Excel $($source.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
  Git @('-C', $syncRepo, 'commit', '-m', $message)
  Git @('-C', $syncRepo, 'push', '-u', 'origin', $DataBranch)
  return $true
}

Write-Host '============================================='
Write-Host ' LTIA Dashboard - Excel Auto Sync to GitHub'
Write-Host '============================================='
Write-Host "Excel : $ExcelPath"
Write-Host "Branch: $DataBranch"
if ($DataBranch -ne 'dashboard-data') { throw 'This sync script is configured for the dashboard-data branch. Do not change DataBranch.' }
Write-Host "Poll  : every $CheckSeconds sec (stable for $StableSeconds sec)"
Write-Host ''

$lastSignature = ''
$firstRun = $true

while ($true) {
  try {
    $file = Get-Item -LiteralPath $ExcelPath
    $signature = "$($file.Length)|$($file.LastWriteTimeUtc.Ticks)"

    if ($firstRun -or $signature -ne $lastSignature) {
      Start-Sleep -Seconds $StableSeconds
      $file2 = Get-Item -LiteralPath $ExcelPath
      $signature2 = "$($file2.Length)|$($file2.LastWriteTimeUtc.Ticks)"

      if ($signature2 -eq $signature) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Excel change detected. Syncing..."
        $changed = SyncExcel
        if ($changed) {
          Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] GitHub data branch updated."
        } else {
          Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] No content change to push."
        }
        $lastSignature = $signature2
        $firstRun = $false
      }
    }
  }
  catch {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Start-Sleep -Seconds ([Math]::Max(5, $CheckSeconds))
  }

  Start-Sleep -Seconds $CheckSeconds
}
