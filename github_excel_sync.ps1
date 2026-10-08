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
$Remote = 'https://github.com/Shi-49-94-444-4444/LTIA-ICT-Dashboard-GitHub-AutoExcel.git'

if ($DataBranch -ne 'dashboard-data') {
  throw 'This sync script is locked to the dashboard-data branch.'
}

function Get-GitExe {
  $candidates = @(
    (Get-Command git.exe -ErrorAction SilentlyContinue).Source,
    'C:\Program Files\Git\cmd\git.exe',
    'C:\Program Files\Git\bin\git.exe',
    "$env:LOCALAPPDATA\Programs\Git\cmd\git.exe"
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path $_) }

  if (-not $candidates) {
    throw 'Git for Windows was not found. Install Git for Windows and reopen this BAT.'
  }

  return $candidates[0]
}

$GitExe = Get-GitExe

# Compatible with Windows PowerShell 5.1 and PowerShell 7+.
# Do NOT use ProcessStartInfo.ArgumentList here because it is unavailable in Windows PowerShell 5.1.
function RunGit([string[]]$GitArgs) {
  # Run Git and print its output, but DO NOT return the output from this
  # helper. PowerShell captures function output when a function is used
  # inside another function, which can corrupt values such as $SyncRepo.
  $oldErrorAction = $ErrorActionPreference
  try {
    $ErrorActionPreference = 'Continue'
    $output = & $GitExe @GitArgs 2>&1
    $code = $LASTEXITCODE
  }
  finally {
    $ErrorActionPreference = $oldErrorAction
  }

  $text = (($output | ForEach-Object { $_.ToString() }) -join "`n").Trim()

  if ($code -ne 0) {
    if ([string]::IsNullOrWhiteSpace($text)) { $text = '(git returned no diagnostic output)' }
    throw "git $($GitArgs -join ' ') failed with exit code $code`n$text"
  }

  if (-not [string]::IsNullOrWhiteSpace($text)) {
    Write-Host $text
  }
}

function RunGitText([string[]]$GitArgs) {
  # Same Git execution as RunGit, but intentionally returns text because
  # the caller needs ls-remote output to detect the remote branch.
  $oldErrorAction = $ErrorActionPreference
  try {
    $ErrorActionPreference = 'Continue'
    $output = & $GitExe @GitArgs 2>&1
    $code = $LASTEXITCODE
  }
  finally {
    $ErrorActionPreference = $oldErrorAction
  }

  $text = (($output | ForEach-Object { $_.ToString() }) -join "`n").Trim()

  if ($code -ne 0) {
    if ([string]::IsNullOrWhiteSpace($text)) { $text = '(git returned no diagnostic output)' }
    throw "git $($GitArgs -join ' ') failed with exit code $code`n$text"
  }

  return $text
}

function EnsureSyncRepo {
  New-Item -ItemType Directory -Force -Path $SyncRoot | Out-Null

  if (-not (Test-Path (Join-Path $SyncRepo '.git'))) {
    if (Test-Path $SyncRepo) { Remove-Item -LiteralPath $SyncRepo -Recurse -Force }
    Write-Host 'Cloning repository...'
    RunGit @('clone', $Remote, $SyncRepo)
  }

  Write-Host 'Fetching origin...'
  RunGit @('-C', $SyncRepo, 'fetch', '--prune', 'origin')

  $remoteBranch = RunGitText @('-C', $SyncRepo, 'ls-remote', '--heads', 'origin', "refs/heads/$DataBranch")

  if (-not [string]::IsNullOrWhiteSpace($remoteBranch)) {
    Write-Host "Using existing remote branch: $DataBranch"
    RunGit @('-C', $SyncRepo, 'checkout', '-B', $DataBranch, "origin/$DataBranch")
    RunGit @('-C', $SyncRepo, 'reset', '--hard', "origin/$DataBranch")
    RunGit @('-C', $SyncRepo, 'clean', '-fd')
  }
  else {
    Write-Host "Remote branch does not exist yet. Creating: $DataBranch"
    RunGit @('-C', $SyncRepo, 'checkout', '-B', $DataBranch, 'origin/main')
    RunGit @('-C', $SyncRepo, 'rm', '-r', '--ignore-unmatch', '--', '.')
    RunGit @('-C', $SyncRepo, 'clean', '-fd')
  }

  $rootName = (& $GitExe -C $ProjectRoot config user.name 2>$null | Out-String).Trim()
  $rootEmail = (& $GitExe -C $ProjectRoot config user.email 2>$null | Out-String).Trim()
  if ([string]::IsNullOrWhiteSpace($rootName)) { $rootName = 'LTIA Dashboard Auto Sync' }
  if ([string]::IsNullOrWhiteSpace($rootEmail)) { $rootEmail = 'ltia-dashboard-sync@users.noreply.github.com' }

  RunGit @('-C', $SyncRepo, 'config', 'user.name', $rootName)
  RunGit @('-C', $SyncRepo, 'config', 'user.email', $rootEmail)

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

  RunGit @('-C', $syncRepo, 'add', '-f', '--', $ExcelFile, 'Data Last Updated.txt')

  & $GitExe -C $syncRepo diff --cached --quiet 2>$null
  $diffCode = $LASTEXITCODE
  if ($diffCode -eq 0) {
    return $false
  }

  $message = "Auto-sync Excel $($source.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
  RunGit @('-C', $syncRepo, 'commit', '-m', $message)

  Write-Host 'Pushing dashboard-data to GitHub...'
  try {
    RunGit @('-C', $syncRepo, 'push', '--set-upstream', 'origin', $DataBranch)
  }
  catch {
    throw @"
PUSH FAILED.
GitHub rejected the push. Most commonly this means GitHub authentication/permission is not set up for this PC.

Try once in a normal CMD window:
  git -C "$syncRepo" push --set-upstream origin $DataBranch

If GitHub asks you to sign in, complete the browser sign-in, then run this BAT again.

Detailed error:
$($_.Exception.Message)
"@
  }

  return $true
}

Write-Host '============================================='
Write-Host ' LTIA Dashboard - Excel Auto Sync to GitHub'
Write-Host '============================================='
Write-Host "Excel : $ExcelPath"
Write-Host "Branch: $DataBranch"
Write-Host 'Repo  : Shi-49-94-444-4444/LTIA-ICT-Dashboard-GitHub-AutoExcel'
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
