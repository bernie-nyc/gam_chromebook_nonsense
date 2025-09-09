<# 
  Apply-ChildOU-LoginRestriction.ps1  (PowerShell 5.1)

  1) GAM -> print orgs (capture output)
  2) Trim to CSV header+rows
  3) Filter to DIRECT children of the base OU
  4) For each child OU: set Chrome Sign-In restriction to allow ONLY that OU name (email)

  Example:
    .\Apply-ChildOU-LoginRestriction.ps1 `
      -BaseOUPath "\Student 1:1 Devices\00 - WLS\GR05" `
      -GamPath "C:\GAM7\gam.exe" -DryRun
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$BaseOUPath,

  [Parameter(Mandatory=$true)]
  [string]$GamPath,

  [string]$CsvPath = ".\child_ous.csv",

  [switch]$DryRun
)

function Now { Get-Date -Format "HH:mm:ss" }
function Log([string]$msg) { Write-Host ("[{0}] {1}" -f (Now), $msg) }

# Normalize base OU to /A/B/C
$base = $BaseOUPath.Trim() -replace '\\','/'
if ($base -notmatch '^/') { $base = "/$base" }
$parts = $base -split '/'
$parts = $parts | Where-Object { $_ -ne '' }
$base  = '/' + ($parts -join '/')

# Resolve CSV path (PS5-safe)
$resolvedCsv = $null
try { $resolvedCsv = (Resolve-Path -LiteralPath $CsvPath -ErrorAction Stop).Path } catch {}
if (-not $resolvedCsv) { $resolvedCsv = $CsvPath }

$mode = if ($DryRun) { "DRY-RUN (no changes will be made)" } else { "LIVE (changes WILL be made)" }

Write-Host ("Base OU: {0}" -f $base)
Write-Host ("GAM   : {0}" -f $GamPath)
Write-Host ("CSV   : {0}" -f $resolvedCsv)
Write-Host ("Mode  : {0}" -f $mode)
Write-Host ""

# 1) Export ALL OUs (capture + clean)
Log "[1/3] Exporting all OUs (this can take a bit on large tenants)..."
$raw = & $GamPath print orgs fields orgUnitPath,parentOrgUnitPath 2>&1
if (-not $raw -or $raw.Count -eq 0) { throw "GAM returned no output for 'print orgs'." }

# locate header
$headerIndex = -1
for ($i=0; $i -lt $raw.Count; $i++) {
  if ($raw[$i] -match '^\s*orgUnitPath\s*,\s*parentOrgUnitPath\s*$') { $headerIndex = $i; break }
}
if ($headerIndex -lt 0) {
  for ($i=0; $i -lt $raw.Count; $i++) {
    if ($raw[$i] -match 'orgUnitPath' -and $raw[$i] -match 'parentOrgUnitPath') { $headerIndex = $i; break }
  }
}
if ($headerIndex -lt 0) {
  $preview = ($raw | Select-Object -First 10) -join "`r`n"
  throw "Couldn't locate CSV header in GAM output. First lines:`r`n$preview"
}

$csvLines = $raw[$headerIndex..($raw.Count-1)]
Set-Content -LiteralPath $CsvPath -Value ($csvLines -join "`r`n") -Encoding UTF8
$resolvedCsv = (Resolve-Path -LiteralPath $CsvPath).Path
Log ("CSV written: {0}" -f $resolvedCsv)

# 2) Filter to DIRECT children
$rows = Import-Csv -LiteralPath $CsvPath
if (-not $rows) { throw "CSV is empty/unreadable: $CsvPath" }
$children = $rows | Where-Object { $_.parentOrgUnitPath -eq $base } | Sort-Object orgUnitPath
$childCount = ($children | Measure-Object).Count
Log ("[2/3] Found {0} direct child OUs under {1}" -f $childCount, $base)
if ($childCount -eq 0) { Log "Nothing to do."; return }

# 3) Apply policy per child OU
# Required GAM syntax (note quotes around OU path and email):
#   gam update chromepolicy `
#     orgunit "<OU PATH>" `
#     chrome.devices.SignInRestriction `
#     deviceallownewusers RESTRICTED_LIST `
#     userallowlist "<email>"

$idx = 0
$failures = @()
foreach ($child in $children) {
  $idx++
  $ouPath = $child.orgUnitPath
  $email  = ($ouPath.TrimEnd('/') -split '/')[-1]

  Log ("[3/3] ({0}/{1}) Target OU: {2}" -f $idx, $childCount, $ouPath)
  Log ("          Replace allowlist with: {0}" -f $email)

  # Build ONE quoted argument string so GAM sees spaces correctly
  $argLine = ('update chromepolicy orgunit "{0}" chrome.devices.SignInRestriction deviceallownewusers RESTRICTED_LIST userallowlist "{1}"' -f $ouPath, $email)

  if ($DryRun) {
    Log ("DRY-RUN CMD > {0} {1}" -f $GamPath, $argLine)
    continue
  }

  $errTmp = [System.IO.Path]::GetTempFileName()
  try {
    $p = Start-Process -FilePath $GamPath -ArgumentList $argLine -NoNewWindow -PassThru -RedirectStandardError $errTmp
    $null = $p.WaitForExit(120000)
    if (-not $p.HasExited) { try { $p.Kill() } catch {}; throw "GAM update timed out." }
    if ($p.ExitCode -ne 0) {
      $err = ""; if (Test-Path $errTmp) { $err = Get-Content $errTmp -Raw }
      Log ("WARNING:   -> GAM returned non-zero exit code on {0} (exit {1})" -f $ouPath, $p.ExitCode)
      if ($err) { Log ("           STDERR: {0}" -f $err.Trim()) }
      $failures += [PSCustomObject]@{ OU=$ouPath; Email=$email; ExitCode=$p.ExitCode; Error=$err }
    } else {
      Log ("SUCCESS:   -> Updated policy on {0}" -f $ouPath)
    }
  } finally {
    if (Test-Path $errTmp) { Remove-Item $errTmp -Force }
  }
}

Write-Host ""
if ($DryRun) {
  Log "DONE (DRY-RUN). No changes were made."
} else {
  if ($failures.Count -gt 0) {
    Log ("Completed with {0} failure(s)." -f $failures.Count)
    $failures | Format-Table -AutoSize
  } else {
    Log "DONE. All child OUs updated successfully."
  }
}
