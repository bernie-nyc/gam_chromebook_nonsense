#.\Apply-ChildOU-LoginRestriction.ps1 -BaseOUPath "\Student 1:1 Devices\00 - WLS\GR05" -GamPath "C:\GAM7\gam.exe" -DryRun
#
<# 
  Apply-ChildOU-LoginRestriction.ps1
  ----------------------------------------------------------
  WHAT THIS DOES (in plain English):
  1) We take a parent OU path you give us (e.g. \Student 1:1 Devices\00 - WLS\GR05).
  2) We ask GAM for ALL org units, dump them to a CSV locally.
  3) We filter that CSV to just the DIRECT children of your parent OU.
  4) For each child OU, we grab the LAST PATH SEGMENT (that's the email address).
  5) We push a Chrome policy to that child OU that:
        - Sets deviceallownewusers to RESTRICTED_LIST
        - Sets userallowlist to "<that email>"
  6) If -DryRun is on, we only PRINT what we would do. No changes.
#>

[CmdletBinding()]
param(
  # Parent OU path you want to process. You can pass with backslashes like "\A\B\C".
  [Parameter(Mandatory=$true)]
  [string]$BaseOUPath,

  # Full path to gam.exe (or just "gam" if it's on PATH)
  [Parameter(Mandatory=$true)]
  [string]$GamPath,

  # Where to write the big "all orgs" CSV we’ll filter
  [string]$CsvPath = ".\child_ous.csv",

  # If present, we only SHOW what we would do.
  [switch]$DryRun
)

### --------- Helper: timestamped log -----------
function Now { Get-Date -Format "HH:mm:ss" }
function Log([string]$msg) { Write-Host ("[{0}] {1}" -f (Now), $msg) }

### --------- Normalize the base OU -------------
# Accept "\A\B\C" or "/A/B/C" and normalize to "/A/B/C"
$base = $BaseOUPath.Trim()
$base = $base -replace '\\','/'          # backslashes -> slashes
if ($base -notmatch '^/') { $base = "/$base" } # ensure leading slash
# collapse accidental double slashes
$base = ($base -split '/' | Where-Object { $_ -ne '' }) -join '/'
$base = "/$base"

Write-Host ("Base OU: {0}" -f $base)
Write-Host ("GAM   : {0}" -f $GamPath)
Write-Host ("CSV   : {0}" -f (Resolve-Path -LiteralPath $CsvPath -ErrorAction SilentlyContinue) ?? $CsvPath)
Write-Host ("Mode  : {0}" -f ($(if($DryRun){"DRY-RUN (no changes will be made)"}else{"LIVE (changes WILL be made)"}))
Write-Host ""

### --------- 1) Export ALL OUs to CSV ----------
Log "[1/3] Exporting all OUs (this can take a bit on large tenants)..."

# We use Start-Process to avoid pipeline/encoding weirdness and to reliably write a file.
# Command we want: gam print orgs fields orgUnitPath,parentOrgUnitPath
$stdOut = New-TemporaryFile
$stdErr = New-TemporaryFile
try {
  $p = Start-Process -FilePath $GamPath `
                     -ArgumentList @('print','orgs','fields','orgUnitPath,parentOrgUnitPath') `
                     -NoNewWindow -PassThru `
                     -RedirectStandardOutput $stdOut `
                     -RedirectStandardError  $stdErr
  $null = $p.WaitForExit(90000) # wait up to 90 seconds; adjust if needed
  if (-not $p.HasExited) {
    try { $p.Kill() } catch {}
    throw "GAM 'print orgs' timed out."
  }
  if ($p.ExitCode -ne 0) {
    $errText = Get-Content $stdErr -Raw
    throw ("GAM 'print orgs' failed (exit {0})`n{1}" -f $p.ExitCode, $errText)
  }
  # Move stdout to our CSV path
  Move-Item -Force $stdOut $CsvPath
} finally {
  # Clean up the temp stderr; stdout may have been moved
  if (Test-Path $stdErr) { Remove-Item $stdErr -Force }
  if (Test-Path $stdOut) { Remove-Item $stdOut -Force -ErrorAction SilentlyContinue }
}

if (-not (Test-Path $CsvPath)) { throw "Failed to produce CSV at $CsvPath" }
Log ("CSV written: {0}" -f (Resolve-Path $CsvPath).Path)

### --------- 2) Filter to DIRECT children -------
# CSV headers from GAM are: orgUnitPath,parentOrgUnitPath
$rows = Import-Csv -LiteralPath $CsvPath
if (-not $rows) { throw "CSV is empty/unreadable: $CsvPath" }

# Keep only direct children (where parentOrgUnitPath == $base)
$children = $rows | Where-Object { $_.parentOrgUnitPath -eq $base } | Sort-Object orgUnitPath

$childCount = ($children | Measure-Object).Count
Log ("[2/3] Found {0} direct child OUs under {1}" -f $childCount, $base)
if ($childCount -eq 0) {
  Log "Nothing to do."
  return
}

### --------- 3) Apply policy per child OU -------
# Working syntax you verified:
# gam update chromepolicy orgunit "<CHILD_OU_PATH>" chrome.devices.SignInRestriction deviceallownewusers RESTRICTED_LIST userallowlist "<EMAIL>"

$index = 0
$failures = @()
foreach ($child in $children) {
  $index++
  $ouPath  = $child.orgUnitPath
  $email   = ($ouPath.TrimEnd('/') -split '/')[-1]  # last segment is the email leaf

  Log ("[3/3] ({0}/{1}) Target OU: {2}" -f $index, $childCount, $ouPath)
  Log ("          Replace allowlist with: {0}" -f $email)

  $args = @(
    'update','chromepolicy',
    'orgunit', $ouPath,
    'chrome.devices.SignInRestriction',
    'deviceallownewusers','RESTRICTED_LIST',
    'userallowlist', $email
  )

  if ($DryRun) {
    Log ("DRY-RUN CMD > {0} {1}" -f $GamPath, ($args -join ' '))
    continue
  }

  $stderrFile = New-TemporaryFile
  try {
    $proc = Start-Process -FilePath $GamPath `
                          -ArgumentList $args `
                          -NoNewWindow -PassThru `
                          -RedirectStandardError $stderrFile
    $null = $proc.WaitForExit(90000) # 90s per OU; bump if needed
    if (-not $proc.HasExited) {
      try { $proc.Kill() } catch {}
      throw "GAM update timed out."
    }
    if ($proc.ExitCode -ne 0) {
      $err = Get-Content $stderrFile -Raw
      Log ("WARNING:   -> GAM returned non-zero exit code on {0} (exit {1})" -f $ouPath, $proc.ExitCode)
      if ($err) { Log ("           STDERR: {0}" -f $err.Trim()) }
      $failures += [PSCustomObject]@{ OU = $ouPath; Email = $email; ExitCode = $proc.ExitCode; Error = $err }
    } else {
      Log ("SUCCESS:   -> Updated policy on {0}" -f $ouPath)
    }
  } finally {
    if (Test-Path $stderrFile) { Remove-Item $stderrFile -Force }
  }
}

### --------- Summary ----------------------------
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

#Tip: if you see timeouts on very large tenants, increase the WaitForExit(90000) values (they’re in milliseconds) for both the export step and the per-OU update loop.

