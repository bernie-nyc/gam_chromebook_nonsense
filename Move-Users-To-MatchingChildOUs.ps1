
# Preview only
#.\Move-Users-To-MatchingChildOUs.ps1 `
 # -BaseOUPath "\Student 1:1 Devices\00 - WLS\GR05" `
  #-GamPath "C:\GAM7\gam.exe" `
  #-DryRun

# Execute for real
#.\Move-Users-To-MatchingChildOUs.ps1 `
 # -BaseOUPath "\Student 1:1 Devices\00 - WLS\GR05" `
  #-GamPath "C:\GAM7\gam.exe"


param(
  # Parent container to scan (can be \ or / separators). We only act on its immediate children.
  [Parameter(Mandatory=$true)]
  [string]$BaseOUPathRaw,                          # e.g. "\Student 1:1 Devices\00 - WLS\GR05"

  # Path to gam.exe (or just "gam" if on PATH)
  [string]$GamPath = "gam",                        # e.g. "C:\GAM7\gam.exe"

  # Where we dump the “immediate children” list for traceability
  [string]$OutCsv  = ".\child_ous.csv",

  # Preview only (don’t make changes)
  [switch]$DryRun
)

# ============================
# helpers
# ============================
function Normalize-OU([string]$p) {
  $p = $p.Trim()
  $p = $p -replace '\\','/'         # backslash -> forward slash
  if (-not $p.StartsWith('/')) { $p = '/' + $p }
  return $p
}

function Run-GamArray([string[]]$ArgsArray, [switch]$Quiet) {
  if ($DryRun) {
    if (-not $Quiet) { Write-Host "[DRYRUN] $GamPath $($ArgsArray -join ' ')" }
    return @()
  }
  $out = & $GamPath @ArgsArray 2>&1
  $global:LASTEXITCODE | Out-Null
  return $out
}

# Simple email-ish check so we can warn early
function Looks-LikeEmail([string]$s) {
  return [bool]($s -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

# ============================
# banner
# ============================
$BaseOUPath = Normalize-OU $BaseOUPathRaw
Write-Host "Base OU: $BaseOUPath"
Write-Host "GAM   : $GamPath"
Write-Host "CSV   : $OutCsv"
if ($DryRun) { Write-Host "Mode  : DRY-RUN (no changes will be made)" }

# ============================
# 1) Export ALL OUs, keep only immediate children of base OU
# ============================
Write-Host "`n[1/3] Exporting all OUs..."
$argsPrintOrgs = @('print','orgs','fields','name,orgUnitPath,parentOrgUnitPath')
$rawLines = Run-GamArray -ArgsArray $argsPrintOrgs -Quiet

if (-not $DryRun) {
  if ($LASTEXITCODE -ne 0) { throw "gam print orgs failed (exit $LASTEXITCODE)" }
  if (-not $rawLines -or $rawLines.Count -eq 0) { throw "gam print orgs returned no output" }
}

# Keep only CSV header + rows; drop progress chatter
$csvLines = if ($DryRun) { @() } else {
  $rawLines | Where-Object {
    $_ -match '^orgUnitPath,name,parentOrgUnitPath$' -or $_ -match '^/'
  }
}

$all = if ($DryRun) { @() } else {
  ($csvLines -join "`n") | ConvertFrom-Csv
}

$children = if ($DryRun) { @() } else {
  $all | Where-Object { $_.parentOrgUnitPath -eq $BaseOUPath } |
        Select-Object orgUnitPath,name,parentOrgUnitPath
}

# Always write CSV (empty on dry-run is fine)
$children | Export-Csv -Path $OutCsv -NoTypeInformation
if ($DryRun) {
  Write-Host "Found (dry-run, not computed) immediate children -> wrote placeholder CSV: $(Resolve-Path $OutCsv)"
} else {
  Write-Host ("Found {0} immediate children. Wrote list to: {1}" -f ($children.Count), (Resolve-Path $OutCsv))
  if (-not $children -or $children.Count -eq 0) {
    Write-Host "No immediate children under $BaseOUPath"
    return
  }
}

# ============================
# 2) For each child OU, move the user (email == child name) into that OU
# ============================
Write-Host "`n[2/3] Moving users into their matching child OUs..."

$index = 0
foreach ($row in $children) {
  $index++
  $childPath = $row.orgUnitPath     # target OU to place the user into
  $email     = $row.name            # child OU name is the email

  Write-Host ("[{0}/{1}] OU: {2}" -f $index,$children.Count,$childPath)

  if ([string]::IsNullOrWhiteSpace($email) -or -not (Looks-LikeEmail $email)) {
    Write-Warning ("           Skipping: child name is not an email -> '{0}'" -f $row.name)
    continue
  }

  Write-Host ("           Move user -> {0}" -f $email)

  # GAM: move user to OU
  # Syntax: gam update user <email> org "<orgUnitPath>"
  $argsMove = @('update','user', $email, 'org', $childPath)

  if ($DryRun) {
    Write-Host "[DRYRUN] $GamPath $($argsMove -join ' ')"
    continue
  }

  $null = Run-GamArray -ArgsArray $argsMove -Quiet
  if ($LASTEXITCODE -ne 0) {
    Write-Warning "  -> GAM returned non-zero exit code (exit $LASTEXITCODE). Check if the user exists and your OU path is correct."
  } else {
    Write-Host "  -> Moved."
  }
}

# ============================
# 3) done
# ============================
Write-Host "`n[3/3] Complete."
if ($DryRun) { Write-Host "No changes were made (dry-run)." }
