## Dry-run (preview only)
#   .\Apply-ChildOU-LoginRestriction.ps1 `
#  -BaseOUPath "\Student 1:1 Devices\00 - WLS\GR05" `
#  -GamPath "C:\GAM7\gam.exe" `
#  -DryRun

#  Execute for real
#   .\Apply-ChildOU-LoginRestriction.ps1 `
#  -BaseOUPath "\Student 1:1 Devices\00 - WLS\GR05" `
#  -GamPath "C:\GAM7\gam.exe"
#
#
#

param(
  [Parameter(Mandatory=$true)]
  [string]$BaseOUPathRaw,                          # e.g. "\Student 1:1 Devices\00 - WLS\GR05"

  [string]$GamPath = "gam",                        # e.g. "C:\GAM7\gam.exe"
  [string]$OutCsv  = ".\child_ous.csv",            # output CSV of immediate children
  [switch]$DryRun                                  # preview only
)

# ---- helpers -----------------------------------------------------------------
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

# ---- banner ------------------------------------------------------------------
$BaseOUPath = Normalize-OU $BaseOUPathRaw
Write-Host "Base OU: $BaseOUPath"
Write-Host "GAM   : $GamPath"
Write-Host "CSV   : $OutCsv"
if ($DryRun) { Write-Host "Mode  : DRY-RUN (no changes will be made)" }

# ---- 1) export ALL OUs, then keep only immediate children --------------------
Write-Host "`n[1/3] Exporting all OUs..."
$argsPrintOrgs = @('print','orgs','fields','name,orgUnitPath,parentOrgUnitPath')
$rawLines = Run-GamArray -ArgsArray $argsPrintOrgs -Quiet

if (-not $DryRun) {
  if ($LASTEXITCODE -ne 0) { throw "gam print orgs failed (exit $LASTEXITCODE)" }
  if (-not $rawLines -or $rawLines.Count -eq 0) { throw "gam print orgs returned no output" }
}

# Keep only CSV header + rows; drop progress chatter like
# "Getting all Organizational Units..." and "Got N Organizational Units"
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

# Always write a CSV for traceability (empty on dry-run is fine)
$children | Export-Csv -Path $OutCsv -NoTypeInformation
if ($DryRun) {
  Write-Host "Found (dry-run, not computed) immediate children -> wrote placeholder CSV: $(Resolve-Path $OutCsv)"
} else {
  Write-Host ("Found {0} immediate children. Wrote list to: {1}" -f ($children.Count), (Resolve-Path $OutCsv))
  if (-not $children -or $children.Count -eq 0) { Write-Host "No immediate children under $BaseOUPath"; return }
}

# ---- 2) apply sign-in restriction per child OU --------------------------------
Write-Host "`n[2/3] Applying Sign-In Restriction on each child OU..."
# Policy: chrome.devices.SignInRestriction
# Fields:
#   deviceAllowNewUsers = RESTRICTED_LIST
#   userAllowlist       = <child name/email>

$index = 0
foreach ($row in $children) {
  $index++
  $childPath = $row.orgUnitPath
  $email     = $row.name  # your tree uses the child OU name as the email

  if ([string]::IsNullOrWhiteSpace($email)) {
    Write-Warning ("[{0}/{1}] Skipping {2} (empty name/email)" -f $index,$children.Count,$childPath)
    continue
  }

  Write-Host ("[{0}/{1}] OU: {2}" -f $index,$children.Count,$childPath)
  Write-Host ("           Replace allowlist with: {0}" -f $email)

  # Build GAM args as an array (no quoting issues)
  $argsUpdate = @(
    'update','chromepolicy','orgunit', $childPath,
    'name','chrome.devices.SignInRestriction',
    'fields',
      'deviceAllowNewUsers','RESTRICTED_LIST',
      'userAllowlist',      $email
  )

  if ($DryRun) {
    Write-Host "[DRYRUN] $GamPath $($argsUpdate -join ' ')"
    continue
  }

  $null = Run-GamArray -ArgsArray $argsUpdate -Quiet
  if ($LASTEXITCODE -ne 0) {
    Write-Warning "  -> GAM returned non-zero exit code on $childPath (exit $LASTEXITCODE)"
  } else {
    Write-Host "  -> Applied."
  }
}

# ---- 3) done ------------------------------------------------------------------
Write-Host "`n[3/3] Complete."
if ($DryRun) { Write-Host "No changes were made (dry-run)." }

