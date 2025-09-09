# One-liner to get immediate children of a base OU into child_ous.csv
$base = '/Student 1:1 Devices/00 - WLS/GR05'
& 'C:\GAM7\gam.exe' print orgs fields name,orgUnitPath,parentOrgUnitPath 2>&1 |
  Where-Object { $_ -match '^orgUnitPath,name,parentOrgUnitPath$' -or $_ -match '^/' } |
  ConvertFrom-Csv |
  Where-Object { $_.parentOrgUnitPath -eq $base } |
  Select-Object orgUnitPath,name,parentOrgUnitPath |
  Export-Csv .\child_ous.csv -NoTypeInformation
