param (
  [switch]$DebugCommand = $false,
  [int]$NumDays = 90
)

$path = Join-Path $env:OneDrive -ChildPath "Documents" | Join-Path -ChildPath "rclone"
$numDays = $numDays * -1

$space_before = "{0:N2} GB" -f (((Get-ChildItem -Path $path -Recurse | Measure-Object -Property Length -Sum).Sum / 1GB), 2)

if ($debugCommand) {
  Write-Host $path
  Get-ChildItem -Path $path -Recurse -Exclude *.ps1 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays($numDays) } | Remove-Item -Recurse -ErrorAction SilentlyContinue -WhatIf

  exit
}

Write-Host "Space before"$space_before

Get-ChildItem -Path $path -Recurse  -Exclude *.ps1 | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays($numDays) } | Remove-Item -Recurse -ErrorAction SilentlyContinue
$space_after = "{0:N2} GB" -f (((Get-ChildItem -Path $path -Recurse | Measure-Object -Property Length -Sum).Sum / 1GB), 2)


Write-Host "Space after"$space_after
