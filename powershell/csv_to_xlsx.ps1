<#
  Author: Peter Wood (peter@peterwood.dev)
  Date: 7/12/2022
  Description:
    PowerShell script to convert any CSVs found in directory ($dirSource) that were created today -
    (1) saves them as an XLSX,
    (2) renames the first tab,
    (3) moves the XLSX to the parent directory
#>

$dirSource = "\\peranda-nas\media\backups"
$filePrefix = "placeholder"

Get-ChildItem $dirSource -Filter $filePrefix*.csv | Where-Object {($_.LastAccessTime.Date -eq [DateTime]::Now.Date)} | ForEach-Object {
    $csv = $_.FullName
    Write-Host "filePath ="$csv

    $dirParent = (Get-Item $csv).Directory.Parent.FullName
    # Write-Host "parentDir ="$dirParent

    $excel = New-Object -ComObject "Excel.Application"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # open the csv
    $workbook = $excel.Workbooks.Open($csv)

    # rename the tab
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "TabName"

    # save & close the CSV as XLSX
    $xlsx = $csv.Replace(".csv", ".xlsx")
    $workbook.SaveAs($xlsx, 51)

    $excel.Quit()

    # move to parent directory
    Move-Item -Path $xlsx -Destination $dirParent -Force
    # Copy-Item -Path $xlsx -Destination $dirParent -Force
    # Remove-Item -Path $xlsx

}