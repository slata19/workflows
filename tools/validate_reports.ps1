param(
    [string]$WorkbookPath
)

Write-Host "Validating reports in $WorkbookPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($WorkbookPath)

# Run VBA function ValidateReports (returns Boolean)
$result = $excel.Run("ValidateReports")

$wb.Close($true)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

if (-not $result) {
    Write-Error "Report validation failed."
    exit 1
}

Write-Host "Report validation passed."
exit 0
