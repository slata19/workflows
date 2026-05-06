param(
    [string]$WorkbookPath,
    [string]$SourceFolder
)

Write-Host "Importing VBA from $SourceFolder into $WorkbookPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($WorkbookPath)
$vbProj = $wb.VBProject

# Remove existing non-document modules
for ($i = $vbProj.VBComponents.Count; $i -ge 1; $i--) {
    $comp = $vbProj.VBComponents.Item($i)
    if ($comp.Type -ne 100) { # 100 = document modules (sheets, ThisWorkbook)
        $vbProj.VBComponents.Remove($comp)
    }
}

# Import .bas, .cls, .frm
Get-ChildItem -Path $SourceFolder -Recurse -Include *.bas, *.cls, *.frm |
    ForEach-Object {
        Write-Host "Importing $($_.FullName)"
        $vbProj.VBComponents.Import($_.FullName) | Out-Null
    }

$wb.Save()
$wb.Close($true)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
