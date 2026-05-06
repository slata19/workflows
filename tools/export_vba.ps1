param(
    [string]$WorkbookPath,
    [string]$OutputFolder
)

Write-Host "Exporting VBA from $WorkbookPath to $OutputFolder"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($WorkbookPath)
$vbProj = $wb.VBProject

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

foreach ($comp in $vbProj.VBComponents) {
    $name = $comp.Name
    $ext = $null

    switch ($comp.Type) {
        1 { $ext = ".bas" } # standard module
        2 { $ext = ".cls" } # class module
        3 { $ext = ".frm" } # userform
        default { continue }
    }

    $path = Join-Path $OutputFolder "$name$ext"
    Write-Host "Exporting $path"
    $comp.Export($path)
}

$wb.Close($true)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
