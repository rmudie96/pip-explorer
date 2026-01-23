$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open("C:\Users\mudie\OneDrive - Local Trust\2026\PiP Explorer\lsoa_msoa_la_region_lookup.xlsx")
$sheet = $wb.Sheets.Item(1)
$range = $sheet.UsedRange

Write-Host "Rows: $($range.Rows.Count), Columns: $($range.Columns.Count)"
Write-Host ""
Write-Host "Columns:"
for ($col = 1; $col -le $range.Columns.Count; $col++) {
    Write-Host "  $col : $($sheet.Cells.Item(1, $col).Text)"
}
Write-Host ""
Write-Host "Sample row 2:"
for ($col = 1; $col -le $range.Columns.Count; $col++) {
    Write-Host "  $($sheet.Cells.Item(1, $col).Text): $($sheet.Cells.Item(2, $col).Text)"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
