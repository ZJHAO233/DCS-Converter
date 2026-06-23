[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$files = Get-ChildItem "E:\Temp\mimo\*.xlsx"
$xlsxPath = $files[0].FullName
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open($xlsxPath)
$ws = $wb.Sheets.Item(2)
$lastRow = $ws.UsedRange.Rows.Count
Write-Host "Sheet: $($ws.Name) Rows:$lastRow"

for ($r = 166; $r -le [Math]::Min($lastRow, 210); $r++) {
    $row = @()
    for ($c = 1; $c -le 7; $c++) {
        $row += $ws.Cells.Item($r, $c).Text
    }
    Write-Host "R$r`: $($row -join '|')"
}

$wb.Close($false)
$excel.Quit()
