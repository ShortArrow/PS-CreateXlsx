using namespace System.Collections.Generic

if ([int]$psversiontable.psversion.major -lt 6) {
    Write-Host "PowerShell version need 6 or later" -BackgroundColor Red -ForegroundColor White
}
else {
    Write-Host "PowerShell version is Fit" -BackgroundColor Green -ForegroundColor White
}

try {
    # [System.__ComObject]$excel = New-Object -ComObject Excel.Application
    $excel = New-Object -ComObject Excel.Application
    # $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    [__ComObject]$book = $excel.Workbooks.Add()
    [__ComObject]$sheet = $book.WorkSheets(1)
    $sheet.Name = "CookingSheet"
    [Microsoft.VisualBasic.VariantType]$Empty = [Microsoft.VisualBasic.VariantType]::Empty
    [Microsoft.VisualBasic.VariantType[]]$array1 = (1, 2, $Empty, 3, $Empty, $Empty, $Empty, 4, $Empty, 5, $Empty, 6, $Empty, $Empty, $Empty, 7, $Empty, 8, $Empty, 9, $Empty, $Empty, $Empty, 10, $Empty, 11)
    for ($i = 0; $i -lt 10; $i++) {
        $sheet.Range($sheet.Cells($i + 1, 1), $sheet.Cells($i + 1, $array1.Count)).Value(10) = $array1
    }
    $sheet.Columns.ColumnWidth = 3

    if (!(Test-Path("output"))) {
        New-Item -Path "output" -ItemType Directory
    }

    $book.SaveAs("$(Get-Location)\output\generated_from_posh_$(Get-Date -Format yyyyMMdd_HHmmss).xlsx")
    $excel.DisplayAlerts = $true
    $excel.ScreenUpdating = $true
    $excel.EnableEvents = $true
    $excel.Quit()
}
catch {
    Write-Host "Error" -ForegroundColor Red
}
finally {
    # $excel = $Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)  | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [GC]::collect()  
}

Write-Host "Finish" -ForegroundColor Green
# [Console]::ReadKey($true) | Out-Null