
$xlsxFile = "C:\" #change this with your desired path


$folderPath = [System.IO.Path]::GetDirectoryName($xlsxFile)


$csvFile = [System.IO.Path]::Combine($folderPath, "ProjectsTemplate.csv")


if (Test-Path $csvFile) {
    Write-Host "The file 'ProjectsTemplate.csv' already exists. Deleting it..." -ForegroundColor Red
    Remove-Item $csvFile -Force
}


$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false 


$Workbook = $Excel.Workbooks.Open($xlsxFile)


$Workbook.SaveAs($csvFile, 6)


$Workbook.Close()
$Excel.Quit()


[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null


$csvContent = Get-Content $csvFile
$csvContent | Set-Content -Path $csvFile -Encoding UTF8

Write-Host "Conversion complete! The CSV file is saved at: $csvFile"
