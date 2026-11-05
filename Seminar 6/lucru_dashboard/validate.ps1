param([string]$Workbook = 'S6_Dashborad_demo (2) - corectat.xlsx')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName WindowsBase
Add-Type -Path 'C:\Program Files\Microsoft Office\root\vfs\Windows\assembly\GAC_MSIL\DocumentFormat.OpenXml\2.8.1.0__8FB06CB64D019A17\DocumentFormat.OpenXml.dll'
$document = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open((Join-Path $pwd $Workbook), $false)
try {
    $validator = New-Object DocumentFormat.OpenXml.Validation.OpenXmlValidator([DocumentFormat.OpenXml.FileFormatVersions]::Office2016)
    $validator.MaxNumberOfErrors = 1000
    $validationErrors = @($validator.Validate($document))
    $sheetCount = @($document.WorkbookPart.WorksheetParts).Count
    $pivotCount = 0
    foreach ($worksheetPart in $document.WorkbookPart.WorksheetParts) {
        $pivotCount += @($worksheetPart.PivotTableParts).Count
    }
    $report = [ordered]@{
        workbook = $Workbook
        validator = 'Microsoft Open XML SDK 2.8.1 / Office 2016'
        errors = $validationErrors.Count
        worksheets = $sheetCount
        pivottables = $pivotCount
        details = @($validationErrors | ForEach-Object {
            [ordered]@{part = $_.Part.Uri.ToString(); path = $_.Path.XPath; description = $_.Description}
        })
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 'lucru_dashboard/validare_microsoft.json'
    $report | ConvertTo-Json -Depth 5
    if ($validationErrors.Count -gt 0) { throw 'Microsoft Open XML validation failed.' }
    if ($sheetCount -ne 15 -or $pivotCount -ne 11) { throw 'Unexpected workbook component count.' }
}
finally { $document.Dispose() }
