$Path = "C:\Work\"
$newPath = 'D:\'
$report = @()

$excelSheets = Get-Childitem -Path $path -Include *.xls,*.xlsx -Recurse
$excel = New-Object -comobject Excel.Application
$excel.visible = $false
foreach($excelSheet in $excelSheets)
{
    $FileUpdated = $false
    $workbook = $Excel.Workbooks.Open($excelSheet) 
    foreach ($LinkPath in $workbook.LinkSources())
    {
        if($LinkPath -like "$Path*")
        {
            $newLink = $LinkPath -replace [regex]::Escape($path), [regex]::Escape($newPath)
            $workbook.ChangeLink($LinkPath, $newLink) 
            $FileUpdated = $true

            $details = @{            
                Excelfile = $excelSheet.FullName              
                OldLink = $LinkPath
                NewLink = $newLink
            }
            $report += New-Object PSObject -Property $details 
        }
    }
    if ($FileUpdated)
    {
        $workbook.save()
    }
    $workbook.close()
}
$report | export-csv -Path $Path\so.csv -NoTypeInformation

$excel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
$excel = $null

