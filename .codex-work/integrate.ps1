$ErrorActionPreference = 'Stop'
$root = 'C:\Users\HP\Desktop\DataCo_Projet'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Copy-Item -LiteralPath "$root\DataCo_Logistique.xlsm" -Destination "$PSScriptRoot\DataCo_avant_suggestions_$stamp.xlsm"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.EnableEvents = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open("$root\DataCo_Logistique.xlsm")
    foreach ($name in @('modData','modInterface','modSimulation','modFormLayout','F_Main','ThisWorkbook')) {
        $ext = if ($name -eq 'F_Main') { 'frm' } elseif ($name -eq 'ThisWorkbook') { 'vba' } else { 'bas' }
        $text = [IO.File]::ReadAllText("$root\VBA_Source\$name.$ext")
        if ($name -eq 'F_Main') { $text = $text.Substring($text.IndexOf('Option Explicit')) }
        $text = (($text -split '\r?\n') | Where-Object { $_ -notmatch '^Attribute ' }) -join "`r`n"
        $cm = $book.VBProject.VBComponents.Item($name).CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1,$cm.CountOfLines) }
        $cm.AddFromString($text)
        Write-Output "Integrated $name"
    }
    $book.Save()
    Write-Output 'Saved macros'
    $selected = [string]$excel.Run("'DataCo_Logistique.xlsm'!Param",'ProduitSimulation')
    $rows = @()
    foreach ($product in @('P001','P002','P003','P004')) {
        $excel.Run("'DataCo_Logistique.xlsm'!SetParam",'ProduitSimulation',$product)
        $excel.Run("'DataCo_Logistique.xlsm'!Calibrate",$true)
        $state = [ordered]@{ product=$product }
        foreach ($key in @('mu','sigma','L','Cpos','Ccmd','S0','s','Q')) {
            $state[$key] = $excel.Run("'DataCo_Logistique.xlsm'!Param",$key)
        }
        $state['policies'] = @()
        $sc = $book.Worksheets.Item('Scenarios').Range('A2:E4').Value2
        for ($r=1; $r -le 3; $r++) { $state['policies'] += ,@($sc[$r,1],$sc[$r,2],$sc[$r,3],$sc[$r,4],$sc[$r,5]) }
        $comparison = $excel.Run("'DataCo_Logistique.xlsm'!DiagnoseComparison")
        if ($comparison -ne 'PASS') { throw "$product comparison: $comparison" }
        $state['comparison'] = $comparison
        $stats = $book.Worksheets.Item('Export_Stats').ListObjects.Item('tblExport_Stats').DataBodyRange.Value2
        $state['results'] = @()
        for($r=1;$r -le 3;$r++) { $record=@(); for($c=1;$c -le $stats.GetLength(1);$c++){ $record += $stats[$r,$c] }; $state['results'] += ,$record }
        $rows += $state
        Write-Output ($state | ConvertTo-Json -Depth 5 -Compress)
    }
    $rows | ConvertTo-Json -Depth 6 | Set-Content "$PSScriptRoot\product-results.json" -Encoding UTF8
    $excel.Run("'DataCo_Logistique.xlsm'!SetParam",'ProduitSimulation',$selected)
    $excel.Run("'DataCo_Logistique.xlsm'!Calibrate",$true)
    $excel.Run("'DataCo_Logistique.xlsm'!CompareScenarios",$true)
    $excel.Run("'DataCo_Logistique.xlsm'!ExportAll",$true)
    $excel.CalculateFull()
    $book.Save()
    Write-Output 'Saved calibrated workbook and refreshed exports'
    $book.Close($false)
} finally { $excel.EnableEvents=$true; $excel.Quit() }
