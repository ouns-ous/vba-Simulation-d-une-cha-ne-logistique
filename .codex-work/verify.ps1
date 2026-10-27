$ErrorActionPreference = 'Stop'
$root = 'C:\Users\HP\Desktop\DataCo_Projet'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.EnableEvents = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open("$root\DataCo_Logistique.xlsm")
    $component = $book.VBProject.VBComponents.Import("$PSScriptRoot\TestSuggestions.bas")
    $result = $excel.Run("'DataCo_Logistique.xlsm'!VerifyProductSuggestions")
    Write-Output $result
    $result | Set-Content "$PSScriptRoot\native-tests.txt" -Encoding UTF8
    $book.VBProject.VBComponents.Remove($component)
    if ($result -notlike 'PASS*') { throw $result }
    $excel.Run("'DataCo_Logistique.xlsm'!CompareScenarios",$true)
    $excel.Run("'DataCo_Logistique.xlsm'!ExportAll",$true)
    $excel.CalculateFull()
    foreach($name in @('Scenarios','Feuille_Controle')) {
        $sheet=$book.Worksheets.Item($name)
        $printArea = $sheet.PageSetup.PrintArea
        $sheet.PageSetup.PrintArea = if($name -eq 'Scenarios') { '$A$1:$E$4' } else { '$B$2:$D$21' }
        $sheet.ExportAsFixedFormat(0,"$PSScriptRoot\$name.pdf")
        $sheet.PageSetup.PrintArea = $printArea
    }
    $book.Save()
    $book.Close($false)
    Write-Output 'Saved and verified'
} finally { $excel.EnableEvents=$true; $excel.Quit() }
