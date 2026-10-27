$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.EnableEvents = $false
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open('C:\Users\HP\Desktop\DataCo_Projet\DataCo_Logistique.xlsm')
    foreach ($name in @('modData','modInterface','F_Main','modSimulation','ThisWorkbook')) {
        $cm = $book.VBProject.VBComponents.Item($name).CodeModule
        [IO.File]::WriteAllText("$PSScriptRoot/native-$name.vba", $cm.Lines(1,$cm.CountOfLines))
    }
    foreach ($name in @('Feuille_Controle','Scenarios')) {
        $sh = $book.Worksheets.Item($name)
        Write-Output $name
        $a = $sh.Range('A1:F22').Value2
        for ($r=1; $r -le 22; $r++) { $line = @(); for($c=1; $c -le 6; $c++) { $line += $a[$r,$c] }; Write-Output ($line -join ' | ') }
    }
    $book.Close($false)
} finally { $excel.Quit() }
