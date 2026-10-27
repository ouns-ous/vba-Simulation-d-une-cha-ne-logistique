$ErrorActionPreference = 'Stop'
$root = 'C:\Users\HP\Desktop\DataCo_Projet'
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $xl.Workbooks.Item('DataCo_Logistique.xlsm')
if ($wb.FullName -ne "$root\DataCo_Logistique.xlsm") { throw 'Unexpected workbook' }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backup = "$root\.codex-work\DataCo_avant_position_liste_$stamp.xlsm"
$wb.SaveCopyAs($backup)
$cm = $wb.VBProject.VBComponents.Item('F_Main').CodeModule
$before = $cm.Lines(1, $cm.CountOfLines)
$needle = '    With mProductSelector'
$position = '        .Move p.Controls("par15").Left, p.Controls("par15").Top, p.Controls("par15").Width, p.Controls("par15").Height'
if (($before.Split([string[]]@($needle), [StringSplitOptions]::None)).Count -ne 2) { throw 'Selector block is not unique' }
$line = 0
for ($i = 1; $i -le $cm.CountOfLines; $i++) {
    if ($cm.Lines($i, 1).TrimEnd() -eq $needle) { $line = $i; break }
}
if (-not $line) { throw 'Selector block missing' }
$cm.InsertLines($line + 1, $position)
$after = $cm.Lines(1, $cm.CountOfLines)
if ($after.Replace($position + "`r`n", '') -cne $before) { throw 'Unexpected code difference' }
$source = "$root\VBA_Source\F_Main.frm"
$sourceText = [IO.File]::ReadAllText($source)
if (-not $sourceText.Contains($needle)) { throw 'Source selector block missing' }
[IO.File]::WriteAllText($source, $sourceText.Replace($needle, $needle + "`r`n" + $position), [Text.UTF8Encoding]::new($false))
$wb.Save()
Write-Output "Saved: $($wb.FullName)"
Write-Output "Backup: $backup"
Write-Output 'Verified: F_Main differs by one layout line only.'
