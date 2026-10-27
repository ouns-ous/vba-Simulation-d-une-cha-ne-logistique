$ErrorActionPreference='Stop'
$root='C:\Users\HP\Desktop\DataCo_Projet'
$xl=New-Object -ComObject Excel.Application
$xl.Visible=$false
$xl.EnableEvents=$false
$xl.DisplayAlerts=$false
try {
 $wb=$xl.Workbooks.Open("$root\DataCo_Logistique.xlsm")
 foreach($name in @('F_Main','modFormLayout','modInterface','modData')) {
  $cm=$wb.VBProject.VBComponents.Item($name).CodeModule
  [IO.File]::WriteAllText("$root\.codex-work\design-before-$name.vba",$cm.Lines(1,$cm.CountOfLines))
 }
 $helper=$wb.VBProject.VBComponents.Add(1)
 $helper.CodeModule.AddFromString(@'
Public Function InspectDesign() As String
 Dim p As Object, c As Object, s As String
 s = "Layout error: " & F_Main.DisplayError & vbCrLf
 For Each p In F_Main.Controls("tabsMain").Pages
 s = s & p.Name & " " & p.Caption & " " & p.InsideWidth & "x" & p.InsideHeight & vbCrLf
 Next
 For Each c In F_Main.Controls("tabsMain").Pages(0).Controls
 s = s & c.Name & " " & TypeName(c) & vbCrLf
 Next
 InspectDesign = s
 Unload F_Main
End Function
'@)
 Write-Output $xl.Run("'DataCo_Logistique.xlsm'!InspectDesign")
 $wb.Close($false)
} finally { $xl.Quit() }
