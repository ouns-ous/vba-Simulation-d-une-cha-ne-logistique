from pathlib import Path
from win32com.client import dynamic
root = Path(__file__).resolve().parent.parent
xl = dynamic.Dispatch('Excel.Application')
xl.Visible = False
xl.EnableEvents = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(root / 'DataCo_Logistique.xlsm'))
    for name in ('F_Main', 'modFormLayout', 'modInterface', 'modData'):
        cm = wb.VBProject.VBComponents(name).CodeModule
        (root / '.codex-work' / f'design-before-{name}.vba').write_text(cm.Lines(1, cm.CountOfLines), encoding='utf-8')
    helper = wb.VBProject.VBComponents.Add(1)
    helper.CodeModule.AddFromString('''Public Function InspectDesign() As String
 Dim p As Object, c As Object, s As String
 s = "Layout error: " & F_Main.DisplayError & vbCrLf
 s = s & ReadProps(F_Main) & ReadProps(F_Main.Controls("tabsMain")) & ReadProps(F_Main.Controls("tabsMain").Pages(0))
 For Each p In F_Main.Controls("tabsMain").Pages
 s = s & p.Name & " " & p.Caption & " " & p.InsideWidth & "x" & p.InsideHeight & vbCrLf
 Next
 For Each c In F_Main.Controls("tabsMain").Pages(0).Controls
 s = s & c.Name & " " & TypeName(c) & vbCrLf
 Next
 InspectDesign = s
 Unload F_Main
End Function
Private Function ReadProps(ByVal host As Object) As String
 On Error Resume Next
 Dim v As Variant, prop As Variant, s As String
 For Each prop In Array("BackColor", "ScrollBars", "KeepScrollBarsVisible", "MultiRow", "Style")
 Err.Clear
 v = CallByName(host, CStr(prop), VbGet)
 s = s & TypeName(host) & "." & CStr(prop) & " error=" & Err.Number & vbCrLf
 Next
 ReadProps = s
End Function''')
    print(xl.Run("'DataCo_Logistique.xlsm'!InspectDesign"))
    wb.Close(False)
finally:
    xl.Quit()
