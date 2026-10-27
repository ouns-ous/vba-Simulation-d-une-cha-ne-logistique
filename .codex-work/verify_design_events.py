from pathlib import Path
from win32com.client import dynamic
import json
root=Path(__file__).resolve().parent.parent
xl=dynamic.Dispatch('Excel.Application')
wb=xl.Workbooks('DataCo_Logistique.xlsm')
helper=wb.VBProject.VBComponents.Add(1)
helper.Name='TempVerifyDesignEvents'
helper.CodeModule.AddFromString('''Public Function VerifyDesignEvents() As String
    Load F_Main
    F_Main.Controls("designNav1").Value = True
    If F_Main.Controls("tabsMain").Value <> 1 Then Err.Raise 530, , "Navigation click event failed"
    If F_Main.Controls("tabsMain").Pages(1).Zoom <> 80 Then Err.Raise 531, , "Default page zoom failed"
    F_Main.Controls("designZoomIn").Value = True
    If F_Main.Controls("tabsMain").Pages(1).Zoom <> 90 Then Err.Raise 532, , "Zoom click event failed"
    F_Main.Controls("designZoomOut").Value = True
    F_Main.Controls("designNav0").Value = True
    If F_Main.Controls("tabsMain").Value <> 0 Then Err.Raise 533, , "Home click event failed"
    If F_Main.Controls("tabsMain").Pages(0).Zoom <> 100 Then Err.Raise 534, , "Home zoom failed"
    VerifyDesignEvents = "Navigation and zoom CommandButton events: OK"
End Function''')
try:
    result=xl.Run("'DataCo_Logistique.xlsm'!VerifyDesignEvents")
    print(result)
    path=root/'.codex-work'/'verification-new-design.json'
    report=json.loads(path.read_text(encoding='utf-8'))
    report['command_button_events']=result
    report['visual_capture']='Unavailable: Windows capture timed out; runtime geometry and accessibility inspected.'
    path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
finally:
    wb.VBProject.VBComponents.Remove(helper)
    wb.Save()
xl.Run("'DataCo_Logistique.xlsm'!OpenMain")
