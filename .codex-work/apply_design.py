from pathlib import Path
from datetime import datetime
from win32com.client import dynamic
import hashlib, json, re
root = Path(__file__).resolve().parent.parent
work = root / '.codex-work'
xl = dynamic.Dispatch('Excel.Application')
xl.Visible = False
xl.EnableEvents = False
xl.DisplayAlerts = False
def normalize(s):
    return s.replace('\r\r\n', '\n').replace('\r\n', '\n').strip()
def module_text(wb, name):
    cm = wb.VBProject.VBComponents(name).CodeModule
    return normalize(cm.Lines(1,cm.CountOfLines)) if cm.CountOfLines else ''
def data_digest(wb):
    values = {}
    for sh in wb.Worksheets:
        for lo in sh.ListObjects:
            values[lo.Name] = hashlib.sha256(repr(lo.Range.Value2).encode()).hexdigest()
    values['parameters'] = hashlib.sha256(repr(wb.Worksheets('Feuille_Controle').Range('C6:C21').Value2).encode()).hexdigest()
    return values
try:
    existing=[book for book in xl.Workbooks if book.FullName.lower()==str(root/'DataCo_Logistique.xlsm').lower()]
    wb = existing[0] if existing else xl.Workbooks.Open(str(root/'DataCo_Logistique.xlsm'))
    helper=wb.VBProject.VBComponents.Add(1)
    helper.CodeModule.AddFromString('Public Sub CloseDesignPreview()\nUnload F_Main\nEnd Sub')
    xl.Run("'DataCo_Logistique.xlsm'!CloseDesignPreview")
    wb.VBProject.VBComponents.Remove(helper)
    xl.WindowState=-4137
    backup = work / ('DataCo_avant_nouveau_design_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsm')
    wb.SaveCopyAs(str(backup))
    original = {c.Name: module_text(wb,c.Name) for c in wb.VBProject.VBComponents}
    before_data = data_digest(wb)
    # The domain procedures are compared independently of presentation changes.
    changed_ui = {'UserForm_Initialize','ArrangeInterface','RefreshPage','RefreshHome','UserForm_QueryClose','InstallDesignControls','HandleDesignAction'}
    def procedures(code):
        return {m.group(2): m.group(0).strip() for m in re.finditer(r'(?m)^(Private|Public) (?:Sub|Function|Property Get) (\w+)[\s\S]*?^End (?:Sub|Function|Property)',code)}
    for name, filename in [('modFormLayout','modFormLayout.bas'),('F_Main','F_Main.frm'),('CDataCoDesignButton','CDataCoDesignButton.cls')]:
        source = (root/'VBA_Source'/filename).read_text(encoding='utf-8')
        source = source[source.index('Option Explicit'):]
        source = '\n'.join(line for line in source.splitlines() if not line.startswith('Attribute '))
        if name=='F_Main':
            before_procs=procedures(original[name]); after_procs=procedures(source)
            for procname,body in before_procs.items():
                if procname not in changed_ui:
                    assert after_procs[procname] == body, f'Domain procedure changed: {procname}'
        try:
            comp=wb.VBProject.VBComponents(name)
        except Exception:
            comp=wb.VBProject.VBComponents.Add(2)
            comp.Name=name
        cm=comp.CodeModule
        if cm.CountOfLines: cm.DeleteLines(1,cm.CountOfLines)
        cm.AddFromString(source)
    helper=wb.VBProject.VBComponents.Add(1)
    helper.Name='TempVerifyNewDesign'
    helper.CodeModule.AddFromString((work/'VerifyNewDesign.bas').read_text())
    try:
        result=xl.Run("'DataCo_Logistique.xlsm'!VerifyNewDesign")
        print(result,flush=True)
        after_data=data_digest(wb)
        assert after_data==before_data, 'Workbook data changed during design verification'
        for name,text in original.items():
            if name not in ('F_Main','modFormLayout'):
                assert module_text(wb,name)==text, f'Unrelated VBA module changed: {name}'
        report={'result':result,'backup':str(backup),'tables_and_parameters_unchanged':True,'domain_procedures_unchanged':True,'unrelated_modules_unchanged':True}
        (work/'verification-new-design.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    finally:
        wb.VBProject.VBComponents.Remove(helper)
    wb.Save()
    # Export the actual saved designer so re-importing the source retains user edits.
    export=work/'new-design-F_Main.frm'
    if export.exists(): export.unlink()
    wb.VBProject.VBComponents('F_Main').Export(str(export))
    frm=export.read_text(encoding='cp1252').replace('new-design-F_Main.frx','F_Main.frx')
    (root/'VBA_Source'/'F_Main.frm').write_text(frm,encoding='utf-8')
    (root/'VBA_Source'/'F_Main.frx').write_bytes(export.with_suffix('.frx').read_bytes())
    print('Saved design. All tables, parameters and domain procedures unchanged.',flush=True)
    wb.Close(False)
finally:
    xl.Quit()
