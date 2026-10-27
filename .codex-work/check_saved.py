import zipfile, xml.etree.ElementTree as E, hashlib, json
from pathlib import Path
root=Path(__file__).resolve().parent.parent
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def read(path):
    out={}
    with zipfile.ZipFile(path) as z:
        shared=E.fromstring(z.read('xl/sharedStrings.xml'))
        strings=[''.join(s.itertext()) for s in shared]
        book=E.fromstring(z.read('xl/workbook.xml'))
        rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={r.attrib['Id']:r.attrib['Target'] for r in rels}
        for s in book.find('s:sheets',ns):
            rid=s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target=relmap[rid]
            p=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            cells={}; digest=hashlib.sha256(); errors=[]
            for event,c in E.iterparse(z.open(p),events=('end',)):
                if c.tag.endswith('}c'):
                    v=c.find('s:v',ns); f=c.find('s:f',ns)
                    val=v.text if v is not None else None
                    if c.attrib.get('t')=='s': val=strings[int(val)]
                    item=(c.attrib.get('t','n'),val,f.text if f is not None else None)
                    digest.update(json.dumps((c.attrib['r'],item),ensure_ascii=False).encode())
                    if c.attrib['r'] in ['C21','C6','C2','C3','C4','E1']: cells[c.attrib['r']]=item
                    if item[0]=='e': errors.append(c.attrib['r'])
                    c.clear()
                elif c.tag.endswith('}row'): c.clear()
            out[s.attrib['name']]={'cells':cells,'hash':digest.hexdigest(),'errors':errors}
    return out
before=read(next((root/'.codex-work').glob('DataCo_avant_suggestions_*.xlsm')))
after=read(root/'DataCo_Logistique.xlsm')
assert before.keys()==after.keys()
for name in ['T_Produits','T_Fournisseurs','T_Commandes','T_Stock','T_Approvisionnements','T_Mouvements']:
    assert before[name]==after[name],name
assert after['Feuille_Controle']['cells']['C21'][1]=='P002'
assert float(after['Feuille_Controle']['cells']['C6'][1])==415
assert all(after['Scenarios']['cells'][f'C{r}'][1]=='415' for r in [2,3,4])
assert after['Export_Decision']['cells']['E1'][1]=='ACTUEL'
errors=[(n,c) for n,cs in after.items() for c in cs['errors']]
assert not errors, errors[:10]
sha=hashlib.sha256((root/'DataCo_Logistique.xlsm').read_bytes()).hexdigest()
verification=json.loads((root/'verification.json').read_text(encoding='utf-8-sig'))
verification['SuggestionsProduits20261022']={
    'NativeExcel':(root/'.codex-work/native-tests.txt').read_text(encoding='utf-8-sig').strip(),
    'Independent':'PASS: 4 historiques, 12 couples s/Q, 4 S0, 360 repetitions',
    'SavedWorkbook':'PASS: P002, S0=415, ACTUEL, aucune erreur de cellule',
    'Preservation':'PASS: 19 feuilles, six tables transactionnelles identiques',
    'PowerBI':'Exports CSV actualises pour P002. Actualisation native Power BI non executee.',
    'WorkbookSha256':sha}
verification['FinalWorkbookSha256']=sha
(root/'verification.json').write_text(json.dumps(verification,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print('PASS: fichier sauvegarde, 19 feuilles, tables de gestion intactes, P002 S0=415, resultats ACTUEL, aucune erreur de cellule. SHA256 '+sha)
