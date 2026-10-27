"""Theme the project actually launched by the user's VBA workbook.

Keep the open report untouched; point the launch project at a separate themed
report so saving the older open window cannot overwrite the themed definition.
The sibling semantic model, exports, workbook and VBA are retained unchanged.
"""
from pathlib import Path
import copy
import hashlib
import json
import shutil
import uuid
import zipfile
from datetime import datetime

SOURCE = Path(__file__).resolve().parent
TARGET = Path(r'C:\Users\HP\Desktop\DataCo_Projet-9dima\PowerBI')
OLD = TARGET / 'DataCo.Report'
NEW = TARGET / 'DataCo.BleuBlanc.Report'
def read(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))
def write(p, v):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(v, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
def lit(v):
    return {'expr': {'Literal': {'Value': str(v) + 'D' if isinstance(v, (int, float)) else "'" + v.replace("'", "''") + "'"}}}
def col(c):
    return {'solid': {'color': lit(c)}}
def position(v, x, y, w, h, z=10):
    v['position'] = dict(x=x, y=y, width=w, height=h, z=z, tabOrder=int(y * 2 + x / 100))
def text(v, lines, size=14):
    v['visual']['objects']['general'][0]['properties']['paragraphs'] = [
        {'textRuns': [{'value': s, 'textStyle': {'fontFamily': 'Segoe UI Semibold' if i == 0 else 'Segoe UI', 'fontSize': f'{size if i == 0 else size - 2}px', 'color': '#123B69' if i == 0 else '#59748F'}}]} for i, s in enumerate(lines)]
def hashes(folder):
    return {str(p.relative_to(folder)): hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.rglob('*') if p.is_file() and '.pbi' not in p.relative_to(folder).parts}

assert not NEW.exists(), 'The themed report already exists; inspect before replacing it.'
before_model = hashes(TARGET / 'DataCo.SemanticModel')
before_exports = hashes(TARGET.parent / 'exports')
stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
backup = TARGET / 'sauvegardes' / f'Projet_avant_theme_VBA_{stamp}.zip'
backup.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(backup, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(TARGET / 'DataCo.pbip', 'DataCo.pbip')
    for p in OLD.rglob('*'):
        if p.is_file() and '.pbi' not in p.relative_to(OLD).parts:
            z.write(p, p.relative_to(TARGET))

shutil.copytree(SOURCE / 'DataCo.Report', NEW, ignore=shutil.ignore_patterns('.pbi'))
pages = NEW / 'definition/pages'
oldpages = OLD / 'definition/pages'
order = read(oldpages / 'pages.json')['pageOrder']
meta = read(pages / 'pages.json')
meta['pageOrder'] = order
write(pages / 'pages.json', meta)
platform = read(NEW / '.platform')
platform['metadata']['displayName'] = 'DataCo · Bleu et blanc'
platform['config']['logicalId'] = str(uuid.uuid4())
write(NEW / '.platform', platform)

# Reuse the proven light formatting, keeping every old analytical query.
templates = {}
for p in (pages / '01_ensemble/visuals').glob('*/visual.json'):
    v = read(p)
    templates.setdefault(v['visual']['visualType'], v)
core = {p.parent.name: read(p) for p in (pages / '01_ensemble/visuals').glob('*/visual.json')}
nav_template = copy.deepcopy(core['nav_01_ensemble'])
nav_labels = ['Vue globale', 'Stock', 'Scénarios', 'Coûts', 'Fournisseurs']
queries_preserved = 0
for page_id in order:
    original = {p.parent.name: read(p) for p in (oldpages / page_id / 'visuals').glob('*/visual.json')}
    if not (pages / page_id).exists():
        write(pages / page_id / 'page.json', read(pages / '01_ensemble/page.json'))
        m = read(pages / page_id / 'page.json')
        m.update(name=page_id, displayName='Fournisseurs')
        write(pages / page_id / 'page.json', m)
        visuals = {n: copy.deepcopy(v) for n, v in core.items() if n.startswith('nav_') or n in ('navigation', 'control_panel', 'control_hint')}
        for name, ov in original.items():
            if not ov['visual'].get('query'):
                continue
            typ = ov['visual']['visualType']
            nv = copy.deepcopy(templates[typ])
            nv['name'] = name
            nv['visual']['query'] = copy.deepcopy(ov['visual']['query'])
            if 'filterConfig' in ov:
                nv['filterConfig'] = copy.deepcopy(ov['filterConfig'])
            title = ov['visual'].get('visualContainerObjects', {}).get('title')
            if title:
                nv['visual']['visualContainerObjects']['title'][0]['properties']['text'] = copy.deepcopy(title[0]['properties']['text'])
            visuals[name] = nv
        heading = copy.deepcopy(core['88efb87d8ca142eea0b4'])
        heading['name'] = 'titre_fournisseurs'
        text(heading, ['DataCo Global  |  Fournisseurs', 'Délais observés, retards et qualité des livraisons'], 24)
        visuals[heading['name']] = heading
        for i in range(4):
            position(visuals[f'transaction_{i}'], 20 + i * 354, 88, 338, 128)
        for name, box in {'delais': (20, 236, 526, 268), 'retard': (562, 236, 524, 268), 'table': (20, 520, 1066, 278), 'fournisseur': (1118, 430, 286, 180)}.items():
            position(visuals[name], *box)
    else:
        visuals = {p.parent.name: read(p) for p in (pages / page_id / 'visuals').glob('*/visual.json')}
        for name, v in visuals.items():
            if name in original and original[name]['visual'].get('query'):
                v['visual']['query'] = copy.deepcopy(original[name]['visual']['query'])
                for key in ('filterConfig',):
                    if key in original[name]:
                        v[key] = copy.deepcopy(original[name][key])
                    else:
                        v.pop(key, None)
                queries_preserved += 1
        # This older model has no Simulation_Contexte table. Avoid broken fields.
        ctx = copy.deepcopy(core['control_hint'])
        ctx['name'] = 'product_context'
        position(ctx, 1118, 298, 286, 112)
        text(ctx, ['Source des données', 'Exports du classeur Excel', 'DataCo_Projet-9dima'], 15)
        visuals['product_context'] = ctx
    for i, dest in enumerate(order):
        n = 'nav_' + dest
        v = copy.deepcopy(nav_template)
        v['name'] = n
        position(v, 290 + i * 224, 840, 208, 48, 30)
        objects = v['visual']['objects']
        for entry in objects['text']:
            if entry.get('selector', {}).get('id') == 'default':
                entry['properties'].update(text=lit(nav_labels[i]), fontColor=col('#FFFFFF' if dest == page_id else '#123B69'))
        for entry in objects['fill']:
            state = entry.get('selector', {}).get('id')
            if state:
                entry['properties']['fillColor'] = col('#238AD1' if state == 'hover' or dest == page_id else '#EAF3FB')
        link = v['visual']['visualContainerObjects']['visualLink'][0]['properties']
        link['navigationSection'] = lit(dest)
        link['tooltip'] = lit('Ouvrir ' + nav_labels[i] + ' (Ctrl + clic dans Desktop)')
        visuals[n] = v
    position(visuals['nav_brand'], 36, 838, 238, 56, 20)
    if page_id == '05_fournisseurs':
        text(visuals['control_hint'], ['Lecture du rapport', 'Choisir un fournisseur pour', 'filtrer les données historiques.', '', 'Les données sont actualisées', 'depuis le classeur Excel.'], 14)
    for name, v in visuals.items():
        write(pages / page_id / 'visuals' / name / 'visual.json', v)

# Route the unchanged VBA entry point to the themed report.
project = read(TARGET / 'DataCo.pbip')
project['artifacts'] = [{'report': {'path': NEW.name}}]
write(TARGET / 'DataCo.pbip', project)
write(TARGET / 'DataCo_BleuBlanc.pbip', project)
shutil.copy2(SOURCE / 'DataCo_Theme_BleuBlanc.json', TARGET / 'DataCo_Theme_BleuBlanc.json')
assert hashes(TARGET / 'DataCo.SemanticModel') == before_model, 'Semantic model changed unexpectedly'
assert hashes(TARGET.parent / 'exports') == before_exports, 'Source exports changed unexpectedly'
result = {'vba_project': str(TARGET / 'DataCo.pbip'), 'fresh_open_project': str(TARGET / 'DataCo_BleuBlanc.pbip'), 'report': str(NEW), 'pages': order, 'backup': str(backup), 'model_unchanged': True, 'exports_unchanged': True, 'original_queries_reused': queries_preserved, 'desktop_preview': 'Not verified: Computer Use kernel unavailable.'}
write(TARGET / 'correction_lancement_vba.json', result)
print(json.dumps(result, ensure_ascii=False, indent=2))
