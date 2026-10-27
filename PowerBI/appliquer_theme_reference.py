"""Apply the supplied blue/white reference to the existing native PBIR report."""
from pathlib import Path
import copy
import json
import zipfile
from datetime import datetime

ROOT = Path(__file__).resolve().parent
REPORT = ROOT / 'DataCo.Report'
PAGES = REPORT / 'definition/pages'
NAVY, BLUE, GREEN, ORANGE = '#123B69', '#238AD1', '#2AA775', '#F3A62C'
BG, LINE, MUTED = '#F1F7FC', '#D7E6F2', '#59748F'
SCHEMA = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.12.0/schema.json'

def read(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))

def write(p, obj):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')

def lit(v):
    if isinstance(v, bool):
        s = str(v).lower()
    elif isinstance(v, (int, float)):
        s = str(v) + 'D'
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {'expr': {'Literal': {'Value': s}}}

def color(c):
    return {'solid': {'color': lit(c)}}

def obj(properties):
    return [{'properties': properties}]

def place(v, x, y, w, h, z=10):
    v['position'] = dict(x=x, y=y, width=w, height=h, z=z, tabOrder=int(y * 2 + x / 100))

def style(v, title=None, background=True, border=True):
    co = v['visual'].setdefault('visualContainerObjects', {})
    co['background'] = obj(dict(show=lit(background), color=color('#FFFFFF'), transparency=lit(0)))
    co['border'] = obj(dict(show=lit(border), color=color(LINE), radius=lit(8)))
    co['dropShadow'] = obj(dict(show=lit(False)))
    co['padding'] = obj({k: lit(12) for k in ('top', 'bottom', 'left', 'right')})
    if title is not None:
        co['title'] = obj(dict(show=lit(True), text=lit(title), fontColor=color(NAVY), fontFamily=lit('Segoe UI Semibold'), fontSize=lit(12)))
    elif 'title' in co:
        co['title'][0]['properties'].update(fontColor=color(NAVY), fontSize=lit(12))
    typ = v['visual']['visualType']
    vo = v['visual'].setdefault('objects', {})
    if typ == 'card':
        vo['labels'] = obj(dict(color=color(NAVY), fontSize=lit(28), fontFamily=lit('Segoe UI Semibold'), labelDisplayUnits=lit(1), labelPrecision=lit(1)))
        vo['categoryLabels'] = obj(dict(show=lit(False)))
        vo['wordWrap'] = obj(dict(show=lit(True)))
    elif typ in ('clusteredColumnChart', 'lineChart', 'scatterChart'):
        vo['legend'] = obj(dict(show=lit(True), position=lit('Bottom'), showTitle=lit(False), fontSize=lit(10), labelColor=color(MUTED)))
        vo['categoryAxis'] = obj(dict(labelColor=color(MUTED), fontSize=lit(10), showAxisTitle=lit(False)))
        vo['valueAxis'] = obj(dict(labelColor=color(MUTED), fontSize=lit(10), showAxisTitle=lit(False), gridlineColor=color('#E9F0F6')))
        if typ == 'clusteredColumnChart':
            vo['labels'] = obj(dict(show=lit(True), color=color(NAVY), fontSize=lit(10)))
        if typ == 'lineChart':
            vo['lineStyles'] = obj(dict(strokeWidth=lit(2)))
    elif typ == 'tableEx':
        vo['columnHeaders'] = obj(dict(backColor=color('#EAF3FB'), fontColor=color(NAVY), fontFamily=lit('Segoe UI Semibold'), fontSize=lit(10), wordWrap=lit(True)))
        vo['values'] = obj(dict(backColorPrimary=color('#FFFFFF'), backColorSecondary=color('#F5F9FD'), fontColorPrimary=color(NAVY), fontColorSecondary=color(NAVY), fontSize=lit(10), wordWrap=lit(True)))
        vo['grid'] = obj(dict(gridVertical=lit(False), gridHorizontal=lit(True), gridHorizontalColor=color('#EAF1F7'), rowPadding=lit(12), outlineColor=color(LINE)))
        vo['total'] = obj(dict(totals=lit(False)))
    return v

def textbox(name, lines, x, y, w, h, bg=False, edge=False, size=14, z=20):
    v = {'$schema': SCHEMA, 'name': name, 'visual': {'visualType': 'textbox'}}
    place(v, x, y, w, h, z)
    style(v, background=bg, border=edge)
    v['visual']['visualContainerObjects']['padding'] = obj(dict(top=lit(4), bottom=lit(4), left=lit(12), right=lit(12)))
    v['visual']['objects']['general'] = obj({'paragraphs': [
        {'textRuns': [{'value': text, 'textStyle': {'fontFamily': 'Segoe UI Semibold' if i == 0 else 'Segoe UI', 'fontSize': f'{size if i == 0 else size - 2}px', 'color': NAVY if i == 0 else MUTED}}]}
        for i, text in enumerate(lines)
    ]})
    return v

def projection(prop, label=None, table='Scenarios', measure=True):
    return {'field': {'Measure' if measure else 'Column': {'Expression': {'SourceRef': {'Entity': table}}, 'Property': prop}}, 'queryRef': table + '.' + prop, 'nativeQueryRef': prop, **({'displayName': label} if label else {})}

def card(name, measure, title, x, y, w=338, h=128, table='Scenarios'):
    v = {'$schema': SCHEMA, 'name': name, 'visual': {'visualType': 'card', 'query': {'queryState': {'Values': {'projections': [projection(measure, table=table)]}}}, 'drillFilterOtherVisuals': True}}
    place(v, x, y, w, h)
    return style(v, title)

def save_visual(page, v):
    write(page / 'visuals' / v['name'] / 'visual.json', v)

stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
backup = ROOT / 'sauvegardes' / f'DataCo_avant_theme_{stamp}.zip'
backup.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(backup, 'w', zipfile.ZIP_DEFLATED) as z:
    for p in REPORT.rglob('*'):
        if p.is_file() and '.pbi' not in p.relative_to(REPORT).parts:
            z.write(p, p.relative_to(ROOT))

theme = {
    'name': 'DataCo · Bleu et blanc',
    'dataColors': [BLUE, GREEN, ORANGE, '#8554BD', '#1B5CAA', '#8BC9F4'],
    'good': GREEN, 'neutral': ORANGE, 'bad': '#D95762',
    'background': '#FFFFFF', 'foreground': NAVY, 'tableAccent': BLUE,
    'textClasses': {
        'callout': {'fontSize': 28, 'fontFace': 'Segoe UI Semibold', 'color': NAVY},
        'title': {'fontSize': 12, 'fontFace': 'Segoe UI Semibold', 'color': NAVY},
        'header': {'fontSize': 11, 'fontFace': 'Segoe UI Semibold', 'color': NAVY},
        'label': {'fontSize': 10, 'fontFace': 'Segoe UI', 'color': MUTED},
    },
    'visualStyles': {'*': {'*': {
        'title': [{'show': True, 'fontFamily': 'Segoe UI Semibold', 'fontSize': 12, 'fontColor': {'solid': {'color': NAVY}}}],
        'background': [{'show': True, 'color': {'solid': {'color': '#FFFFFF'}}, 'transparency': 0}],
        'border': [{'show': True, 'color': {'solid': {'color': LINE}}, 'radius': 8}],
    }}}
}
write(REPORT / 'StaticResources/RegisteredResources/DataCoGlobal-cf3d8bc0.json', theme)
write(ROOT / 'DataCo_Theme_BleuBlanc.json', theme)

page_titles = {
    '01_ensemble': ('Vue globale', 'Coûts simulés, service et politique de réapprovisionnement'),
    '02_stock': ('Niveau de stock', 'Évolution quotidienne, ruptures et couverture'),
    '03_comparaison': ('Comparaison des scénarios', 'Coût, service et intervalles de confiance à 95 %'),
    '04_couts_fiabilite': ('Coûts et fiabilité', 'Structure du coût et sensibilité des politiques'),
}
nav_names = list(page_titles)
nav_labels = ['Vue globale', 'Stock', 'Scénarios', 'Coûts et fiabilité']
main = PAGES / '01_ensemble'
context_template = read(PAGES / '02_stock/visuals/product_context/visual.json')

for page_id, (heading, subheading) in page_titles.items():
    page = PAGES / page_id
    meta = read(page / 'page.json')
    meta.update(width=1440, height=940, displayOption='FitToPage')
    meta['objects']['background'] = obj(dict(color=color(BG), transparency=lit(0)))
    write(page / 'page.json', meta)
    vs = {p.parent.name: read(p) for p in (page / 'visuals').glob('*/visual.json')}
    for v in vs.values():
        style(v)
    # Reuse the native navigation objects, with their existing page actions intact.
    vs['navigation'] = textbox('navigation', [''], 20, 820, 1400, 90, True, True, z=0)
    vs['nav_brand'] = textbox('nav_brand', ['Accès rapide', 'Explorer les analyses'], 36, 838, 254, 56, size=16)
    vs['nav_help'] = textbox('nav_help', ['DataCo Global  |  Pilotage logistique                 Navigation dans Desktop : Ctrl + clic'], 20, 916, 1100, 24, size=11)
    for i, dest in enumerate(nav_names):
        v = vs['nav_' + dest]
        place(v, 322 + i * 270, 840, 254, 48, 30)
        objects = v['visual']['objects']
        for entry in objects['text']:
            if entry.get('selector', {}).get('id') == 'default':
                entry['properties'].update(text=lit(nav_labels[i]), fontSize=lit(12), fontColor=color('#FFFFFF' if dest == page_id else NAVY))
        for entry in objects['fill']:
            state = entry.get('selector', {}).get('id')
            if state:
                entry['properties']['fillColor'] = color(BLUE if state == 'hover' or dest == page_id else '#EAF3FB')
        style(v, background=False, border=False)
    # A compact title, four equally spaced KPI cards, then the dashboard body.
    old_titles = [n for n, v in vs.items() if v['visual']['visualType'] == 'textbox' and n not in ('navigation', 'nav_brand', 'nav_help')]
    for n in old_titles:
        vs[n] = textbox(n, ['DataCo Global  |  ' + heading, subheading], 20, 12, 1400, 64, size=24)
    top = [('Cout moyen', 'Coût annuel moyen (USD)'), ('Service moyen', 'Service en volume'), ('Ruptures moyennes', 'Unités manquantes · moyenne'), ('Commandes moyennes', 'Commandes · moyenne')]
    for i, (measure, label) in enumerate(top):
        vs[f'kpi_{i}'] = card(f'kpi_{i}', measure, label, 20 + i * 354, 88)
    # The right panel is a native background container plus a live product card and slicer.
    vs['control_panel'] = textbox('control_panel', ['Panneau de contrôle'], 1102, 236, 318, 562, True, True, size=17, z=0)
    ctx = copy.deepcopy(context_template)
    place(ctx, 1118, 298, 286, 112)
    style(ctx, 'Produit simulé', background=False, border=False)
    ctx['visual']['objects']['labels'][0]['properties']['fontSize'] = lit(15)
    vs['product_context'] = ctx
    sc = vs['scenario']
    place(sc, 1118, 430, 286, 180)
    style(sc, 'Scénario de simulation', background=False, border=False)
    sc['visual']['objects']['data'] = obj({'mode': lit('Basic')})
    sc['visual']['objects']['header'] = obj({'show': lit(False)})
    sc['visual']['objects']['items'] = obj({'fontColor': color(NAVY), 'textSize': lit(12), 'background': color('#FFFFFF')})
    vs['control_hint'] = textbox('control_hint', ['Lecture du rapport', 'Sélectionner un scénario pour', 'filtrer les visuels de la page.', '', 'Simulation : relancer depuis Excel,', 'puis actualiser dans Power BI.'], 1118, 642, 286, 142, size=14)

    if page_id == '01_ensemble':
        layout = {'cout': (20, 236, 590, 268), 'service': (626, 236, 460, 268), 'disponibilite': (626, 520, 222, 110), 'couverture': (864, 520, 222, 110), 'decision': (626, 646, 460, 152)}
        style(vs['disponibilite'], 'Disponibilité quotidienne')
        style(vs['couverture'], 'Couverture (jours)')
        style(vs['decision'], 'Politique recommandée')
        vs['decision']['visual']['objects']['labels'][0]['properties']['color'] = color(GREEN)
        summary = {'$schema': SCHEMA, 'name': 'synthese_scenarios', 'visual': {'visualType': 'tableEx', 'query': {'queryState': {'Values': {'projections': [projection('Nom', 'Scénario', measure=False), projection('Cout moyen', 'Coût (USD)'), projection('Service moyen', 'Service'), projection('Stock moyen', 'Stock moyen'), projection('Commandes moyennes', 'Commandes')]}}}}}
        place(summary, 20, 520, 590, 278)
        vs[summary['name']] = style(summary, 'Synthèse des scénarios')
    elif page_id == '02_stock':
        layout = {'stock': (20, 236, 1066, 300), 'ruptures': (20, 552, 738, 246), 'stock_moyen': (774, 552, 312, 110)}
        vs['stock_note'] = textbox('stock_note', ['Journal de simulation', 'Les courbes représentent le run 1.', 'Le stock moyen porte sur l’ensemble', 'des répétitions simulées.'], 774, 678, 312, 120, True, True, size=14)
    elif page_id == '03_comparaison':
        layout = {'ic': (20, 236, 1066, 210), 'apparie': (20, 462, 500, 336), 'cout_service': (536, 462, 550, 336)}
    else:
        layout = {'decomposition': (20, 236, 590, 266), 'sensibilite': (626, 236, 460, 266), 'indicateurs': (20, 518, 1066, 280)}
        style(vs['sensibilite'], 'Sensibilité · commandes simultanées')
    for name, pos in layout.items():
        place(vs[name], *pos)
    for v in vs.values():
        save_visual(page, v)

print(json.dumps({'backup': str(backup), 'pages': len(page_titles), 'theme': str(ROOT / 'DataCo_Theme_BleuBlanc.json')}, ensure_ascii=False))
