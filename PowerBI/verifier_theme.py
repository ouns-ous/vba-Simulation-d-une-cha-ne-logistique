"""Validate PBIR, visual formatting, model references and layout against Microsoft schemas."""
import json
import re
import argparse
from pathlib import Path
from datetime import datetime, timezone
import requests
import jsonschema

parser = argparse.ArgumentParser()
parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parent)
parser.add_argument('--report', default='DataCo.Report')
args = parser.parse_args()
ROOT = args.root.resolve()
REPORT = ROOT / args.report
cache = {}
def schema(url):
    url = url.replace('schema.embedded.json', 'schema-embedded.json')
    if url not in cache:
        response = requests.get(url, timeout=40)
        response.raise_for_status()
        cache[url] = json.loads(response.text.replace('schema.embedded.json', 'schema-embedded.json'))
    return cache[url]

errors = []
count = 0
for p in (REPORT / 'definition').rglob('*.json'):
    data = json.loads(p.read_text(encoding='utf-8-sig'))
    if '$schema' in data:
        sch = schema(data['$schema'])
        resolver = jsonschema.RefResolver(base_uri=data['$schema'], referrer=sch, handlers={'https': schema})
        validator = jsonschema.Draft7Validator(sch, resolver=resolver)
        for e in validator.iter_errors(data):
            errors.append(f'{p.relative_to(ROOT)}: {list(e.path)}: {e.message[:500]}')
        count += 1

theme_schema = schema('https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/Report%20Theme%20JSON%20Schema/reportThemeSchema-2.159.json')
theme = json.loads((ROOT / 'DataCo_Theme_BleuBlanc.json').read_text(encoding='utf-8'))
for e in jsonschema.Draft7Validator(theme_schema).iter_errors(theme):
    errors.append(f'Theme: {list(e.path)}: {e.message[:500]}')

def unwrap(v):
    if isinstance(v, list):
        return [unwrap(x) for x in v]
    if not isinstance(v, dict):
        return v
    if set(v) == {'expr'} and 'Literal' in v['expr']:
        raw = v['expr']['Literal']['Value']
        if raw in ('true', 'false'):
            return raw == 'true'
        if raw.startswith("'"):
            return raw[1:-1].replace("''", "'")
        return float(raw.rstrip('DL'))
    return {k: unwrap(x) for k, x in v.items()}

fields = {}
for p in (ROOT / 'DataCo.SemanticModel/definition/tables').glob('*.tmdl'):
    text = p.read_text(encoding='utf-8-sig')
    fields[p.stem] = set(a or b for a, b in re.findall(r"^\t(?:column|measure) (?:'([^']+)'|([^\s=]+))", text, re.M))

def walk(v, p):
    if isinstance(v, dict):
        for kind in ('Column', 'Measure'):
            if kind in v:
                f = v[kind]
                entity = f.get('Expression', {}).get('SourceRef', {}).get('Entity')
                if entity and f.get('Property') not in fields.get(entity, set()):
                    errors.append(f'{p}: unknown field {entity}.{f.get("Property")}')
        for value in v.values():
            walk(value, p)
    elif isinstance(v, list):
        for x in v:
            walk(x, p)

visual_count = 0
for page in (REPORT / 'definition/pages').iterdir():
    if not page.is_dir():
        continue
    meta = json.loads((page / 'page.json').read_text(encoding='utf-8'))
    foreground = []
    for p in (page / 'visuals').glob('*/visual.json'):
        v = json.loads(p.read_text(encoding='utf-8'))
        pos = v['position']
        if pos['x'] < 0 or pos['y'] < 0 or pos['x'] + pos['width'] > meta['width'] or pos['y'] + pos['height'] > meta['height']:
            errors.append(f'{page.name}/{v["name"]}: outside canvas')
        if pos['z'] > 0:
            foreground.append((v['name'], pos))
        walk(v.get('visual', {}).get('query', {}), f'{page.name}/{v["name"]}')
        visual_count += 1
        typ = v['visual']['visualType']
        if typ in ('card', 'tableEx', 'slicer', 'lineChart', 'scatterChart', 'clusteredColumnChart'):
            styles = {k: [unwrap(o['properties']) for o in objs] for k, objs in v['visual'].get('objects', {}).items()}
            trial = {'name': 'Validation', 'visualStyles': {typ: {'*': styles}}}
            for e in jsonschema.Draft7Validator(theme_schema).iter_errors(trial):
                errors.append(f'{page.name}/{v["name"]} format: {list(e.path)}: {e.message[:500]}')
    for i, (a, pa) in enumerate(foreground):
        for b, pb in foreground[i + 1:]:
            if min(pa['x'] + pa['width'], pb['x'] + pb['width']) - max(pa['x'], pb['x']) > 0.1 and min(pa['y'] + pa['height'], pb['y'] + pb['height']) - max(pa['y'], pb['y']) > 0.1:
                errors.append(f'{page.name}: overlap {a}/{b}')

result = {'checked_at_utc': datetime.now(timezone.utc).isoformat(), 'json_files': count, 'visuals': visual_count, 'errors': errors, 'native_desktop_preview': 'Not verified: Computer Use kernel unavailable (failed to write kernel assets).'}
(ROOT / 'verification_theme.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(result, ensure_ascii=False, indent=2))
raise SystemExit(bool(errors))
