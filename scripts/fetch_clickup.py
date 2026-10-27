# -*- coding: utf-8 -*-
# Leitura (GET) da lista do ClickUp. O token vem da variavel de ambiente, nunca e gravado em arquivo.
import os, json, urllib.request, urllib.error, sys
TOKEN = os.environ.get('CLICKUP_TOKEN', '')
LIST_ID = '901321384887'
OUT = os.path.join(os.path.dirname(__file__), 'clickup_abertas.json')

def get(url):
    req = urllib.request.Request(url, headers={'Authorization': TOKEN, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode('utf-8'))

tasks = []
page = 0
while True:
    d = get('https://api.clickup.com/api/v2/list/%s/task?include_closed=false&subtasks=true&page=%d' % (LIST_ID, page))
    batch = d.get('tasks', [])
    tasks += batch
    if len(batch) < 100 or d.get('last_page'):
        break
    page += 1

def cf_value(cf):
    v = cf.get('value')
    if v is None or v == '':
        return ''
    t = cf.get('type')
    if t in ('drop_down',):
        opts = cf.get('type_config', {}).get('options', [])
        for o in opts:
            if o.get('orderindex') == v or o.get('id') == v:
                return o.get('name', '')
        return str(v)
    if t == 'users':
        return ', '.join(u.get('username', '') for u in v) if isinstance(v, list) else str(v)
    if t == 'labels':
        opts = {o['id']: o.get('label', '') for o in cf.get('type_config', {}).get('options', [])}
        return ', '.join(opts.get(x, str(x)) for x in v) if isinstance(v, list) else str(v)
    return str(v)

rows = []
for t in tasks:
    st = (t.get('status') or {}).get('status', '')
    cfs = {c.get('name'): cf_value(c) for c in t.get('custom_fields', [])}
    rows.append(dict(
        id=t['id'], name=t.get('name', '').strip(), status=st, url=t.get('url'),
        tags=[g['name'] for g in t.get('tags', [])],
        assignees=[a.get('username', '') for a in t.get('assignees', [])],
        due=t.get('due_date'), created=t.get('date_created'),
        parent=t.get('parent'), custom=cfs))
json.dump(rows, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False)
from collections import Counter
print(len(rows), 'tarefas abertas lidas')
print(Counter(r['status'] for r in rows))
print('campos personalizados vistos:', sorted({k for r in rows for k in r['custom']}))
