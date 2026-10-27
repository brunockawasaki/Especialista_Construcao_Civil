# -*- coding: utf-8 -*-
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
import extract_bpm as E

OUTC = os.path.join(os.path.dirname(E.OUT), 'flows_compact')
os.makedirs(OUTC, exist_ok=True)

def compact(path):
    xml = E.diagram_xml(path)
    acts, lanes, pools, trans, arts = E.parse(xml)
    out = []
    # ordem de fluxo: ordenar atividades por BFS a partir de StartEvents
    nxt = {}
    for f, t, n, c in trans:
        nxt.setdefault(f, []).append((t, n or c))
    starts = [i for i, a in acts.items() if a['kind'].startswith('Evento:StartEvent')]
    order, seen = [], set()
    stack = list(reversed(starts))
    while stack:
        x = stack.pop()
        if x in seen or x not in acts: continue
        seen.add(x); order.append(x)
        for t, _ in reversed(nxt.get(x, [])):
            stack.append(t)
    for i in acts:
        if i not in seen: order.append(i)
    bylane = {}
    for i in order:
        a = acts[i]
        if a['kind'].startswith('Gateway') or a['kind'].startswith('Evento:Intermediate'):
            continue
        if not a['name'].strip(): continue
        bylane.setdefault(lanes.get(a['lane'], '(sem raia)'), []).append(a)
    for l, items in bylane.items():
        out.append('RAIA %s: ' % l + ' > '.join(x['name'].replace('\n', ' ').strip() + ('{' + x['doc'][:170] + '}' if x['doc'] else '') for x in items))
    # decisoes
    dec = []
    for i in order:
        a = acts[i]
        if a['kind'].startswith('Gateway'):
            outs = nxt.get(i, [])
            labels = [(acts.get(t, {}).get('name', '?').replace('\n', ' ').strip(), lb) for t, lb in outs]
            if len(outs) > 1:
                nm = a['name'].replace('\n', ' ').strip()
                dec.append('  ? %s%s -> %s' % (nm or '[decisao]', '', '; '.join('%s=>%s' % (lb or '-', n or '[gw]') for n, lb in labels)))
    if dec:
        out.append('DECISOES:'); out += dec
    notes = [x for x in arts if x[1] or x[2]]
    if notes:
        out.append('NOTAS: ' + ' || '.join((x[1] + ' ' + x[2]).strip().replace('\n', ' ') for x in notes))
    return '\n'.join(out)

tot = 0
for dp, dn, fn in os.walk(E.ROOT):
    for f in sorted(fn):
        if f.lower().endswith('.bpm'):
            p = os.path.join(dp, f)
            rel = os.path.relpath(p, E.ROOT)
            try:
                t = compact(p)
            except Exception as e:
                t = 'ERRO ' + str(e)
            dest = os.path.join(OUTC, rel.replace('\\', '__') + '.txt')
            open(dest, 'w', encoding='utf-8').write(t)
            tot += len(t)
print('chars total compacto:', tot)
