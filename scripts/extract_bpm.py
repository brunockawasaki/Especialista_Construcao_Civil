# -*- coding: utf-8 -*-
import os, io, zipfile, re, sys
import xml.etree.ElementTree as ET

ROOT = r"C:\Users\bruno.kawasaki\OneDrive - Grupo Gadens\Processos - Documentos\Processos Grupo Gadens\1. Projetos\1. Projeto Next\2. Processos"
OUT = r"C:\Users\BRUNO~1.KAW\AppData\Local\Temp\claude\C--Users-bruno-kawasaki-OneDrive---Grupo-Gadens-1--Bruno-Kawasaki-Projeto-teste\b3f3cea1-f242-4fed-9cfc-7187c96cb2ce\scratchpad\flows"
os.makedirs(OUT, exist_ok=True)

def ln(tag):
    return tag.split('}')[-1]

def diagram_xml(path):
    z = zipfile.ZipFile(path)
    for n in z.namelist():
        if n.endswith('.diag'):
            raw = z.read(n)
            z2 = zipfile.ZipFile(io.BytesIO(raw))
            return z2.read('Diagram.xml')
    return None

def text_of(el):
    return ' '.join((el.text or '').split()) if el is not None else ''

def parse(xml):
    root = ET.fromstring(xml)
    # id -> info
    acts = {}
    lanes = {}
    pools = {}
    trans = []
    arts = []
    for el in root.iter():
        t = ln(el.tag)
        if t == 'Pool':
            pools[el.get('Id')] = el.get('Name') or ''
        elif t == 'Lane':
            lanes[el.get('Id')] = el.get('Name') or ''
        elif t == 'Activity':
            kind = 'Tarefa'
            doc = ''
            for ch in el:
                c = ln(ch.tag)
                if c == 'Route':
                    kind = 'Gateway'
                    gt = ch.get('GatewayType') or ''
                    kind = 'Gateway' + (':' + gt if gt else '')
                elif c == 'Event':
                    kind = 'Evento'
                    for e2 in ch:
                        kind = 'Evento:' + ln(e2.tag)
                elif c == 'BlockActivity':
                    kind = 'Subprocesso'
                elif c == 'Implementation':
                    for e2 in ch:
                        k2 = ln(e2.tag)
                        if k2 in ('Task',):
                            for e3 in e2:
                                kind = 'Tarefa:' + ln(e3.tag)
                        else:
                            kind = 'Tarefa'
                elif c == 'Documentation':
                    import html as _h
                    doc = ' '.join(_h.unescape(re.sub(r'<[^>]+>', ' ', ''.join(ch.itertext()))).split())
            lane = ''
            for g in el.iter():
                if ln(g.tag) == 'NodeGraphicsInfo':
                    lane = g.get('LaneId') or lane
                    break
            acts[el.get('Id')] = dict(name=el.get('Name') or '', kind=kind, doc=doc, lane=lane)
        elif t == 'Transition':
            cond = ''
            for ch in el:
                if ln(ch.tag) == 'Condition':
                    cond = (ch.text or '').strip()
            trans.append((el.get('From'), el.get('To'), el.get('Name') or '', cond))
        elif t == 'Artifact':
            nm = el.get('Name') or ''
            txt = ''
            for ch in el.iter():
                if ln(ch.tag) == 'TextAnnotation':
                    txt = (ch.text or '').strip()
            arts.append((el.get('ArtifactType') or '', nm, txt))

    # --- atribuicao de raia por geometria ---
    def ycoord(el):
        for g in el.iter():
            if ln(g.tag) == 'NodeGraphicsInfo':
                for c in g:
                    if ln(c.tag) == 'Coordinates':
                        try: return float(c.get('YCoordinate'))
                        except: return None
        return None
    pool_lanes = {}
    for pl in root.iter():
        if ln(pl.tag) == 'Pool':
            ls = []
            for ln_el in pl.iter():
                if ln(ln_el.tag) == 'Lane':
                    y = ycoord(ln_el); h = None
                    for g in ln_el.iter():
                        if ln(g.tag) == 'NodeGraphicsInfo': h = float(g.get('Height') or 0); break
                    if y is not None and h is not None:
                        ls.append((y, y + h, ln_el.get('Name') or '', ln_el.get('Id')))
            pool_lanes[pl.get('Process')] = ls
    for wp in root.iter():
        if ln(wp.tag) == 'WorkflowProcess':
            ls = pool_lanes.get(wp.get('Id'), [])
            if not ls: continue
            for a in wp.iter():
                if ln(a.tag) == 'Activity' and a.get('Id') in acts:
                    y = ycoord(a)
                    if y is None: continue
                    yc = y + 15
                    for (y0, y1, nm, lid) in ls:
                        if y0 <= yc < y1:
                            lanes[lid] = nm
                            acts[a.get('Id')]['lane'] = lid
                            break
    return acts, lanes, pools, trans, arts

def render(path):
    xml = diagram_xml(path)
    if xml is None:
        return 'SEM .diag'
    acts, lanes, pools, trans, arts = parse(xml)
    out = []
    out.append('POOLS: ' + ' | '.join(v for v in pools.values() if v))
    # group by lane
    bylane = {}
    for i, a in acts.items():
        bylane.setdefault(lanes.get(a['lane'], '(sem raia)'), []).append(a)
    for l, items in bylane.items():
        out.append('RAIA: %s' % l)
        for a in items:
            if a['name'] or a['doc']:
                s = '  - [%s] %s' % (a['kind'], a['name'].replace('\n', ' '))
                if a['doc']:
                    s += '  // DOC: ' + a['doc']
                out.append(s)
    out.append('FLUXO (de -> para):')
    for f, t, n, c in trans:
        fa = acts.get(f, {}).get('name', '?') or '[%s]' % acts.get(f, {}).get('kind', '?')
        ta = acts.get(t, {}).get('name', '?') or '[%s]' % acts.get(t, {}).get('kind', '?')
        lab = (' (%s)' % (n or c)) if (n or c) else ''
        out.append('  %s -> %s%s' % (fa.replace('\n', ' '), ta.replace('\n', ' '), lab))
    if arts:
        out.append('ANOTACOES/ARTEFATOS:')
        for a in arts:
            if a[1] or a[2]:
                out.append('  - %s: %s %s' % (a[0], a[1], a[2]))
    return '\n'.join(out), len(acts), len(trans)

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        p = os.path.join(ROOT, '0. Macro Grupo Gadens', 'FL.GG.PRO.01 - Macroprocesso Grupo Gadens Incorporadoras.bpm')
        r = render(p)
        print(r[0][:6000]); print(r[1], r[2])
        sys.exit()
    idx = []
    for dp, dn, fn in os.walk(ROOT):
        for f in sorted(fn):
            if f.lower().endswith('.bpm'):
                p = os.path.join(dp, f)
                rel = os.path.relpath(p, ROOT)
                try:
                    res = render(p)
                    if isinstance(res, str):
                        idx.append((rel, 0, 0, res)); continue
                    txt, na, nt = res
                except Exception as e:
                    idx.append((rel, 0, 0, 'ERRO ' + str(e)[:80])); continue
                dest = os.path.join(OUT, rel.replace('\\', '__') + '.txt')
                with open(dest, 'w', encoding='utf-8') as fh:
                    fh.write(txt)
                idx.append((rel, na, nt, ''))
    with open(os.path.join(OUT, '_INDEX.txt'), 'w', encoding='utf-8') as fh:
        for rel, na, nt, err in idx:
            fh.write('%4d atv %4d lig  %s %s\n' % (na, nt, rel, err))
    print(len(idx), 'fluxos;', sum(1 for i in idx if i[1] > 0), 'com atividades')
