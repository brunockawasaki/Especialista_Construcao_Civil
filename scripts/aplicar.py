# -*- coding: utf-8 -*-
import json, re, collections
from score_base import *
NAT = {'C':'Controle','D':'Documento normativo','A':'Automação ou sistema','F':'Definição ou decisão','T':'Treinamento','I':'Indicador'}
def carregar():
    T = json.load(open('clickup_todas.json', encoding='utf-8'))
    byid = {t['id']: t for t in T}
    man = {}
    for ln in open('notas_manuais.txt', encoding='utf-8'):
        ln = ln.strip()
        if not ln or ln.startswith('#'): continue
        ids, rest = ln.split(':', 1)
        f = rest.split()
        R, F, O, L, D, E = map(int, f[:6]); N = NAT[f[6]]; cef = 'Sim' if len(f) > 7 and f[7] == 'S' else 'Não'
        for i in [x.strip() for x in ids.split(',')]:
            man[i] = dict(R=R, F=F, O=O, L=L, D=D, E=E, nat=N, cef=cef)
    return T, byid, man
def pontuar():
    T, byid, man = carregar()
    out = []
    for t in T:
        p = proc_of(t)
        if t['id'] in man:
            s = dict(man[t['id']]); origem = 'Manual (proposta)'
        else:
            s = base_scores(t); origem = 'Regra base'
            for k in 'RFOLD': s[k] = min(s[k], 3)
            s['R'] = min(s['R'], 3)
        media, nota = final(s, p)
        out.append(dict(t=t, proc=p, s=s, media=media, nota=nota, origem=origem))
    return out, byid, man
if __name__ == '__main__':
    out, byid, man = pontuar()
    miss = [i for i in man if i not in byid]
    print('IDs manuais sem correspondencia:', miss)
    print('manuais:', sum(1 for o in out if o['origem'].startswith('Manual')), 'de', len(out))
    out.sort(key=lambda o: -o['nota'])
    c = collections.Counter()
    for o in out:
        n = o['nota']; c['>=4.0' if n >= 4 else '3.3-3.99' if n >= 3.3 else '2.5-3.29' if n >= 2.5 else '<2.5'] += 1
    print(c)
    veto = sum(1 for o in out if o['s']['R'] == 5)
    print('risco 5:', veto, '| cef sim:', sum(1 for o in out if o['s']['cef']=='Sim'))
    print('nota rank 73:', out[72]['nota'], '| rank 269:', out[268]['nota'])
    nm = [ (i,o) for i,o in enumerate(out,1) if o['origem']=='Regra base' and i<=330]
    print('nao-manuais nos 330 primeiros:', len(nm))
    for i,o in nm[:40]:
        print(i, o['t']['id'], '%.2f'%o['nota'], o['t']['status'][:2], re.sub(r'\s+',' ',o['t']['name'])[:80])
