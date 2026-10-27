# -*- coding: utf-8 -*-
import json, re, unicodedata, os

HERE = os.path.dirname(__file__)

def norm(s):
    s = unicodedata.normalize('NFD', s.lower())
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')

PROC_MAP = {'novos negócios': 'Novos Negócios', 'incorporação make it': 'Incorporação Make It', 'incorporação gadens': 'Incorporação Gadens',
            'engenharia': 'Engenharia', 'gmo': 'GMO', 'suprimentos': 'Suprimentos', 'financeiro': 'Financeiro',
            'comercial gadens': 'Comercial Gadens', 'comercial make it': 'Comercial Make It'}

def proc_of(t):
    p = re.sub(r'\s+', ' ', (t['custom'].get('Processo') or '')).strip()
    if p.startswith('Engenharia'): p = 'Engenharia'   # inclui "Engenharia e GMO"
    if p in PROC_MAP.values(): return p
    for g in t['tags']:
        if g in PROC_MAP: return PROC_MAP[g]
    return 'Sem processo definido'

def any_(p, n): return re.search(p, n) is not None

def nature_of(t, n):
    tipo = norm(re.sub(r'\s+', ' ', (t['custom'].get('Tipo de Ação') or '')).strip())
    nm = t['name'].strip()
    if tipo == 'documento normativo' or re.match(r'^\W*elaborar (pop|it|politica|regimento)', n): return 'Documento normativo'
    if tipo == 'indicador' or re.match(r'^\W*implementar indicador|^\W*indice de', n): return 'Indicador'
    if 'treinamento' in n or 'treinar' in n: return 'Treinamento'
    if any_(r'alcada|trava|bloque|segrega|perfis? de acesso|concilia|due diligence|aprovacao de|dados banc|homologa|auditoria|catraca|controle de acesso|duplicidade|penalidade|nao conformidade|autonomia', n): return 'Controle'
    if any_(r'automat|integra|parametr|criar campo|campo no|habilitar|modulo|rpa|\bbi\b|clicksign|approvo|\bcv\b|mega|prevision|robo|plataforma|sistema|g-control|implantar|implantacao', n): return 'Automação ou sistema'
    return 'Definição ou decisão'

STRONG_CTRL = r'trava|bloque|alcada|comissao|segrega|autonomia'

def base_scores(t):
    n = norm(t['name'])
    nat = nature_of(t, n)
    avaliar = bool(re.match(r'^\W*(avaliar|estudar|analisar a possib|verificar|testar|discutir|entender)', n))
    # ---- risco
    R = 1
    if any_(r'due diligence|medic|aditivo|distrato|repasse|permuta|spe\b|consorcio|venda direta|credito|\bnf\b|nota fiscal|notas fiscais|impost|fiscal|tribut|lgpd|compliance|inventario|estoque|perdas|desvio|cotac|compras|pedido|preco|orcamento|contrato|minuta|fornecedor|empreiteir|seguranca|nao conformidade|vistoria|cvco|\bcef\b|financeiro|pagamento|caixa|cnpj|escritura|aquisicao|terreno', n): R = 3
    elif any_(r'\bpop\b|politica|regimento|\bit\b|automat|integra|parametr|campo|modulo|sistema|cadastro|plataforma|indicador', n): R = 2
    if any_(r'alcada|autonomia|trava|bloque|segrega|perfis? de acesso|revisao dos acessos|catraca|controle de acesso|dados banc|conta de pj|pagamentos? manua|adiantament|concilia|comissao|devolucao de sinal|nao pagamento de sinal|duplicidade|homologa|cadastro de fornecedor|auditoria|valor do contrato|contrato livre|sem contrato|penalidad|politica comercial|controles para garantir|cartao corporativo|mutuo|boleto|aliquota|retencao|pagamentos? indevid|excec|autenticidade|due diligence de fornecedor|implementar sistema de due', n): R = 4
    if nat in ('Indicador', 'Treinamento') and not any_(r'pagamentos? indevid|avanco fisico|nfs com erro|duplicidade', n): R = min(R, 2)
    if avaliar and not any_(STRONG_CTRL, n): R = min(R, 3)
    # ---- financeiro
    F = 1
    if any_(r'nf\b|nota|fiscal|fornecedor|empreiteir|obra|planejamento|contrato|minuta|spe\b|cnpj|due diligence|projeto|vistoria|seguranca|qualidade', n): F = 3
    elif any_(r'\bpop\b|politica|regimento|automat|integra|sistema|cadastro|campo', n): F = 2
    if any_(r'pagamento|adiantament|medic|repasse|permuta|devolucao|sinal|comissao|compras|cotac|orcamento|preco|saving|custo|viabilidade|aquisicao|terreno|multas|juros|impost|concilia|contas a pagar|mutuo|inventario|estoque|perdas|aditivo|credito|venda', n): F = 4
    if nat == 'Indicador' and not any_(r'pagamentos? indevid', n): F = min(F, 3)
    if nat == 'Treinamento': F = min(F, 2)
    if avaliar: F = min(F, 3)
    # ---- obra e entrega
    O = 1
    if any_(r'venda|comercial|proposta|fornecedor', n): O = 2
    if any_(r'\bobra\b|obras|projeto|qualidade|seguranca|estoque|almoxarifado|medic|orcamento|comite', n): O = 3
    if any_(r'planejamento|histograma|inicio de obra|pre-?obra|projeto executivo|aprovacao de projetos|obra sem contrato|servico iniciado|abandono|empreiteir|contratacao de|suprimentos|solicitac\w* de compra|entrega de chaves|cronograma|lancamento|restric|alvara', n): O = 4
    if any_(r'interrupcao de obra|cvco|memorial|memoriais cef|\bpls\b|\brae\b', n): O = 5
    if nat in ('Indicador', 'Treinamento'): O = min(O, 3)
    if avaliar: O = min(O, 3)
    # ---- legal / CEF / compliance
    L = 1
    if any_(r'\bpop\b|\bit\b|procedimento|cadastro', n): L = 2
    if any_(r'politica|regimento|alcada|homologa|credito|aditivo', n): L = 3
    if any_(r'lgpd|compliance|\bcef\b|\bcno\b|fiscal|tribut|impost|escritura|cvco|alvara|spe\b|cnpj|memorial|distrato|notificac|contrato|due diligence|serasa|spc|certidao|matricula|licenc', n): L = 4
    if any_(r'seguranca do trabalho|interrupcao de obra por motivos de seguranca', n): L = 5
    if nat in ('Indicador', 'Treinamento'): L = min(L, 2)
    if avaliar: L = min(L, 3)
    # ---- dependencia
    D = 1
    if any_(r'avaliar|treinamento|indicador|comunicacao', n): D = 2
    if any_(r'\bpop\b|regimento|automat|trava|bloque|alerta|relatorio|sistema', n): D = 3
    if any_(r'definir|definicao|padroniz|politica|integr|parametr|cadastro|fluxo|modulo|plataforma|alcada|campo no cv', n): D = 4
    if any_(r'plano de alcadas|alcada de aprovacao|estrutura do empreendimento|definir .*responsab|cadastro de itens|cadastro de fornecedor|fluxo de contratos|integrar cv com mega|integracao prevision', n): D = 5
    if nat == 'Indicador': D = min(D, 2)
    if nat == 'Treinamento': D = min(D, 2)
    # ---- esforco
    if nat == 'Documento normativo': E = 4 if 'politica' in n else 3
    elif nat == 'Indicador': E = 2
    elif nat == 'Treinamento': E = 2
    elif avaliar or n.startswith(('definir', 'alinhar', 'incluir', 'formalizar')): E = 2
    elif any_(r'automat|integra|parametr|criar campo|habilitar|modulo|implantar|implementar', n): E = 3
    else: E = 3
    if any_(r'contratacao de uma pessoa|criacao de equipe|bim|aplicativo|portal de fornecedores|banco de dados', n): E = 5
    return dict(R=R, F=F, O=O, L=L, D=D, E=E, cef='Não', nat=nat)

FACTORS = {'Incorporação Make It': 1.1, 'Incorporação Gadens': 1.1, 'Engenharia': 1.084, 'Novos Negócios': 1.067, 'Suprimentos': 1.05,
           'Financeiro': 1.05, 'GMO': 1.034, 'Comercial Make It': 1.034, 'Comercial Gadens': 1.017, 'Sem processo definido': 1.0}
ADJ = {'Controle': 0.3, 'Documento normativo': 0.0, 'Automação ou sistema': 0.0, 'Definição ou decisão': 0.0, 'Treinamento': -0.2, 'Indicador': -0.3}
W = dict(R=0.30, F=0.20, O=0.20, L=0.15, D=0.15)

from decimal import Decimal, ROUND_HALF_UP
def rnd(x, d):
    q = Decimal(1).scaleb(-d)
    return float(Decimal(repr(round(x, 9))).quantize(q, rounding=ROUND_HALF_UP))

PROC_NOTES = {'Incorporação Make It': (5, 5, 5), 'Incorporação Gadens': (5, 5, 5), 'Engenharia': (5, 5, 4), 'Novos Negócios': (5, 3, 5),
              'Suprimentos': (5, 4, 3), 'Financeiro': (4, 5, 3), 'GMO': (4, 4, 3), 'Comercial Make It': (4, 4, 3),
              'Comercial Gadens': (4, 3, 3), 'Sem processo definido': (3, 3, 3)}
FACTORS = {k: rnd(0.9 + 0.2 * (rnd(sum(v) / 3, 2) - 1) / 4, 3) for k, v in PROC_NOTES.items()}

def final(sc, proc):
    media = rnd(sc['R'] * W['R'] + sc['F'] * W['F'] + sc['O'] * W['O'] + sc['L'] * W['L'] + sc['D'] * W['D'], 2)
    nota = rnd(media * FACTORS.get(proc, 1.0) + ADJ[sc['nat']], 2)
    return media, nota
