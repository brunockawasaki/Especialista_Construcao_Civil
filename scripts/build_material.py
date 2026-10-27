# -*- coding: utf-8 -*-
import sys, os, json, warnings
warnings.filterwarnings('ignore')
import openpyxl
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

XLSX, OUT = sys.argv[1], sys.argv[2]
SCR = os.path.dirname(__file__)
wv = openpyxl.load_workbook(XLSX, data_only=True)
M, P, T = wv['Matriz'], wv['Processos'], wv['Tarefas']
urls = {t['id']: t['url'] for t in json.load(open(os.path.join(SCR, 'clickup_abertas.json'), encoding='utf-8'))}

NAVY = '1F3864'
doc = Document()

# ---------------- estilos base ----------------
def set_fonts(rpr, size, bold=None, italic=None, color=None, underline=False):
    f = rpr.find(qn('w:rFonts'))
    if f is None:
        f = OxmlElement('w:rFonts'); rpr.insert(0, f)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        f.set(qn(a), 'Arial')
    for tag in ('w:sz', 'w:szCs'):
        e = rpr.find(qn(tag))
        if e is None:
            e = OxmlElement(tag); rpr.append(e)
        e.set(qn('w:val'), str(int(round(size * 2))))

st = doc.styles['Normal']
st.font.name = 'Arial'; st.font.size = Pt(10)
set_fonts(st.element.get_or_add_rPr(), 10)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15

def style_heading(name, size, color):
    s = doc.styles[name]
    s.font.name = 'Arial'; s.font.size = Pt(size); s.font.bold = True
    s.font.color.rgb = RGBColor.from_string(color)
    rp = s.element.get_or_add_rPr()
    set_fonts(rp, size)
    rf = rp.find(qn('w:rFonts'))
    for a in ('w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme', 'w:cstheme'):
        if rf.get(qn(a)) is not None: del rf.attrib[qn(a)]
    s.paragraph_format.space_before = Pt(14 if name == 'Heading 1' else 10)
    s.paragraph_format.space_after = Pt(6)
    s.paragraph_format.keep_with_next = True
style_heading('Heading 1', 13, NAVY)
style_heading('Heading 2', 11, NAVY)
for n in ('List Bullet', 'List Number'):
    s = doc.styles[n]; s.font.name = 'Arial'; s.font.size = Pt(10)
    set_fonts(s.element.get_or_add_rPr(), 10)
    s.paragraph_format.space_after = Pt(3)

sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = Cm(1.8); sec.bottom_margin = Cm(1.8); sec.left_margin = Cm(2.5); sec.right_margin = Cm(1.8)

def R(p, text, size=10, bold=False, italic=False, color=None, underline=False):
    r = p.add_run(text)
    r.font.name = 'Arial'; r.font.size = Pt(size); r.bold = bold; r.italic = italic
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if underline: r.underline = True
    set_fonts(r._r.get_or_add_rPr(), size)
    return r

def para(text='', size=10, bold=False, italic=False, color=None, align=None, after=None, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if text: R(p, text, size, bold, italic, color)
    if align: p.alignment = align
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p

def rich(parts, style=None, after=None):
    """parts: lista de (texto, negrito)"""
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    for t, b in parts: R(p, t, bold=b)
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p

def bullet(text, bold_lead=None):
    p = doc.add_paragraph(style='List Bullet')
    if bold_lead: R(p, bold_lead, bold=True)
    R(p, text)
    return p

_num = [0]
def reset_num(): _num[0] = 0
def numbered(text, bold_lead=None):
    _num[0] += 1
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8); p.paragraph_format.first_line_indent = Cm(-0.8)
    p.paragraph_format.space_after = Pt(3)
    R(p, '%d.  ' % _num[0])
    if bold_lead: R(p, bold_lead, bold=True)
    R(p, text)
    return p

def h1(t): return doc.add_heading(t, level=1)
def h2(t): return doc.add_heading(t, level=2)

def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:shd')): tcPr.remove(e)
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)

def borders(tbl, color='BFBFBF', sz=4):
    tblPr = tbl._tbl.tblPr
    for e in tblPr.findall(qn('w:tblBorders')): tblPr.remove(e)
    b = OxmlElement('w:tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement('w:' + side)
        e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), str(sz)); e.set(qn('w:space'), '0'); e.set(qn('w:color'), color)
        b.append(e)
    anchor = None
    for tag in ('w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription'):
        anchor = tblPr.find(qn(tag))
        if anchor is not None: break
    if anchor is not None: anchor.addprevious(b)
    else: tblPr.append(b)

def cell_write(cell, text, size=9, bold=False, color=None, align=None, italic=False, link=None):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(1); p.paragraph_format.space_before = Pt(1); p.paragraph_format.line_spacing = 1.0
    if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if link:
        add_link(p, link, text, size)
    else:
        R(p, str(text), size, bold, italic, color)

def add_link(p, url, text, size):
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), rid)
    r = OxmlElement('w:r'); rpr = OxmlElement('w:rPr')
    c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); rpr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rpr.append(u)
    set_fonts(rpr, size)
    r.append(rpr)
    t = OxmlElement('w:t'); t.text = text; t.set(qn('xml:space'), 'preserve'); r.append(t)
    h.append(r); p._p.append(h)

def table(headers, rows, widths_cm, size=9, head_size=None, first_col_bold=False, center_cols=(), shade_first=False, caption=None):
    if caption:
        cp = para(caption, 9, True, color=NAVY, after=3); cp.paragraph_format.keep_with_next = True
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        cell_write(c, h, head_size or size, True, 'FFFFFF', 'c'); shade(c, NAVY)
    # cabecalho repete
    trPr = t.rows[0]._tr.get_or_add_trPr(); th = OxmlElement('w:tblHeader'); th.set(qn('w:val'), 'true'); trPr.append(th)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            link = None
            if isinstance(v, tuple): v, link = v
            cell_write(cells[i], '' if v is None else v, size, bold=(first_col_bold and i == 0), align=('c' if i in center_cols else None), link=link)
            if shade_first and i == 0: shade(cells[i], 'D9E1F2')
    for row in t.rows:
        # nao quebrar linha entre paginas
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement('w:cantSplit'); cs.set(qn('w:val'), 'true'); trPr.append(cs)
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def box(text_parts, fill='FFF2CC'):
    t = doc.add_table(rows=1, cols=1); t.autofit = False; borders(t, 'BF9000', 6)
    c = t.rows[0].cells[0]; c.width = Cm(16.7); shade(c, fill)
    c.text = ''
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(2)
    for tx, b in text_parts: R(p, tx, 10, b)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def fnum(x, d=2):
    return ('%.*f' % (d, x)).replace('.', ',')

# ---------------- dados da planilha ----------------
crit = []
for r in range(5, 10):
    crit.append(dict(nome=M.cell(r, 1).value, perg=M.cell(r, 2).value, peso=M.cell(r, 3).value, esc=[M.cell(r, c).value for c in range(4, 9)]))
cuts = [(M.cell(r, 1).value, M.cell(r, 2).value, M.cell(r, 3).value) for r in range(16, 20)]
veto_risco, veto_obra = M['C23'].value, M['C24'].value
nats = [(M.cell(r, 1).value, M.cell(r, 2).value, M.cell(r, 3).value) for r in range(28, 34)]
procs = [(P.cell(r, 1).value, P.cell(r, 2).value, P.cell(r, 3).value, P.cell(r, 4).value, P.cell(r, 5).value, P.cell(r, 6).value, P.cell(r, 7).value) for r in range(8, 18)]
fmin, fmax = P['B4'].value, P['B5'].value
anch = []
for r in range(5, T.max_row + 1):
    if T.cell(r, 24).value:
        g = lambda c: T.cell(r, c).value
        anch.append(dict(id=g(1), nome=g(2), proc=g(3), resp=g(4), status=g(5), nat=g(7), esf=g(8), n=[g(9), g(10), g(11), g(12), g(13)], cef=g(14),
                         media=g(15), fator=g(16), aj=g(17), nota=g(18), catn=g(19), catf=g(20), prio=g(21), ordem=g(23), why=g(25)))
anch.sort(key=lambda a: a['ordem'])
pesos = [c['peso'] for c in crit]
assert sum(pesos) == 100

# =====================================================================
# CAPA / ABERTURA
# =====================================================================
para('Como priorizar as atividades do ClickUp', 20, True, color=NAVY, after=2)
para('A lógica da matriz, os números e 12 atividades de exemplo', 11, italic=True, color='595959', after=2)
para('Projeto Next · Processos · 21/10/2026 · versão para análise e validação', 9, color='7F7F7F', after=10)

box([('Leia isto primeiro. ', True),
     ('Tudo neste material é proposta para você analisar e contestar. Os pesos, os cortes e as notas de processo são hipóteses minhas, não dado oficial. '
      'As atividades vêm do ClickUp (lista "Implantação de Melhorias (5W2H)", lida em 21/10/2026, 401 atividades abertas). '
      'A prioridade que já estava no ClickUp não foi usada em nenhuma conta.', False)])

para('Roteiro', 11, True, color=NAVY, after=2)
for t in ['1. O problema que a matriz resolve', '2. A ideia em uma página', '3. Os 5 critérios, os pesos e a régua de cada nota',
          '4. O peso do processo', '5. A natureza da ação', '6. Cortes e regras que passam por cima da conta', '7. Esforço e ordem de execução',
          '8. Um exemplo calculado passo a passo', '9. As 12 atividades de amostra', '10. O que é hipótese e como validar', '11. Decisões que dependem de você', 'Anexos: glossário e mapa da planilha']:
    para(t, 10, after=1)

# =====================================================================
h1('1. O problema que a matriz resolve')
para('O ClickUp tem hoje 963 atividades no total, sendo 401 abertas (sem contar concluídas e canceladas). Todas têm uma prioridade, mas a prioridade mistura três perguntas diferentes:')
bullet(' o quanto a atividade protege ou ganha para a empresa.', 'Importância:')
bullet(' quando precisa estar pronta.', 'Prazo:')
bullet(' o quanto custa fazer.', 'Esforço:')
para('Quando as três ficam na mesma nota, uma atividade importante e difícil perde para outra fácil e pouco importante. O Plano de Alçadas é um exemplo: vários fluxos oficiais (venda, compras, adiantamento, novos negócios) decidem "conforme alçada", mas a regra ainda não está pronta. É importante, é difícil e depende da diretoria. Numa nota que mistura esforço, ela cai.')
rich([('Objetivo da matriz: ', True), ('uma regra simples e escrita, que qualquer pessoa aplique e chegue à mesma categoria. As quatro categorias são as que você pediu: 1 Baixa, 2 Normal, 3 Alta e 4 Urgente.', False)])

# =====================================================================
h1('2. A ideia em uma página')
para('A matriz segue seis princípios. Cada número das próximas seções nasce de um deles.')
reset_num()
numbered(' A nota mede só a importância. Prazo e esforço ficam de fora e só servem para ordenar as atividades dentro de cada categoria.', 'Uma pergunta por vez.')
numbered(' O risco da atividade pesa mais que o processo em que ela está. Um processo de apoio, como o Financeiro, pode ter risco alto.', 'O risco manda.')
numbered(' Cada nota de 1 a 5 tem uma frase que diz o que ela significa. Duas pessoas pontuando a mesma atividade devem chegar perto.', 'Régua escrita.')
numbered(' Algumas situações não podem ficar em prioridade baixa, qualquer que seja a conta. Por isso há duas regras que passam por cima do cálculo.', 'Regras de veto.')
numbered(' Pesos, cortes e notas ficam em células editáveis na planilha. Mudou um peso, tudo recalcula.', 'Tudo editável.')
numbered(' A planilha mostra cada passo da conta. Dá para conferir qualquer resultado com uma calculadora.', 'Conta visível.')
para('O caminho completo de uma atividade é este:', after=3)

flow = ['5 notas\n(1 a 5)', 'Média\nponderada', '× Fator do\nprocesso', '+ Ajuste da\nnatureza', '= Nota\nfinal', 'Cortes', 'Regras de\nveto', 'Prioridade\n1 a 4']
ft = doc.add_table(rows=1, cols=len(flow)); ft.autofit = False; borders(ft, NAVY, 6)
for i, x in enumerate(flow):
    c = ft.rows[0].cells[i]; c.width = Cm(2.08)
    cell_write(c, x, 8.5, True, 'FFFFFF' if i in (0, 7) else NAVY, 'c'); shade(c, NAVY if i in (0, 7) else 'D9E1F2')
para('', after=2)
para('Esforço e prazo não entram nessa linha. Eles definem a ordem dentro da categoria (seção 7).', 9, italic=True, color='595959')

# =====================================================================
h1('3. Os 5 critérios, os pesos e a régua de cada nota')
para('Cada atividade recebe uma nota de 1 a 5 em cada critério. A média é ponderada pelos pesos abaixo, que somam 100%.')
table(['Critério', 'Pergunta que ele responde', 'Peso'],
      [(c['nome'], c['perg'], '%d%%' % c['peso']) for c in crit] + [('Soma', '', '100%')],
      [4.6, 9.6, 2.5], size=9.5, first_col_bold=True, center_cols=(2,), shade_first=True, caption='Tabela 1 · Critérios e pesos (aba Matriz)')

h2('Por que esses pesos')
bullet(' É o critério que mais separa o que pode esperar do que não pode. Perda por fraude ou erro grave não se recupera com facilidade, e o ponto de maior perda costuma ser pequeno, repetido e sem barreira. Por isso tem o maior peso.', 'Risco de fraude e erro (30%).')
bullet(' Mede o tamanho do dinheiro envolvido. Sozinho esconderia risco pequeno e repetido, então pesa menos que o risco.', 'Impacto financeiro (20%).')
bullet(' O negócio depende de entregar obra e de liberar o repasse da CEF. Atraso aqui trava caixa.', 'Obra e entrega (20%).')
bullet(' Tem consequência vinda de fora: multa, embargo, exigência de contrato ou da CEF. Pesa menos porque a maior parte das atividades não depende disso.', 'Exigência legal, CEF e compliance (15%).')
bullet(' Uma atividade que destrava várias outras vale mais do que uma isolada. É o critério que dá valor às definições de regra (alçada, responsável, cadastro).', 'Dependência (15%).')
para('Risco e exigência legal somam 45%. Isso é intencional: o trabalho de Processos inclui controle interno e prevenção a fraude, e esses são riscos que não se recuperam com facilidade.', after=4)

h2('A régua de cada nota')
head = ['Critério', 'Nota 1', 'Nota 2', 'Nota 3', 'Nota 4', 'Nota 5']
table(head, [[c['nome']] + c['esc'] for c in crit], [2.9, 2.6, 2.7, 2.9, 2.7, 2.9], size=8, first_col_bold=True, shade_first=True,
      caption='Tabela 2 · O que significa cada nota (aba Matriz)')
para('As faixas de R$ da linha de Impacto financeiro são exemplos. O Financeiro precisa confirmar os valores reais.', 9, italic=True, color='595959')

# =====================================================================
h1('4. O peso do processo')
para('Cada processo recebe três notas de 1 a 5. A média das três vira um fator que multiplica a nota da atividade.')
bullet(' quanto do dinheiro da empresa passa por este processo.', 'Dinheiro:')
bullet(' se ele parar, quantos outros processos param junto.', 'Dependência:')
bullet(' se o erro tem volta ou se o dano fica (comprar o terreno errado não se desfaz).', 'Irreversibilidade:')
para('Conta do fator: fator = %s + (%s − %s) × (média − 1) ÷ 4. Com isso, média 1 dá fator %s e média 5 dá fator %s.' % (fnum(fmin, 1), fnum(fmax, 1), fnum(fmin, 1), fnum(fmin, 1), fnum(fmax, 1)))
table(['Processo', 'Dinheiro', 'Depend.', 'Irrevers.', 'Média', 'Fator', 'Por que (hipótese)'],
      [[p[0], p[1], p[2], p[3], fnum(p[4]), fnum(p[5], 3), p[6]] for p in procs], [3.2, 1.7, 1.8, 1.8, 1.3, 1.3, 5.6], size=8.5,
      first_col_bold=True, center_cols=(1, 2, 3, 4, 5), shade_first=True, caption='Tabela 3 · Notas e fator por processo (aba Processos)')
rich([('Por que a faixa é pequena (%s a %s): ' % (fnum(fmin, 1), fnum(fmax, 1)), True),
      ('no macroprocesso, Financeiro e Suprimentos aparecem como apoio. Mas é no Financeiro que o dinheiro sai. Se o fator variasse muito, uma atividade de alto risco num processo de apoio ficaria escondida. Com a faixa pequena, o processo apenas desempata. Quem manda é o risco da atividade.', False)])

# =====================================================================
h1('5. A natureza da ação')
para('A natureza diz o que a atividade faz, e isso muda o valor dela. O ajuste é somado à nota depois do fator do processo.')
table(['Natureza', 'Quando usar', 'Ajuste'],
      [[n[0], n[1], ('%+.1f' % n[2]).replace('.', ',') if n[2] else '0,0'] for n in nats], [3.8, 11.0, 1.9], size=9, first_col_bold=True, center_cols=(2,), shade_first=True,
      caption='Tabela 4 · Ajuste por natureza (aba Matriz)')
para('Por que controle sobe: alçada, trava, segregação e conciliação são o que impede a fraude de acontecer. Por que indicador desce: medir sem proteger não reduz o risco por si só. Se o indicador detecta desvio (como "avanço físico x avanço financeiro"), ele deve ser classificado como Controle. A coluna Natureza da planilha vem de uma sugestão automática minha e deve ser revisada.')

# =====================================================================
h1('6. Cortes e regras que passam por cima da conta')
table(['Categoria', 'Quando se aplica', 'Nota mínima'],
      [[c[0], c[1], fnum(c[2], 1)] for c in cuts], [3.3, 10.4, 3.0], size=9.5, first_col_bold=True, center_cols=(2,), shade_first=True,
      caption='Tabela 5 · Cortes das categorias (aba Matriz)')
para('Regra: a nota final cai na maior categoria cujo mínimo ela alcança. Exemplo: 3,66 é maior que 3,3 e menor que 4,0, então é categoria 3 · Alta.')
h2('Regras que passam por cima da conta (vetos)')
para('Alguns casos não podem ficar em prioridade baixa, mesmo que a média fique baixa. Duas regras elevam a categoria para um mínimo:')
table(['Regra', 'Quando vale', 'Categoria mínima'],
      [['Risco no nível máximo', 'A nota de "Risco de fraude e erro" é 5.', '%d · Alta' % veto_risco],
       ['Para obra ou repasse da CEF', 'A atividade está marcada "Sim" em "Para obra ou repasse CEF?".', '%d · Urgente' % veto_obra]],
      [4.6, 8.8, 3.3], size=9.5, first_col_bold=True, center_cols=(2,), shade_first=True, caption='Tabela 6 · Vetos (aba Matriz)')
para('Por que existe: a média dilui. Uma atividade com risco 5 e nota 1 nos outros quatro critérios daria uma média baixa, e é justamente o tipo de atividade que não deve esperar. O veto impede isso. Segue a mesma ideia dos "critérios de veto" que aparecem no material de Protocolos do Grupo (versão antiga encontrada na pasta do Projeto Next).')

# =====================================================================
h1('7. Esforço e ordem de execução')
para('O esforço (1 a 5) não entra na nota. Ele serve para ordenar as atividades que ficaram na mesma categoria:')
reset_num()
numbered(' primeiro a categoria mais alta;', 'Primeiro:')
numbered(' dentro dela, a maior nota final;', 'Depois:')
numbered(' em caso de empate, o menor esforço (ganho rápido).', 'Por último:')
para('Escala de esforço: 1 = até 1 semana, 2 = até 1 mês, 3 = até 3 meses, 4 = mais de 3 meses, 5 = depende de contratação, sistema externo ou decisão da diretoria. O resultado aparece na coluna "Ordem geral" da planilha.')

# =====================================================================
h1('8. Um exemplo calculado passo a passo')
ex = anch[0]
nome_c = [c['nome'] for c in crit]
para('Atividade: ' + ex['nome'], 10, True, after=2)
para('Processo: %s · Natureza: %s' % (ex['proc'], ex['nat']), 9, italic=True, color='595959')
step_rows = []
tot = 0
for i, c in enumerate(crit):
    parte = ex['n'][i] * c['peso'] / 100
    tot += parte
    step_rows.append([c['nome'], ex['n'][i], '%d%%' % c['peso'], fnum(parte)])
step_rows.append(['Média ponderada', '', '100%', fnum(tot)])
table(['Critério', 'Nota', 'Peso', 'Nota × peso'], step_rows, [8.0, 2.2, 2.4, 3.0], size=9.5, center_cols=(1, 2, 3), caption='Tabela 7 · Passo 1: média ponderada')
nota = ex['nota']
pas = [
    ('Passo 2.', ' Fator do processo %s = %s. Conta: %s × %s = %s.' % (ex['proc'], fnum(ex['fator'], 3), fnum(ex['media']), fnum(ex['fator'], 3), fnum(ex['media'] * ex['fator']))),
    ('Passo 3.', ' Ajuste da natureza (%s): %s. Conta: %s %s %s = %s.' % (ex['nat'], ('%+.1f' % ex['aj']).replace('.', ','), fnum(ex['media'] * ex['fator']), '+' if ex['aj'] >= 0 else '−', fnum(abs(ex['aj'])), fnum(nota))),
    ('Passo 4.', ' Nota final %s. O corte de Urgente é %s, então a categoria pela nota é %d.' % (fnum(nota), fnum(cuts[3][2], 1), ex['catn'])),
    ('Passo 5.', ' Vetos: o risco é %d, então o veto de risco máximo vale (mínimo Alta). Como a categoria já é %d, o veto não muda o resultado. A atividade não está marcada como "para obra ou repasse CEF".' % (ex['n'][0], ex['catn'])),
    ('Resultado:', ' categoria %d · %s.' % (ex['catf'], ex['prio'])),
]
for a, b in pas: rich([(a, True), (b, False)], after=3)

# =====================================================================
# SECAO PAISAGEM: AMOSTRAS
# =====================================================================
ls = doc.add_section(WD_SECTION.NEW_PAGE)
ls.orientation = WD_ORIENT.LANDSCAPE
ls.page_width, ls.page_height = Cm(29.7), Cm(21)
ls.left_margin = Cm(1.8); ls.right_margin = Cm(1.8); ls.top_margin = Cm(1.8); ls.bottom_margin = Cm(1.8)

h1('9. As 12 atividades de amostra')
para('Escolhi 12 atividades abertas do ClickUp para mostrar a matriz funcionando. Cobrem de propósito os dois extremos: o que não pode esperar e o que pode. Por isso o resultado delas não representa a distribuição de todas as 401. As notas dos critérios são minhas. Servem de âncora: se você discordar de uma nota, mude na planilha (aba Tarefas, células amarelas) e a categoria recalcula.')
rows = []
for i, a in enumerate(anch, 1):
    rows.append([i, (a['nome'] if len(a['nome']) < 70 else a['nome'][:67].rstrip() + '...', urls.get(a['id'])), a['proc'], a['nat'], a['esf']] + a['n'] +
                [fnum(a['media']), fnum(a['fator'], 3), ('%+.1f' % a['aj']).replace('.', ',') if a['aj'] else '0,0', fnum(a['nota']), a['catn'], a['catf'], a['prio']])
hdr = ['#', 'Atividade (link no ClickUp)', 'Processo', 'Natureza', 'Esf.', 'Risco', 'Fin.', 'Obra', 'Legal', 'Dep.', 'Média pond.', 'Fator', 'Ajuste', 'Nota final', 'Cat. pela nota', 'Cat. final', 'Prioridade']
wd = [0.7, 5.2, 2.3, 2.2, 1.0, 1.2, 1.0, 1.1, 1.2, 1.1, 1.3, 1.2, 1.3, 1.2, 1.3, 1.1, 1.7]
t = table(hdr, rows, wd, size=8, head_size=7.5, center_cols=tuple([0] + list(range(4, 17))), caption='Tabela 8 · As 12 atividades de amostra, com a conta de cada uma, em ordem de execução (aba Tarefas)')
# cor da prioridade
colors = {'Urgente': 'F4B6B6', 'Alta': 'F9CB9C', 'Normal': 'FFE599', 'Baixa': 'B6D7A8'}
for row in t.rows[1:]:
    c = row.cells[16]; txt = c.text.strip()
    if txt in colors: shade(c, colors[txt])
para('Como ler: Média pond. = média ponderada dos 5 critérios. Nota final = Média pond. × Fator + Ajuste. Cat. pela nota = categoria pelos cortes. Cat. final = maior entre a categoria pela nota e os vetos. Esf. = esforço de 1 a 5.', 8.5, italic=True, color='595959')

h2('Por que cada atividade recebeu essas notas')
for i, a in enumerate(anch, 1):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4); p.paragraph_format.keep_together = True
    R(p, '%d. ' % i, 9.5, True)
    add_link(p, urls.get(a['id'], ''), a['nome'], 9.5)
    R(p, ' (%s · %s). ' % (a['proc'], a['status']), 9.5, italic=True, color='595959')
    R(p, 'Notas: risco %d, financeiro %d, obra %d, legal %d, dependência %d; esforço %d. ' % tuple(a['n'] + [a['esf']]), 9.5)
    R(p, 'Por quê: ', 9.5, True)
    R(p, a['why'] + ' ', 9.5)
    R(p, 'Resultado: %s (nota %s).' % (a['prio'], fnum(a['nota'])), 9.5, True)

para('Duas leituras que o exemplo mostra:', 10, True, after=2)
bullet(' o POP de pagamentos manuais tem nota 3,26, que daria Normal. Como o risco é 5, o veto sobe para Alta. É o caso em que a média dilui um risco grande.', 'O veto funcionou:')
pc = [a for a in anch if 'Política Comercial' in a['nome']][0]
bullet(' o controle da Política Comercial tem nota %s, que é Alta porque o ajuste de Controle (+0,3) a levou ao corte de %s. Sem o ajuste a nota seria %s, Normal. É o caso em que a natureza da ação decide a categoria.' % (fnum(pc['nota']), fnum(cuts[2][2], 1), fnum(pc['media'] * pc['fator'])), 'A natureza mudou o resultado:')

# =====================================================================
# VOLTA PARA RETRATO
# =====================================================================
ps = doc.add_section(WD_SECTION.NEW_PAGE)
ps.orientation = WD_ORIENT.PORTRAIT
ps.page_width, ps.page_height = Cm(21), Cm(29.7)
ps.left_margin = Cm(2.5); ps.right_margin = Cm(1.8); ps.top_margin = Cm(1.8); ps.bottom_margin = Cm(1.8)

h1('10. O que é hipótese e como validar')
para('Nada na lista abaixo veio de dado oficial. Cada item mostra a origem e a forma de validar.')
table(['Item', 'Origem', 'Como validar'],
      [['Pesos dos 5 critérios', 'Minha proposta, pelo raciocínio da seção 3.', 'Comparação par a par: com 5 critérios são 10 perguntas ("entre risco e prazo, qual pesa mais e quanto?"). Chefe, Controladoria e Processos respondem separados e a planilha calcula os pesos.'],
       ['Cortes das categorias', 'Minha proposta.', 'Pontuar todas as abertas e olhar a distribuição. Meta sugerida: Urgente até 10%, Alta até 25%. Se passar, subir o corte.'],
       ['Notas de processo e fator', 'Minha leitura do macroprocesso, dos fluxos e das dores.', 'Gestores validam as 3 notas do próprio processo. Trocar por dado real: VGV por empreendimento e % do custo em Suprimentos e mão de obra.'],
       ['Faixas de R$ do impacto financeiro', 'Exemplos meus.', 'Financeiro define as faixas reais.'],
       ['Notas das 12 âncoras', 'Minha leitura do nome e do contexto da atividade.', 'Responsável e gestor da área contestam. Se uma âncora mudar muito, ajustar a régua da seção 3.'],
       ['Ajustes de natureza', 'Minha proposta.', 'Testar tirando o ajuste: se muita atividade mudar de categoria, o ajuste pesa demais.'],
       ['Natureza de cada atividade', 'Sugestão automática por tipo de ação e palavras do nome.', 'Revisão por amostra: 30 atividades, conferir quantas estão erradas.']],
      [4.0, 4.6, 8.1], size=9, first_col_bold=True, shade_first=True, caption='Tabela 9 · Hipóteses e validação')
para('Teste final recomendado: usar a matriz nas 401 atividades e olhar onde ela discorda do que a equipe sente que é prioridade. Nos pontos de discordância, uma das duas está errada, e é ali que se aprende mais.')

h1('11. Decisões que dependem de você')
reset_num()
for t in ['A matriz serve para ordenar a entrega (o que fazer primeiro) ou para medir exposição (onde a empresa está mais vulnerável)? A segunda dá mais peso a risco e compliance.',
          'Os pesos dos 5 critérios estão bons, ou algum deve subir ou descer?',
          'Os dois vetos fazem sentido? Falta algum? Exemplo possível: atividade exigida por norma de segurança do trabalho.',
          'Quem pontua: Processos propõe e o gestor da área contesta, ou outra regra?',
          'Quem decide os pesos finais: você, o chefe, a diretoria ou o gestor de cada área?',
          'A prioridade final deve ser gravada no campo Prioridade do ClickUp, ou ficar só na planilha?',
          'Gadens e Make It: uma nota única por ação, com a marca da empresa, ou notas separadas?']:
    numbered(t)

# =====================================================================
h1('Anexos')
h2('A. Glossário')
table(['Termo', 'Significado neste material'],
      [['Alçada', 'Limite de valor ou de tipo de decisão que cada cargo pode aprovar.'],
       ['Segregação de funções', 'Regra de que a mesma pessoa não faz tudo no mesmo fluxo (por exemplo, pedir, aprovar e pagar).'],
       ['Âncora', 'Atividade já pontuada, usada como referência para calibrar as demais.'],
       ['Veto', 'Regra que eleva a categoria a um mínimo, qualquer que seja a nota calculada.'],
       ['Natureza da ação', 'Tipo de coisa que a atividade faz: controle, documento normativo, automação, definição, treinamento ou indicador.'],
       ['Irreversibilidade', 'Se o erro tem volta. Comprar um terreno errado não se desfaz; corrigir um texto de campanha, sim.'],
       ['Média ponderada', 'Média em que cada nota é multiplicada pelo peso do critério.'],
       ['RACM', 'Matriz de riscos e controles: lista cada risco e o controle que o cobre.']],
      [4.0, 12.7], size=9, first_col_bold=True, shade_first=True)
h2('B. Mapa da planilha')
para('O arquivo "Matriz de Prioridades - ClickUp Projeto Next.xlsx" espelha este material. Cada tabela daqui tem a aba correspondente:')
table(['Aba da planilha', 'O que tem', 'Seção deste material'],
      [['Leia-me', 'Como usar, cores e o que é proposta.', '1 e 2'],
       ['Matriz', 'Critérios, pesos, régua das notas, cortes, vetos e ajustes de natureza.', '3, 5 e 6'],
       ['Processos', 'Notas e fator de cada processo.', '4'],
       ['Resumo', 'Quantidade por categoria e por processo, com meta de distribuição.', '10'],
       ['Tarefas', 'As 401 atividades abertas com link, as 12 âncoras pontuadas e as fórmulas.', '7, 8 e 9']],
      [3.5, 9.4, 3.8], size=9, first_col_bold=True, shade_first=True)

# rodape com numero de pagina (todas as secoes)
def footer(section):
    section.footer.is_linked_to_previous = False
    p = section.footer.paragraphs[0]; p.text = ''; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    R(p, 'Projeto Next · Matriz de prioridades · material para validação · página ', 8, color='7F7F7F')
    r = p.add_run(); set_fonts(r._r.get_or_add_rPr(), 8); r.font.color.rgb = RGBColor.from_string('7F7F7F')
    for typ, txt in (('begin', None), (None, ' PAGE '), ('end', None)):
        if typ:
            fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), typ); r._r.append(fc)
        else:
            it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = txt; r._r.append(it)
for s in doc.sections: footer(s)

doc.save(OUT)
print('salvo', OUT)
