"""
Atualiza docs/MANUAL.docx in-place para refletir as mudancas do v1.0.0:
  1. Reescreve o paragrafo do controle "Ano" pra descrever o year picker
     estilo calendario (substituiu o NumericUpDown).
  2. Adiciona uma regra explicita sobre o filtro estrito por ano.
  3. Substitui a screenshot embedada por docs/screenshot.png atualizado.
  4. Corrige o rodape que ainda fala em MANUAL.md como fonte (a fonte
     virou o proprio .docx).
  5. Adiciona instrucoes de instalacao via MSI (com descricao do dialogo
     de selecao de atalhos) e renumera as subsecoes 1.x.

Idempotente: rodar mais de uma vez nao duplica conteudo.
"""
import sys
from pathlib import Path
from docx import Document

ROOT = Path(__file__).resolve().parent.parent
DOCX = ROOT / 'docs' / 'MANUAL.docx'
NEW_PNG = ROOT / 'docs' / 'screenshot.png'

doc = Document(str(DOCX))

# ---------- 1) Paragrafo do controle "Ano" ----------
target_old = 'Ano - ano que aparece no titulo do relatorio e no Gantt'
new_text = (
    'Ano - ano que aparece no titulo do relatorio, no Gantt e que filtra '
    'a planilha. Padrao = ano atual. Clicar no botao "ANO ▾" abre um '
    'calendario 4x3 com 12 anos por pagina; usar as setas no topo do '
    'popup para navegar de decada em decada e clicar no ano desejado. '
    'O popup fecha sozinho ao escolher um ano, ao apertar Esc ou ao '
    'clicar fora dele.'
)
hits = 0
for p in doc.paragraphs:
    if p.text.startswith(target_old):
        # zera os runs e reescreve preservando o estilo
        for r in list(p.runs):
            r.text = ''
        if p.runs:
            p.runs[0].text = new_text
        else:
            p.add_run(new_text)
        hits += 1
        break
if hits == 0:
    print('AVISO: paragrafo do controle Ano nao encontrado', file=sys.stderr)

# ---------- 2) Regra do filtro estrito por ano ----------
# Insere apos o paragrafo "2.3 Regras" como mais um item Compact.
regra_filtro = (
    'O relatorio so traz ferias do ano selecionado no picker. Mesmo que '
    'a planilha tenha linhas de outros anos, elas sao filtradas fora '
    'antes de gerar o relatorio.'
)
inserted = False
for i, p in enumerate(doc.paragraphs):
    if p.text.strip() == 'Salvar e fechar a planilha antes de gerar o relatorio (Excel trava o arquivo)':
        # Insere um novo paragrafo logo apos, com o mesmo estilo "Compact".
        new_p = p.insert_paragraph_before('')  # placeholder so pra ter o objeto
        # insert_paragraph_before insere ANTES; entao trocamos a abordagem:
        # cria via XML clonando o atual.
        from copy import deepcopy
        clone = deepcopy(p._element)
        # zera o texto do clone
        for t in clone.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            t.text = regra_filtro
            break
        # remove os outros <w:t> alem do primeiro pra nao concatenar lixo
        ts = list(clone.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
        for extra in ts[1:]:
            extra.getparent().remove(extra)
        p._element.addnext(clone)
        # remove o placeholder vazio que criei sem querer
        new_p._element.getparent().remove(new_p._element)
        inserted = True
        break
if not inserted:
    print('AVISO: regra do filtro nao foi inserida', file=sys.stderr)

# ---------- 3) Rodape ----------
for p in doc.paragraphs:
    if p.text.startswith('Este manual e gerado a cada release a partir de docs/MANUAL.md'):
        for r in list(p.runs):
            r.text = ''
        if p.runs:
            p.runs[0].text = (
                'Este manual e mantido em docs/MANUAL.docx (fonte direta). '
                'Para atualizar, edite o .docx e rode docs/update-manual.py se '
                'precisar reaplicar mudancas em massa.'
            )
        break

# ---------- 4) Substitui a screenshot embedada ----------
# python-docx nao tem API publica pra trocar a imagem; manipulo a parte do pacote.
new_bytes = NEW_PNG.read_bytes()
swapped = False
for rel_id, rel in doc.part.rels.items():
    target = rel.target_ref
    if 'media' in target and target.endswith('.png'):
        # rel.target_part eh o ImagePart
        rel.target_part._blob = new_bytes
        swapped = True
        print(f'Screenshot substituida em {target} ({len(new_bytes)} bytes)')
if not swapped:
    print('AVISO: nenhuma imagem encontrada pra substituir', file=sys.stderr)

# ---------- 5) Instalacao via MSI ----------
# Plano:
#   - Reescreve o paragrafo "Clicar em FeriasAutomacao-latest.zip..." pra
#     listar as DUAS opcoes (MSI recomendado + ZIP portatil).
#   - Insere uma nova secao "1.2 Instalar via MSI (recomendado)" antes
#     da atual "1.2 Extrair o ZIP", com paragrafos Compact descrevendo
#     o fluxo do instalador (incluindo a tela de selecao de atalhos).
#   - Renumera "1.2 Extrair o ZIP" -> "1.3 Alternativa: extrair o ZIP"
#     e "1.3 Conferir o conteudo" -> "1.4 Conferir o conteudo".
#
# Idempotente: se o heading "1.2 Instalar via MSI" ja existir, pula
# tudo; se os renumerados ja existirem, nao reescreve.

from copy import deepcopy
W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def set_text(p, txt):
    """Reescreve um paragrafo preservando o estilo do primeiro run."""
    for r in list(p.runs):
        r.text = ''
    if p.runs:
        p.runs[0].text = txt
    else:
        p.add_run(txt)

def clone_with_text(template_p, txt):
    """Clona um paragrafo (preservando o estilo dele) e poe o texto novo."""
    clone = deepcopy(template_p._element)
    ts = list(clone.iter(f'{W_NS}t'))
    if ts:
        ts[0].text = txt
        for extra in ts[1:]:
            extra.getparent().remove(extra)
    else:
        # se nao tem <w:t>, cria um run+text minimo dentro do paragrafo
        from docx.oxml.ns import qn
        r = clone.makeelement(qn('w:r'), {})
        t = clone.makeelement(qn('w:t'), {})
        t.text = txt
        r.append(t)
        clone.append(r)
    return clone

# ja foi feito antes? se sim, sai cedo
already_done = any(p.text.strip() == '1.2 Instalar via MSI (recomendado)'
                   for p in doc.paragraphs)

if not already_done:
    # 5.1) reescreve o "Clicar em FeriasAutomacao-latest.zip..."
    for p in doc.paragraphs:
        if p.text.strip().startswith('Clicar em FeriasAutomacao-latest.zip'):
            set_text(p, (
                'Baixar uma das opcoes disponiveis na pagina (recomendado: '
                'FeriasAutomacao-1.0.0.msi para instalacao automatica; '
                'FeriasAutomacao-latest.zip para versao portatil sem '
                'instalar nada).'
            ))
            break

    # 5.2) localiza o "1.2 Extrair o ZIP" pra usar como ponto de insercao
    p_extrair = None
    p_conferir = None
    for p in doc.paragraphs:
        if p.text.strip() == '1.2 Extrair o ZIP':
            p_extrair = p
        elif p.text.strip() == '1.3 Conferir o conteudo':
            p_conferir = p

    if p_extrair is None:
        print('AVISO: heading "1.2 Extrair o ZIP" nao encontrado, nao da pra inserir MSI', file=sys.stderr)
    else:
        # pega um Compact ja existente como template de estilo
        compact_template = None
        for p in doc.paragraphs:
            if p.style and p.style.name == 'Compact' and p.text.strip():
                compact_template = p
                break
        if compact_template is None:
            print('AVISO: nao achei nenhum paragrafo Compact pra usar como template', file=sys.stderr)

        # 5.3) renumera "1.2 Extrair o ZIP" -> "1.3 Alternativa: ..."
        set_text(p_extrair, '1.3 Alternativa: extrair o ZIP')

        # 5.4) renumera "1.3 Conferir o conteudo" -> "1.4 ..."
        if p_conferir is not None:
            set_text(p_conferir, '1.4 Conferir o conteudo')

        # 5.5) insere a nova secao ANTES de p_extrair.
        # ordem de insercao: heading + N paragrafos Compact + paragrafo de
        # transicao. Tudo via .addprevious() pra manter a ordem natural.
        msi_blocks = [
            ('Heading 3', '1.2 Instalar via MSI (recomendado)'),
            ('First Paragraph',
             'O .msi instala o app por usuario (nao precisa de admin) em '
             '%LocalAppData%\\Programs\\FeriasAutomacao e cria os atalhos '
             'no menu Iniciar e na area de trabalho automaticamente.'),
            ('Compact', 'Dar duplo clique em FeriasAutomacao-1.0.0.msi.'),
            ('Compact',
             'O instalador mostra uma tela de selecao de features tipo '
             'arvore. Por padrao tudo vem marcado:'),
            ('Compact',
             '- Aplicativo Ferias Automacao (obrigatorio): copia a planilha, '
             'os scripts, o icone e o instalador offline do Pandoc.'),
            ('Compact',
             '- Atalhos no Menu Iniciar (recomendado): cria a pasta '
             '"Programas > Ferias Automacao" com os atalhos Gerar Relatorio, '
             'Manual e Pasta de instalacao.'),
            ('Compact',
             '- Atalho na Area de Trabalho (recomendado): cria o atalho '
             '"Gerar Relatorio" direto na area de trabalho.'),
            ('Compact',
             '- Atalho em "Enviar para" (opcional, vem desmarcado): '
             'adiciona o app ao menu de clique-direito do Windows. '
             'Marcar so se quiser abrir uma planilha .xlsx direto pelo '
             'Enviar para.'),
            ('Compact',
             'Clicar em Avancar > Instalar e aguardar. Ao terminar, abrir '
             'o atalho Gerar Relatorio (menu Iniciar ou desktop) e seguir '
             'para a secao 2.'),
            ('First Paragraph',
             'Quem usar o MSI pode pular as secoes 1.3 e 1.4 abaixo - elas '
             'sao so para quem prefere a versao portatil em ZIP.'),
        ]

        # acha um paragrafo template pra cada estilo usado
        templates = {}
        for p in doc.paragraphs:
            if p.style and p.text.strip():
                name = p.style.name
                if name in {'Heading 3', 'First Paragraph', 'Compact'} and name not in templates:
                    templates[name] = p
        # checa que temos os 3
        missing = {'Heading 3', 'First Paragraph', 'Compact'} - templates.keys()
        if missing:
            print(f'AVISO: estilos nao encontrados: {missing}', file=sys.stderr)

        for style, txt in msi_blocks:
            tpl = templates.get(style)
            if tpl is None:
                continue
            new_p = clone_with_text(tpl, txt)
            p_extrair._element.addprevious(new_p)

        print('OK: secao "1.2 Instalar via MSI (recomendado)" inserida.')
else:
    print('Secao MSI ja presente; nada a fazer em (5).')

# ---------- 6.1) Atualiza descricao da aba Ferias na secao 2.1 ----------
# Antes mencionava apenas "uma aba chamada Ferias". Agora a planilha tem
# 5 abas (Ferias + Squads + Integrantes + Status + Instrucoes).
old_aba = ('A planilha deve ter uma aba chamada Ferias '
           '(e dela que o app le os dados).')
new_aba = ('A planilha tem 5 abas: Ferias (a principal, onde o app le os dados), '
           'Squads (lista de squads), Integrantes (lista de pessoas com seu squad), '
           'Status (Aprovada/Planejada/Solicitada) e Instrucoes (texto de ajuda). '
           'As 3 abas de listas alimentam os dropdowns das colunas Colaborador, '
           'Squad e Status da aba Ferias - pra adicionar uma pessoa nova ou um '
           'squad novo, edite a aba correspondente.')
for p in doc.paragraphs:
    if p.text.strip() == old_aba:
        for r in list(p.runs):
            r.text = ''
        if p.runs:
            p.runs[0].text = new_aba
        else:
            p.add_run(new_aba)
        break

# ---------- 6.2) Renomeacao da planilha (ferias-2026.xlsx -> Ferias-template.xlsx) ----------
# Substitui o nome em todos os runs de todos os paragrafos. Idempotente: se
# o nome ja foi trocado, nao acha mais e nao faz nada.
old_name = 'ferias-2026.xlsx'
new_name = 'Ferias-template.xlsx'
swaps = 0
for p in doc.paragraphs:
    for r in p.runs:
        if old_name in r.text:
            r.text = r.text.replace(old_name, new_name)
            swaps += 1
print(f'Renomeacao da planilha: {swaps} run(s) atualizados.')

# ---------- 7) Saida em pasta com timestamp (results\Ferias-{ts}\) ----------
# Antes os arquivos saiam soltos em results\. Agora cada execucao tem a
# sua subpasta results\Ferias-{ts}\ com .md, .html e .xlsx (snapshot da
# planilha) dentro. Atualiza:
#   - paragrafo descritivo "Cada execucao cria arquivos com timestamp..."
#   - bloco Source Code com a estrutura de pastas
#   - referencia em "Subir o arquivo Ferias-{timestamp}.pdf"
old_desc = 'Cada execucao cria arquivos com timestamp na pasta results\\:'
new_desc = ('Cada execucao cria uma subpasta results\\Ferias-{timestamp}\\ '
            'com o relatorio em multiplos formatos + a planilha-fonte:')
for p in doc.paragraphs:
    if p.text.strip() == old_desc:
        for r in list(p.runs):
            r.text = ''
        if p.runs:
            p.runs[0].text = new_desc
        else:
            p.add_run(new_desc)
        break

# Bloco Source Code (a estrutura literal). Reescrevemos pra refletir o
# novo layout. Procura pelo paragrafo Source Code que comeca com 'results\'
# e ainda mostra arquivos soltos sem subpasta.
new_tree = (
    'results\\\n'
    '  Ferias-20260425-143012\\\n'
    '    Ferias-20260425-143012.md     <- versao Markdown (texto puro)\n'
    '    Ferias-20260425-143012.html   <- versao HTML formatada (interativa)\n'
    '    Ferias-20260425-143012.xlsx   <- copia da planilha que gerou o relatorio\n'
    '    Ferias-20260425-143012.pdf    <- (opcional, com a opcao "Gerar PDF tambem" marcada)\n'
    '  Ferias-20260425-150133\\         <- proxima execucao, em nova subpasta\n'
    '    ...'
)
for p in doc.paragraphs:
    if p.style.name == 'Source Code' and p.text.lstrip().startswith('results\\') \
            and 'Ferias-20260425-143012\\' not in p.text:
        # zera todos os runs e poe o novo conteudo no primeiro
        for r in list(p.runs):
            r.text = ''
        if p.runs:
            p.runs[0].text = new_tree
        else:
            p.add_run(new_tree)
        break

# "Subir o arquivo Ferias-{timestamp}.pdf na biblioteca do SharePoint"
old_share = 'Subir o arquivo Ferias-{timestamp}.pdf na biblioteca do SharePoint'
new_share = ('Subir o arquivo Ferias-{timestamp}.pdf de dentro da subpasta '
             'Ferias-{timestamp}\\ na biblioteca do SharePoint')
for p in doc.paragraphs:
    if p.text.strip() == old_share:
        for r in list(p.runs):
            r.text = ''
        if p.runs:
            p.runs[0].text = new_share
        else:
            p.add_run(new_share)
        break

doc.save(str(DOCX))
print(f'OK: {DOCX} atualizado.')
