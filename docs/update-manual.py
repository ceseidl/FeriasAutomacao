"""
Atualiza docs/MANUAL.docx in-place para refletir as mudancas do v1.0.0:
  1. Reescreve o paragrafo do controle "Ano" pra descrever o year picker
     estilo calendario (substituiu o NumericUpDown).
  2. Adiciona uma regra explicita sobre o filtro estrito por ano.
  3. Substitui a screenshot embedada por docs/screenshot.png atualizado.
  4. Corrige o rodape que ainda fala em MANUAL.md como fonte (a fonte
     virou o proprio .docx).
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

doc.save(str(DOCX))
print(f'OK: {DOCX} atualizado.')
