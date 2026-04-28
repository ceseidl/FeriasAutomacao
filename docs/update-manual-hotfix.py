"""
Hotfix do MANUAL.docx pra v1.0.0 (final, com hotfixes pos-release inicial).

Aplica:
  1. Limpa paragrafos duplicados de "O relatorio so traz ferias..."
     (bug de idempotencia em rodadas anteriores do update-manual.py).
  2. Atualiza secao 1.2 (Instalar via MSI):
     - Substitui descricao do FeatureTree pelo InstallDir (tela dedicada
       pra escolher pasta de instalacao).
     - Renomeia "Programas > Ferias Automacao" pra "AI R" (nova pasta no
       Menu Iniciar).
     - Adiciona descricao do atalho "Desinstalar Ferias Automacao".
     - Adiciona descricao do checkbox "Iniciar Ferias Automacao agora"
       no fim do instalador.
  3. Substitui "icone da palmeira" -> "logo da AI R" (apos troca do
     emoji pelo logo vetorial no relatorio).
  4. Adiciona aviso sobre single-instance lock (apenas uma janela por PC).
  5. Adiciona aviso sobre lock estrito ao template oficial (so aceita
     Ferias-template.xlsx com as 5 abas obrigatorias).
  6. Adiciona troubleshooting de reparo via msiexec /f.

Idempotente: rodar mais de uma vez nao duplica conteudo.

Uso:
    python docs/update-manual-hotfix.py
"""
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
DOCX = ROOT / 'docs' / 'MANUAL.docx'

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

doc = Document(str(DOCX))


def set_text(p, txt):
    """Reescreve um paragrafo preservando o estilo do primeiro run."""
    for r in list(p.runs):
        r.text = ''
    if p.runs:
        p.runs[0].text = txt
    else:
        p.add_run(txt)


def clone_with_text(template_p, txt):
    """Clona um paragrafo preservando estilo e poe o texto novo."""
    clone = deepcopy(template_p._element)
    ts = list(clone.iter(f'{W_NS}t'))
    if ts:
        ts[0].text = txt
        for extra in ts[1:]:
            extra.getparent().remove(extra)
    else:
        r = clone.makeelement(qn('w:r'), {})
        t = clone.makeelement(qn('w:t'), {})
        t.text = txt
        r.append(t)
        clone.append(r)
    return clone


def find_template(style_name):
    """Devolve o primeiro paragrafo nao-vazio com o estilo dado, pra clonar."""
    for p in doc.paragraphs:
        if p.style and p.style.name == style_name and p.text.strip():
            return p
    return None


# ============================================================
# 1) Limpa duplicados de "O relatorio so traz ferias..."
# ============================================================
target_dup = 'O relatorio so traz ferias do ano selecionado no picker'
encontrados = [p for p in doc.paragraphs if p.text.strip().startswith(target_dup)]
removidos = 0
for p in encontrados[1:]:  # mantem o primeiro, remove os outros
    p._element.getparent().remove(p._element)
    removidos += 1
print(f'(1) Removidos {removidos} paragrafos duplicados (de {len(encontrados)} encontrados).')


# ============================================================
# 2) Atualiza secao 1.2 (Instalar via MSI)
# ============================================================
substitutions_1_2 = {
    'O .msi instala o app por usuario (nao precisa de admin) em %LocalAppData%\\Programs\\FeriasAutomacao e cria os atalhos no menu Iniciar e na area de trabalho automaticamente.':
        'O .msi instala o app por usuario (sem prompt de UAC). A pasta padrao '
        'e %LocalAppData%\\Programs\\FeriasAutomacao, mas o instalador permite '
        'escolher outra pasta. Os atalhos do Menu Iniciar ficam na pasta "AI R" '
        '(submenu Iniciar > AI R) e na area de trabalho.',

    'O instalador mostra uma tela de selecao de features tipo arvore. Por padrao tudo vem marcado:':
        'Fluxo do instalador (5 telas): Welcome -> License Agreement -> Choose '
        'installation folder (botao "Change..." pra trocar a pasta) -> '
        'Verify Ready -> Finish.',

    '- Aplicativo Ferias Automacao (obrigatorio): copia a planilha, os scripts, o icone e o instalador offline do Pandoc.':
        'A pasta "AI R" no Menu Iniciar tem 4 atalhos: Gerar Relatorio, Manual '
        'do usuario, Pasta de instalacao e Desinstalar Ferias Automacao.',

    '- Atalhos no Menu Iniciar (recomendado): cria a pasta "Programas > Ferias Automacao" com os atalhos Gerar Relatorio, Manual e Pasta de instalacao.':
        'Atalho "Desinstalar Ferias Automacao": chama o assistente de remocao '
        'do Windows. Equivalente a abrir Configuracoes > Aplicativos e mandar '
        'desinstalar.',

    '- Atalho na Area de Trabalho (recomendado): cria o atalho "Gerar Relatorio" direto na area de trabalho.':
        'Na tela final do instalador, o checkbox "Iniciar Ferias Automacao '
        'agora" ja vem marcado. Se mantiver marcado e clicar Finish, o app '
        'abre direto. Desmarcar so se preferir abrir manualmente depois pelo '
        'atalho.',

    '- Atalho em "Enviar para" (opcional, vem desmarcado): adiciona o app ao menu de clique-direito do Windows. Marcar so se quiser abrir uma planilha .xlsx direto pelo Enviar para.':
        'Para reabrir depois: Menu Iniciar > AI R > Gerar Relatorio (ou o '
        'atalho na area de trabalho).',

    'Clicar em Avancar > Instalar e aguardar. Ao terminar, abrir o atalho Gerar Relatorio (menu Iniciar ou desktop) e seguir para a secao 2.':
        'Apos a instalacao, seguir para a secao 2 (preencher a planilha) e '
        '3 (gerar o relatorio).',
}

aplicados = 0
for p in doc.paragraphs:
    txt = p.text.strip()
    if txt in substitutions_1_2:
        set_text(p, substitutions_1_2[txt])
        aplicados += 1
print(f'(2) Atualizados {aplicados} paragrafos da secao 1.2 MSI.')


# ============================================================
# 3) Substitui "palmeira" por "logo da AI R"
# ============================================================
trocas_logo = 0
for p in doc.paragraphs:
    for r in p.runs:
        if 'palmeira' in r.text.lower():
            txt = r.text
            txt = txt.replace('icone da palmeira', 'logo da AI R')
            txt = txt.replace('icone de palmeira', 'logo da AI R')
            txt = txt.replace('palmeira', 'logo da AI R')
            r.text = txt
            trocas_logo += 1
print(f'(3) Substituidas {trocas_logo} mencoes a palmeira pelo logo AI R.')


# ============================================================
# 4) Single-instance lock
# ============================================================
ja_tem_single = any('Apenas uma janela do app por PC' in p.text for p in doc.paragraphs)
if not ja_tem_single:
    for p in doc.paragraphs:
        if p.text.strip().startswith('Duplo-clique em Gerar Relatorio'):
            compact_tpl = find_template('Compact')
            if compact_tpl is not None:
                txt = ('Apenas uma janela do app por PC: se ja tiver uma janela aberta e voce '
                       'der duplo-clique no atalho de novo, a janela existente volta pra frente '
                       'em vez de abrir outra. Mensagem mostrada: "Ferias Automacao ja esta em '
                       'execucao."')
                new_p = clone_with_text(compact_tpl, txt)
                p._element.addnext(new_p)
                print('(4) Paragrafo single-instance inserido.')
            break
else:
    print('(4) Single-instance ja documentado, pulando.')


# ============================================================
# 5) Lock estrito ao template oficial
# ============================================================
ja_tem_lock = any('so aceita a planilha-template oficial' in p.text for p in doc.paragraphs)
if not ja_tem_lock:
    for p in doc.paragraphs:
        if p.text.strip() == 'Abrir Ferias-template.xlsx no Excel.':
            block_tpl = find_template('Block Text')
            if block_tpl is not None:
                txt = ('Importante: o app so aceita a planilha-template oficial, com nome '
                       'exato Ferias-template.xlsx e as 5 abas obrigatorias (Ferias, Squads, '
                       'Integrantes, Status, Instrucoes). Renomear o arquivo ou mexer na '
                       'estrutura faz o app rejeitar com mensagem clara antes de gerar o '
                       'relatorio.')
                new_p = clone_with_text(block_tpl, txt)
                p._element.addnext(new_p)
                print('(5) Aviso de lock do template oficial inserido.')
            break
else:
    print('(5) Lock do template ja documentado, pulando.')


# ============================================================
# 6) Reparo MSI no troubleshooting (secao 7)
# ============================================================
ja_tem_reparo = any('msiexec /f' in p.text for p in doc.paragraphs)
if not ja_tem_reparo:
    p_problemas = None
    for p in doc.paragraphs:
        if p.text.strip() == '7. Problemas comuns':
            p_problemas = p
            break
    if p_problemas is not None:
        compact_tpl = find_template('Compact')
        if compact_tpl is not None:
            txt = ('App parou de abrir / arquivo da pasta de instalacao foi deletado: '
                   'rodar de novo o .msi do release (oferece "Repair") ou abrir um console '
                   'e executar msiexec /f "FeriasAutomacao-1.0.0.1.msi" pra reinstalar os '
                   'arquivos faltantes. Como ultima alternativa, desinstalar pelo Painel '
                   'de Controle e reinstalar.')
            new_p = clone_with_text(compact_tpl, txt)
            p_problemas._element.addnext(new_p)
            print('(6) Troubleshooting de reparo MSI inserido.')
else:
    print('(6) Reparo MSI ja documentado, pulando.')


# ============================================================
# 7) Remove mencao ao checkbox "Iniciar agora" (que foi removido do MSI)
# ============================================================
# O update anterior tinha mencionado um checkbox "Iniciar Ferias Automacao
# agora" no fim do instalador. Esse checkbox foi removido depois (4 tentativas
# de fazer ele disparar o app falharam). Agora o usuario abre manualmente
# pelo atalho do Desktop ou Menu Iniciar.
substituicoes_7 = {
    'Na tela final do instalador, o checkbox "Iniciar Ferias Automacao agora" ja vem marcado. Se mantiver marcado e clicar Finish, o app abre direto. Desmarcar so se preferir abrir manualmente depois pelo atalho.':
        'Na tela final, clicar Finish pra concluir a instalacao. Apos o '
        'instalador fechar, abrir o app pelo atalho "Gerar Relatorio" '
        '(Area de Trabalho -- sempre criado -- ou Menu Iniciar > AI R).',
}
trocas_7 = 0
for p in doc.paragraphs:
    txt = p.text.strip()
    if txt in substituicoes_7:
        set_text(p, substituicoes_7[txt])
        trocas_7 += 1
print(f'(7) Substituidas {trocas_7} mencoes ao checkbox removido.')


# ============================================================
# 8) Atualiza o nome do MSI nas referencias (release v1.0.0.1)
# ============================================================
# Substitui FeriasAutomacao-1.0.0.msi -> FeriasAutomacao-1.0.0.1.msi
# em todos os runs de todos os paragrafos. Idempotente: se ja foi
# trocado, nao acha mais.
old_msi = 'FeriasAutomacao-1.0.0.msi'
new_msi = 'FeriasAutomacao-1.0.0.1.msi'
trocas_8 = 0
for p in doc.paragraphs:
    for r in p.runs:
        if old_msi in r.text:
            r.text = r.text.replace(old_msi, new_msi)
            trocas_8 += 1
print(f'(8) Substituidos {trocas_8} run(s) com nome do MSI antigo.')


# ============================================================
# Salva
# ============================================================
doc.save(str(DOCX))
print(f'OK: {DOCX} atualizado.')
