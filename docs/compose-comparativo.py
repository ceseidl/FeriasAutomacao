"""
Monta uma imagem side-by-side antes (v1.0.0) / depois (v1.0.0.1) a partir
de 2 screenshots gerados do relatorio Ferias.html. Util pra apresentar pra
chefia o que mudou entre as duas versoes.

Uso:
    python docs/compose-comparativo.py
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
RESULTS = ROOT / 'results'

ANTES  = RESULTS / 'antes-1.0.0.png'
DEPOIS = RESULTS / 'depois-1.0.0.1.png'
OUT    = RESULTS / 'comparativo-1.0.0-vs-1.0.0.1.png'

# Identidade visual (mesma da apresentacao executiva)
PRIMARY = (0x1F, 0x49, 0x7D)
ACCENT  = (0x2E, 0x75, 0xB6)
WHITE   = (0xFF, 0xFF, 0xFF)
INK     = (0x26, 0x26, 0x26)
MUTED   = (0x66, 0x66, 0x66)
SUCCESS = (0x1E, 0x7E, 0x34)

GAP        = 30          # espaco horizontal entre as duas imagens
PADDING    = 30          # margem externa
HEADER_H   = 110         # faixa superior com titulos
LABEL_H    = 70          # caixa de label (antes/depois) acima de cada imagem


def load_font(size, bold=False):
    """Tenta carregar Arial Bold/Regular do Windows. Fallback default."""
    candidates = [
        r'C:\Windows\Fonts\arialbd.ttf' if bold else r'C:\Windows\Fonts\arial.ttf',
        r'C:\Windows\Fonts\arial.ttf',
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


antes  = Image.open(ANTES).convert('RGB')
depois = Image.open(DEPOIS).convert('RGB')

# Iguala alturas (escolhe a maior, redimensiona a menor mantendo proporcao)
target_h = max(antes.height, depois.height)
def fit_height(im, h):
    if im.height == h:
        return im
    new_w = int(im.width * h / im.height)
    return im.resize((new_w, h), Image.LANCZOS)

antes  = fit_height(antes, target_h)
depois = fit_height(depois, target_h)

# Dimensoes finais
inner_w  = antes.width + GAP + depois.width
total_w  = inner_w + 2 * PADDING
total_h  = HEADER_H + LABEL_H + target_h + PADDING

canvas = Image.new('RGB', (total_w, total_h), WHITE)
draw   = ImageDraw.Draw(canvas)

# ============================================================
# Faixa superior corporativa
# ============================================================
draw.rectangle([(0, 0), (total_w, HEADER_H)], fill=PRIMARY)
draw.rectangle([(0, HEADER_H - 4), (total_w, HEADER_H)], fill=ACCENT)

font_h1 = load_font(36, bold=True)
font_h2 = load_font(18, bold=False)
draw.text((PADDING, 22), 'Ferias Automacao  --  Comparativo de versoes',
          fill=WHITE, font=font_h1)
draw.text((PADDING, 70),
          'v1.0.0 (em producao)  ->  v1.0.0.1 (proxima release)',
          fill=WHITE, font=font_h2)

# ============================================================
# Labels acima de cada imagem
# ============================================================
font_label   = load_font(20, bold=True)
font_sublbl  = load_font(13, bold=False)

label_y = HEADER_H + 8
# antes
left_x = PADDING
draw.rectangle(
    [(left_x, label_y), (left_x + antes.width, label_y + LABEL_H - 14)],
    fill=(0xF2, 0xF2, 0xF2),
    outline=(0xBF, 0xBF, 0xBF), width=1,
)
draw.text((left_x + 16, label_y + 8),
          'ANTES  --  v1.0.0', fill=PRIMARY, font=font_label)
draw.text((left_x + 16, label_y + 32),
          'Em producao hoje. Sem secao de vencimento.',
          fill=MUTED, font=font_sublbl)

# depois
right_x = left_x + antes.width + GAP
draw.rectangle(
    [(right_x, label_y), (right_x + depois.width, label_y + LABEL_H - 14)],
    fill=(0xE7, 0xF3, 0xE7),
    outline=SUCCESS, width=2,
)
draw.text((right_x + 16, label_y + 8),
          'DEPOIS  --  v1.0.0.1', fill=SUCCESS, font=font_label)
draw.text((right_x + 16, label_y + 32),
          'Nova secao "Controle de Vencimento de Ferias" (CLT 12/24).',
          fill=MUTED, font=font_sublbl)

# ============================================================
# Imagens
# ============================================================
img_y = HEADER_H + LABEL_H
canvas.paste(antes, (left_x, img_y))
canvas.paste(depois, (right_x, img_y))

# Borda discreta nas imagens
draw.rectangle([(left_x, img_y),
                (left_x + antes.width, img_y + antes.height)],
               outline=(0xBF, 0xBF, 0xBF), width=1)
draw.rectangle([(right_x, img_y),
                (right_x + depois.width, img_y + depois.height)],
               outline=SUCCESS, width=2)

# ============================================================
# Salva
# ============================================================
canvas.save(OUT, 'PNG', optimize=True)
size_kb = OUT.stat().st_size // 1024
print(f'OK: {OUT}  ({total_w}x{total_h}px, {size_kb} KB)')
