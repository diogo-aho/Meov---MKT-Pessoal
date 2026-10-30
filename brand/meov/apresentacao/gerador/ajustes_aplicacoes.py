"""Ajustes da R05 sobre os mockups já atualizados (img/aplicacoes).

- notebook.jpg: pagina-25 sem a frase "sua voz, no ritmo certo." na tela.
- linkedin.jpg: pagina-26 com um banner mais leve, sobre a voz e a imagem
  de cada pessoa ganhando o mundo, na paleta atual.

Uso: python3 ajustes_aplicacoes.py   (requer Pillow, numpy e as fontes Manrope e Geist Mono em ~/.fonts)
"""
import os, math
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageFilter

AQUI = os.path.dirname(os.path.abspath(__file__))
APL = os.path.join(AQUI, "img", "aplicacoes")
FONTS = os.path.expanduser("~/.fonts")
MINERAL, PEDRA, MUSGO, ARDOSIA, ROSA, VERDE = (250, 248, 244), (216, 218, 205), (28, 33, 30), (89, 98, 90), (255, 25, 177), (219, 255, 0)


def hfill(im, box, band=6, seed=4):
    """Refaz box interpolando, linha a linha, entre a faixa à esquerda e a faixa à direita."""
    a = np.array(im).astype(float)
    x0, y0, x1, y1 = box
    rng = np.random.default_rng(seed)
    t = np.linspace(0, 1, x1 - x0)[:, None]
    for y in range(y0, y1):
        l = np.median(a[y, x0 - band:x0], axis=0)
        r = np.median(a[y, x1:x1 + band], axis=0)
        a[y, x0:x1] = l[None] * (1 - t) + r[None] * t + rng.normal(0, 1.2, (x1 - x0, 3))
    return Image.fromarray(np.clip(a, 0, 255).astype("uint8"))


def notebook():
    im = Image.open(os.path.join(APL, "pagina-25.jpg")).convert("RGB")
    im = hfill(im, (736, 404, 830, 494))
    im.save(os.path.join(APL, "notebook.jpg"), quality=92)


def banner(w, h):
    """Banner claro: uma linha que nasce reta e vira a onda da Meov, levando a voz para as pessoas."""
    s = 2
    W, H = w * s, h * s
    a = np.zeros((H, W, 3))
    t = np.linspace(0, 1, W)[None, :, None]
    a[:] = np.array(PEDRA)[None, None] * (1 - t) + np.array(MINERAL)[None, None] * t
    a += np.random.default_rng(7).normal(0, 1.6, a.shape)
    im = Image.fromarray(np.clip(a, 0, 255).astype("uint8"))
    d = ImageDraw.Draw(im)
    # a onda: reta à esquerda, pulsando cada vez mais à direita
    y0, x_start, x_end = H * 0.56, W * 0.30, W * 0.97
    pts = []
    for i in range(1400):
        x = x_start + (x_end - x_start) * i / 1399
        u = (x - x_start) / (x_end - x_start)
        amp = H * 0.30 * (max(0, u - 0.08) ** 1.6)
        pts.append((x, y0 - amp * math.sin(u * 9.5 * math.pi)))
    d.line([(W * 0.02, y0)] + pts, fill=MUSGO, width=int(3 * s), joint="curve")
    # pessoas que a voz alcança
    rng = np.random.default_rng(3)
    placed = 0
    while placed < 9:
        k = placed
        x = W * (0.62 + 0.34 * rng.random())
        y = H * (0.10 + 0.80 * rng.random())
        if min(math.hypot(x - px_, y - py_) for px_, py_ in pts[::10]) < H * 0.12:
            continue
        placed += 1
        r = s * (5 + 4 * rng.random())
        c = VERDE if k == 5 else MUSGO
        d.ellipse((x - r, y - r, x + r, y + r), fill=c, outline=MUSGO if c == VERDE else None, width=s)
    # o detalhe rosa: onde a voz começa a pulsar
    px, py = pts[160]
    d.ellipse((px - 7 * s, py - 7 * s, px + 7 * s, py + 7 * s), fill=ROSA)
    f = ImageFont.truetype(os.path.join(FONTS, "Manrope-300.ttf"), 30 * s)
    fm = ImageFont.truetype(os.path.join(FONTS, "GeistMono-Regular.ttf"), 11 * s)
    d.text((W * 0.30, H * 0.17), "sua voz, no ritmo certo.", font=f, fill=MUSGO)
    d.text((W * 0.30, H * 0.17 + 44 * s), "MARCA PESSOAL PARA ESPECIALISTAS", font=fm, fill=ARDOSIA)
    return im.resize((w, h), Image.LANCZOS)


def linkedin():
    im = Image.open(os.path.join(APL, "pagina-26.jpg")).convert("RGB")
    orig = im.copy()
    x0, y0, x1, y1, r = 166, 92, 1753, 490, 34
    b = banner(x1 - x0, y1 - y0)
    mask = Image.new("L", b.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, b.width - 1, b.height + r), radius=r, fill=255)
    im.paste(b, (x0, y0), mask)
    # a foto do perfil volta por cima do banner, com a moldura branca
    ImageDraw.Draw(im).rectangle((215, 341, 520, 645), fill=(255, 255, 255))
    im.paste(orig.crop((220, 347, 516, 641)), (220, 347))
    im.save(os.path.join(APL, "linkedin.jpg"), quality=92)


if __name__ == "__main__":
    notebook()
    linkedin()
    print("ok")
