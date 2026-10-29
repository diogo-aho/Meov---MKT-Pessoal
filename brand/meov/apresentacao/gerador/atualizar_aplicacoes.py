"""Atualiza os mockups da identidade de 2023 para a Meov Marca Pessoal.

Renderiza as páginas de aplicação do PDF de 2023 e troca só os textos antigos
(@meovelab, "experiências audiovisuais", "produtora de vídeos", "KM design" etc.).
O fundo de cada trecho é refeito linha a linha a partir de uma faixa vizinha do
mesmo mockup, então nada é esticado nem distorcido.

Uso: python3 atualizar_aplicacoes.py <pasta_saida>
Requer pdftoppm, Pillow, numpy e as fontes Manrope instaladas.
"""
import os, subprocess, sys, tempfile, glob
import numpy as np
from PIL import Image, ImageDraw, ImageFont

AQUI = os.path.dirname(os.path.abspath(__file__))
PDF = os.path.join(AQUI, "..", "meov-identidade-visual-2023.pdf")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(AQUI, "img", "aplicacoes")
FONTS = os.path.expanduser("~/.fonts")
MINERAL, MUSGO, VERDE, CINZA = (250, 248, 244), (28, 33, 30), (219, 255, 0), (138, 138, 138)

def font(peso, size):
    return ImageFont.truetype(os.path.join(FONTS, f"Manrope-{peso}.ttf"), size)

def refill(im, box, band, noise=2.5, seed=1):
    """Pinta box (x0,y0,x1,y1) linha a linha com a mediana da faixa band (x0,x1) da mesma linha."""
    a = np.array(im).astype(float)
    x0, y0, x1, y1 = box
    rng = np.random.default_rng(seed)
    for y in range(y0, y1):
        c = np.median(a[y, band[0]:band[1]], axis=0)
        a[y, x0:x1] = c + rng.normal(0, noise, (x1 - x0, 3))
    return Image.fromarray(np.clip(a, 0, 255).astype("uint8"))

def vfill(im, box, pad=3, band=6, noise=None, seed=2):
    """Refaz box interpolando, coluna a coluna, entre a faixa logo acima e a faixa logo abaixo."""
    a = np.array(im).astype(float)
    x0, y0, x1, y1 = box
    top = np.median(a[y0 - pad - band:y0 - pad, x0:x1], axis=0)
    bot = np.median(a[y1 + pad:y1 + pad + band, x0:x1], axis=0)
    if noise is None:
        noise = float(np.std(a[y0 - pad - band:y0 - pad, x0:x1] - top, axis=(0, 1)).mean())
    t = np.linspace(0, 1, y1 - y0)[:, None, None]
    rng = np.random.default_rng(seed)
    a[y0:y1, x0:x1] = top[None] * (1 - t) + bot[None] * t + rng.normal(0, max(noise, 0.5), (y1 - y0, x1 - x0, 3))
    return Image.fromarray(np.clip(a, 0, 255).astype("uint8"))

def flat(im, box, color):
    ImageDraw.Draw(im).rectangle(box, fill=color)
    return im

def text(im, xy, s, size, color, peso=300, spacing=1.15, anchor="la", angle=0):
    f = font(peso, size)
    if not angle:
        d = ImageDraw.Draw(im)
        y = xy[1]
        for line in s.split("\n"):
            d.text((xy[0], y), line, font=f, fill=color, anchor=anchor)
            y += int(size * spacing)
        return im
    l, t, r, b = f.getbbox(s)
    layer = Image.new("RGBA", (r + 8, b + 8), (0, 0, 0, 0))
    ImageDraw.Draw(layer).text((4, 4), s, font=f, fill=color + (255,))
    layer = layer.rotate(-angle, resample=Image.BICUBIC, expand=True)
    im = im.convert("RGBA")
    im.alpha_composite(layer, (xy[0] - layer.width // 2, xy[1] - layer.height // 2))
    return im.convert("RGB")

def rot_patch(im, center, size, angle, color, feather=4, noise=2.0, seed=3):
    """Cobre um trecho inclinado com a cor do fundo, com bordas suaves e granulado."""
    from PIL import ImageFilter
    w, h = size
    pad = feather * 3
    mask = Image.new("L", (w + 2 * pad, h + 2 * pad), 0)
    ImageDraw.Draw(mask).rectangle((pad, pad, pad + w, pad + h), fill=255)
    mask = mask.rotate(-angle, resample=Image.BICUBIC, expand=True).filter(ImageFilter.GaussianBlur(feather))
    rng = np.random.default_rng(seed)
    c0, c1 = (color, color) if isinstance(color[0], int) else color
    t = np.linspace(0, 1, mask.width)[None, :, None]
    base = np.array(c0, float)[None, None] * (1 - t) + np.array(c1, float)[None, None] * t
    fill = np.clip(base + rng.normal(0, noise, (mask.height, mask.width, 3)), 0, 255).astype("uint8")
    layer = Image.fromarray(fill).convert("RGBA")
    layer.putalpha(mask)
    im = im.convert("RGBA")
    im.alpha_composite(layer, (center[0] - layer.width // 2, center[1] - layer.height // 2))
    return im.convert("RGB")

def avg(im, box):
    a = np.array(im.crop(box)).reshape(-1, 3)
    return tuple(int(v) for v in np.median(a, axis=0))

EDITS = {
    # Instagram
    22: lambda im: text(text(text(
        refill(refill(refill(im, (598, 180, 836, 230), (900, 1000), 1), (600, 506, 1260, 546), (1270, 1330), 1), (600, 586, 850, 624), (1270, 1330), 1),
        (604, 184), "meov.mp", 36, MINERAL, 500), (604, 508), "marca pessoal para especialistas.", 29, MINERAL, 400), (604, 590), "meov.com.br", 29, (150, 150, 150), 400),
    # Stories: tela do meio
    23: lambda im: text(text(
        vfill(vfill(im, (860, 194, 1060, 224)), (856, 498, 1104, 582)),
        (960, 198), "marca pessoal", 15, (200, 200, 200), 400, anchor="ma"), (960, 508), "sua voz, no\nritmo certo", 27, MINERAL, 400, anchor="ma"),
    # Navegador: aba
    24: lambda im: text(refill(im, (750, 186, 1296, 244), (1300, 1345), 1), (758, 193), "meov · marca pessoal", 40, MINERAL, 400),
    # Notebook: título do site
    25: lambda im: text(vfill(im, (736, 400, 858, 516)), (742, 408), "sua voz,\nno ritmo\ncerto.", 21, MINERAL, 300, 1.2),
    # LinkedIn (banner rosa)
    26: lambda im: text(text(
        flat(vfill(im, (1536, 312, 1690, 424)), (212, 724, 640, 762), (255, 255, 255)),
        (1676, 330), "sua voz,\nno ritmo certo.", 21, VERDE, 400, 1.2, anchor="ra"), (216, 728), "Marca pessoal • São Paulo, SP", 21, CINZA, 400),
    # LinkedIn (banner com play)
    27: lambda im: text(text(
        flat(vfill(im, (1536, 312, 1690, 424)), (212, 724, 640, 762), (255, 255, 255)),
        (1676, 330), "sua voz,\nno ritmo certo.", 21, VERDE, 400, 1.2, anchor="ra"), (216, 728), "Marca pessoal • São Paulo, SP", 21, CINZA, 400),
    # Cartaz
    28: lambda im: text(text(text(
        vfill(vfill(vfill(im, (708, 210, 868, 266)), (1034, 210, 1148, 266)), (708, 858, 818, 892)),
        (714, 216), "sua voz, no\nritmo certo.", 18, MINERAL, 400, 1.2), (1040, 216), "marca\npessoal", 18, VERDE, 400, 1.2), (714, 864), "@meov.mp", 18, MINERAL, 400),
    # Papelaria
    30: lambda im: text(text(
        vfill(vfill(im, (364, 962, 490, 996)), (934, 678, 1038, 712)),
        (370, 968), "@meov.mp", 18, (90, 90, 90), 400), (940, 684), "@meov.mp", 18, MUSGO, 400),
    # Indoor
    36: lambda im: text(text(
        vfill(vfill(im, (1048, 440, 1252, 572)), (1274, 770, 1364, 802)),
        (1056, 448), "sua voz,\nno ritmo\ncerto.", 28, MINERAL, 300, 1.2), (1356, 776), "@meov.mp", 15, MINERAL, 400, anchor="ra"),
    # Conceito do símbolo
    14: lambda im: text(vfill(im, (1196, 580, 1364, 622)), (1202, 584), "voz", 30, VERDE, 300),
}

def tablet(im):
    ang = 19.5
    im = rot_patch(im, (596, 418), (172, 40), ang, (avg(im, (508, 392, 532, 420)), avg(im, (672, 440, 700, 466))))
    im = rot_patch(im, (1262, 332), (96, 26), ang, avg(im, (1190, 350, 1220, 372)), feather=3)
    im = text(im, (600, 420), "Meov", 26, MINERAL, 400, angle=ang)
    im = text(im, (1262, 332), "@meov.mp", 16, MINERAL, 400, angle=ang)
    return im
EDITS[32] = tablet

PAGINAS = [6, 7, 14, 16, 17] + list(range(22, 48))

def main():
    os.makedirs(OUT, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        for p in PAGINAS:
            subprocess.run(["pdftoppm", "-png", "-r", "300", "-f", str(p), "-l", str(p), PDF, os.path.join(tmp, "p")], check=True, capture_output=True)
            src = glob.glob(os.path.join(tmp, "p-*.png"))[0]
            im = Image.open(src).convert("RGB")
            os.remove(src)
            if p in EDITS: im = EDITS[p](im)
            im.save(os.path.join(OUT, f"pagina-{p:02d}.jpg"), quality=88)
    print("ok", len(PAGINAS))

if __name__ == "__main__":
    main()
