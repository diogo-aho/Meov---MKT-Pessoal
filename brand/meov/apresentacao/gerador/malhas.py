"""Gera as malhas de construção do símbolo e do logotipo da Meov a partir dos SVGs."""
import re, sys, math, os
BASE = os.path.join(os.path.dirname(__file__), "..", "..", "logo")
OUT = sys.argv[1] if len(sys.argv) > 1 else "."

def paths(svg, cls="12"):
    return re.findall(r'<path class="fil[' + cls + r']" d="([^"]+)"', svg)

def segments(d):
    toks = re.findall(r'[MmLlCcZz]|-?\d+(?:\.\d+)?', d)
    i, x, y, cmd, segs, pts = 0, 0.0, 0.0, None, [], []
    def num():
        nonlocal i; v = float(toks[i]); i += 1; return v
    while i < len(toks):
        t = toks[i]
        if re.match(r'[A-Za-z]', t):
            cmd = t; i += 1
            if cmd in 'Zz': continue
        if cmd == 'M': x, y = num(), num(); cmd = 'L'
        elif cmd == 'm': x, y = x + num(), y + num(); cmd = 'l'
        elif cmd in 'Ll':
            nx, ny = num(), num()
            if cmd == 'l': nx, ny = x + nx, y + ny
            segs.append(((x, y), (nx, ny))); x, y = nx, ny
        elif cmd == 'c':
            vals = [num() for _ in range(6)]; x, y = x + vals[4], y + vals[5]
        elif cmd == 'C':
            vals = [num() for _ in range(6)]; x, y = vals[4], vals[5]
        pts.append((x, y))
    return segs, pts

def build(kind, src, fg, bg, guide, out):
    svg = open(os.path.join(BASE, src)).read()
    ps = paths(svg)
    allpts, diag = [], []
    for d in paths(svg, "1"):
        segs, pts = segments(d); allpts += pts
        for (a, b) in segs:
            if math.hypot(b[0]-a[0], b[1]-a[1]) > 700 and abs(b[0]-a[0]) > 100 and abs(b[1]-a[1]) > 100: diag.append((a, b))
    xs = [p[0] for p in allpts]; ys = [p[1] for p in allpts]
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    w, h = x1 - x0, y1 - y0
    vw = w * 1.9; vh = vw * 9 / 16
    vx = x0 - (vw - w) / 2; vy = y0 - (vh - h) / 2
    L = []
    def line(xa, ya, xb, yb, dash=False):
        L.append(f'<line x1="{xa:.1f}" y1="{ya:.1f}" x2="{xb:.1f}" y2="{yb:.1f}" stroke="#{guide}" stroke-width="{vw/1400:.1f}"' + (f' stroke-dasharray="{vw/200:.0f} {vw/200:.0f}"' if dash else '') + '/>')
    for yy in (y0, y1): line(vx, yy, vx + vw, yy)
    for xx in (x0, x1): line(xx, vy, xx, vy + vh)
    seen = set()
    for (a, b) in diag:
        dx, dy = b[0]-a[0], b[1]-a[1]
        k = (round(a[0] - a[1]*dx/dy), round(dx/dy, 2))
        if any(abs(k[0]-s[0]) < 40 and k[1] == s[1] for s in seen): continue
        seen.add(k)
        t0, t1 = (vy - a[1]) / dy, (vy + vh - a[1]) / dy
        line(a[0] + dx*t0, vy, a[0] + dx*t1, vy + vh)
    if kind == 'logotipo':
        mid = (y0 + y1) / 2; line(vx, mid, vx + vw, mid, dash=True)
    body = "\n".join(f'<path fill="#{fg}" fill-rule="evenodd" d="{d}"/>' for d in ps)
    out_svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="{vx:.0f} {vy:.0f} {vw:.0f} {vh:.0f}" width="1920" height="1080">
<rect x="{vx:.0f}" y="{vy:.0f}" width="{vw:.0f}" height="{vh:.0f}" fill="#{bg}"/>
{chr(10).join(L)}
{body}
</svg>'''
    open(out, 'w').write(out_svg)

os.makedirs(OUT, exist_ok=True)
build('simbolo', 'simbolo/meov-simbolo-musgo-sem-fundo.svg', 'FAF8F4', '1C211E', '59625A', os.path.join(OUT, 'malha-simbolo.svg'))
build('logotipo', 'logotipo/meov-logotipo-musgo-sem-fundo.svg', 'FAF8F4', '1C211E', '59625A', os.path.join(OUT, 'malha-logotipo.svg'))
print("ok")
