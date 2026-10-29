from PIL import Image, ImageDraw, ImageFont
import sys
import os
F=os.environ.get('FONT_DIR', os.path.expanduser('~/.fonts/'))
light=lambda s: ImageFont.truetype(F+'Manrope_wght_300.ttf', s)
mono=lambda s: ImageFont.truetype(F+'Geist_Mono_wght_400.ttf', s)
W,H=1920,1080
cols=[('Branco mineral','#FAF8F4','34%','#1C211E','#59625A'),('Pedra líquen','#D8DACD','32%','#1C211E','#59625A'),
      ('Preto musgo','#1C211E','27%','#FAF8F4','#D8DACD'),('Cinza ardósia verde','#59625A','6%','#FAF8F4','#D8DACD'),
      ('Verde vibrante','#DBFF00','0,8%','#1C211E','#59625A'),('Rosa choque','#FF19B1','0,2%','#FAF8F4','#FAF8F4')]
props=[34,32,27,6,0.8,0.2]
im=Image.new('RGB',(W,H)); d=ImageDraw.Draw(im)
def block(x,y,w,h,c,big,nsize=44,psize=130):
    n,bg,p,fg,sub=c
    d.rectangle([x,y,x+w,y+h],fill=bg)
    pad=48 if w>400 else 36
    d.text((x+pad,y+62),n,font=light(nsize),fill=fg)
    d.text((x+pad,y+62+nsize+30),bg,font=mono(20),fill=sub)
    d.text((x+pad-4,y+h-pad-10),p,font=light(psize),fill=fg,anchor='ls')
TH=600
w1=round(W*34/66)
block(0,0,w1,TH,cols[0],True); block(w1,0,W-w1,TH,cols[1],True)
d.text((w1-48,70),'PALETA MEOV · 2026',font=mono(18),fill='#59625A',anchor='ra')
ws=[800,500,340,280]; x=0
for c,w,ns,ps in zip(cols[2:],ws,[44,44,34,30],[130,130,76,64]):
    block(x,TH,w,H-TH,c,False,ns,ps); x+=w
# faixa com a proporção real
y0=H-10; x=0
for c,p in zip(cols,props):
    w=W*p/100; d.rectangle([x,y0,x+w,H],fill=c[1]); x+=w
im.save(sys.argv[1])
