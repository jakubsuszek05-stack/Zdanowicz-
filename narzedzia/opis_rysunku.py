from PIL import Image, ImageDraw, ImageFont
S=2
im=Image.open('zdjecia/toyota_zacisk.jpg').convert('RGB')
im=im.resize((im.width*S,im.height*S),Image.LANCZOS)
d=ImageDraw.Draw(im)
f=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',44)
R=34
# (numer, pozycja kółka, [cele])
L=[(1,(880,470),[(725,480)]),
   (2,(880,330),[(760,322)]),
   (3,(300,560),[(500,545)]),
   (4,(290,230),[(468,205)]),
   (4,(340,450),[(445,488)]),
   (5,(700,90),[(606,222)]),
   (5,(410,400),[(566,456)])]
for n,(cx,cy),ts in L:
    for tx,ty in ts:
        d.line([(cx*S,cy*S),(tx*S,ty*S)],fill=(200,0,0),width=5)
        d.ellipse([(tx*S-7,ty*S-7),(tx*S+7,ty*S+7)],fill=(200,0,0))
for n,(cx,cy),ts in L:
    d.ellipse([(cx*S-R*S/2*1.6,cy*S-R*S/2*1.6),(cx*S+R*S/2*1.6,cy*S+R*S/2*1.6)],fill='white',outline=(200,0,0),width=5)
    d.text((cx*S,cy*S),str(n),font=f,fill=(200,0,0),anchor='mm')
im.save('zdjecia/toyota_zacisk_opis.jpg',quality=92)
