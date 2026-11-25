"""python3 -I 排版/网格.py 页 页 …（两页一行）：课本 PDF 页加 5% 网格（标比例），拼成 网格/起-止.png，用来定裁图坐标"""
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
H = Path(__file__).resolve().parent
pages = [int(x) for x in sys.argv[1:]]; cols = 2
src = H / "页图100"
font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 13)
ims = []
for k in pages:
    im = Image.open(src / f"p-{k:03d}.png").convert("RGB"); W, Hh = im.size
    d = ImageDraw.Draw(im, "RGBA")
    for i in range(1, 20):
        x, y = W * i / 20, Hh * i / 20
        c = (255, 0, 0, 150) if i % 2 == 0 else (0, 120, 255, 70)
        d.line([(x, 0), (x, Hh)], fill=c, width=1); d.line([(0, y), (W, y)], fill=c, width=1)
        if i % 2 == 0:
            d.text((x + 2, 2), f"{i*5}", fill=(220, 0, 0), font=font); d.text((2, y + 1), f"{i*5}", fill=(220, 0, 0), font=font)
    d.text((W - 60, Hh - 20), f"P{k}", fill=(0, 0, 200), font=font)
    ims.append(im)
W, Hh = ims[0].size
rows = (len(ims) + cols - 1) // cols
c = Image.new("RGB", (cols * (W + 8), rows * (Hh + 8)), "#555")
for i, im in enumerate(ims):
    c.paste(im, ((i % cols) * (W + 8), (i // cols) * (Hh + 8)))
(H / "网格").mkdir(exist_ok=True)
out = H / "网格" / ("-".join(map(str, pages)) + ".png"); c.save(out); print(out)
