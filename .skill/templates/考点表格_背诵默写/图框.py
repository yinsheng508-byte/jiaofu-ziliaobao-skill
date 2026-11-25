"""列出课本页里的位图框（页面比例坐标），并在页图上标号：python3 -I 排版/图框.py 课本页...
输出 排版/图框/b<页>.png 与终端坐标（x0 y0 x1 y1）"""
import re, subprocess, sys, tempfile
from pathlib import Path
from PIL import Image, ImageDraw
H = Path(__file__).resolve().parent
pdf = H.parent / "素材/S01_课本.pdf"
for b in sys.argv[1:]:
    pi = int(b) + 5
    x2 = subprocess.run(["pdftohtml", "-xml", "-stdout", "-f", str(pi), "-l", str(pi), str(pdf)], capture_output=True, text=True, errors="ignore").stdout
    pg = re.search(r'<page[^>]*height="(\d+)" width="(\d+)"', x2); PH, PW = int(pg[1]), int(pg[2])
    boxes = [(int(l)/PW, int(t)/PH, (int(l)+int(w))/PW, (int(t)+int(h))/PH) for t, l, w, h in re.findall(r'<image top="(-?\d+)" left="(-?\d+)" width="(\d+)" height="(\d+)"', x2)]
    im = Image.open(H / "页图72" / f"p-{pi:03d}.png").convert("RGB"); W, Hh = im.size
    im = im.resize((W * 2, Hh * 2)); W, Hh = im.size; d = ImageDraw.Draw(im)
    for k, (a, bb, c, e) in enumerate(boxes):
        if (c - a) < 0.08: continue
        d.rectangle([a*W, bb*Hh, c*W, e*Hh], outline="red", width=3); d.text((a*W+4, bb*Hh+4), str(k), fill="red")
        print(b, k, f"{a:.3f} {bb:.3f} {c:.3f} {e:.3f}")
    im.save(H / "图框" / f"b{b}.png")
