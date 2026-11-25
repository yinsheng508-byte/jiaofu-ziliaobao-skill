"""按 手工裁图.tsv（PDF页 x0 y0 x1 y1 图名 [抹白框]，页面比例）裁图到 插图/M_图名.png
PDF页写 S02:1 / S04:2 表示参考资料第几页（200dpi 页图）"""
from pathlib import Path
from PIL import Image, ImageDraw
H = Path(__file__).resolve().parent
(H / "插图").mkdir(exist_ok=True)
for l in (H / "手工裁图.tsv").read_text(encoding="utf-8").splitlines():
    if not l.strip() or l.startswith("#"): continue
    cols = l.split("\t"); p, x0, y0, x1, y1, name = cols[:6]; er = cols[6] if len(cols) > 6 else ""
    if p.startswith("S02:"):
        im = Image.open(H / "S02页图" / f"q-{int(p[4:]):02d}.png")
    elif p.startswith("S04:"):
        im = Image.open(H / "S04页图" / f"q-{int(p[4:])}.png")
    else:
        im = Image.open(H / "页图300" / f"p-{int(p):03d}.png")
    im = im.convert("RGB"); W, Hh = im.size
    if er:
        d = ImageDraw.Draw(im)
        for b in er.split(";"):
            a = [float(v) for v in b.split(",")]
            col = im.getpixel((max(0, int(a[0] * W) - 4), min(Hh - 1, int(a[3] * Hh) + 4))) if len(a) > 4 else "white"
            d.rectangle([a[0] * W, a[1] * Hh, a[2] * W, a[3] * Hh], fill=col)
    c = [max(0.0, min(1.0, float(v))) for v in (x0, y0, x1, y1)]
    im.crop((int(c[0] * W), int(c[1] * Hh), int(c[2] * W), int(c[3] * Hh))).save(H / "插图" / f"M_{name}.png")
