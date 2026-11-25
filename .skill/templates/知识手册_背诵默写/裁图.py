"""从课本页图里自动裁出插图，按图注命名。
  python3 -I 排版/裁图.py           （在项目目录运行）
依据：pdftotext -bbox-layout 的 bbox.html（文字块坐标，pt）+ 页图300/p-NNN.png
做法：正文字（行高 ≥11pt、非图注）抹白 → 剩下的墨迹连通成块 → 每个「图x.x」图注配最近的块
输出：课本插图_全部/p<课本页>_<图号>_<图名>.png（没配上图注的块：p<课本页>_块<k>.png）+ 清单.tsv
"""
import re
from collections import deque
from pathlib import Path
from PIL import Image

HERE = Path(__file__).resolve().parent
OUT = HERE / "课本插图_全部"
OUT.mkdir(exist_ok=True)
R = 300 / 72
html = (HERE / "bbox.html").read_text(encoding="utf-8")
pages = re.split(r"<page ", html)[1:]
W_PT, H_PT = 521.575, 737.008
G = 1.0  # 网格 1pt


def lines_of(pg):
    out = []
    for ln in re.findall(r"<line (.*?)</line>", pg, re.S):
        ws = re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', ln)
        if not ws:
            continue
        ws = [(float(a), float(b), float(c), float(d), t) for a, b, c, d, t in ws]
        h = sum(w[3] - w[1] for w in ws) / len(ws)
        out.append(dict(x0=min(w[0] for w in ws), y0=min(w[1] for w in ws), x1=max(w[2] for w in ws),
                        y1=max(w[3] for w in ws), h=h, t="".join(w[4] for w in ws)))
    return out


rows = []
for pi, pg in enumerate(pages, 1):
    book = pi - 5
    import sys
    if sys.argv[1:] and str(pi) not in sys.argv[1:]:
        continue
    if book < 1 or pi > 128:
        continue
    f = HERE / "页图300" / f"p-{pi:03d}.png"
    img = Image.open(f).convert("RGB")
    small = img.resize((int(W_PT / G), int(H_PT / G)), Image.BOX)
    W, H = small.size
    px = small.load()
    ink = bytearray(W * H)
    for y in range(H):
        for x in range(W):
            r, g, b = px[x, y]
            if 255 * 3 - (r + g + b) > 70 and min(r, g, b) < 225:
                ink[y * W + x] = 1
    ls = lines_of(pg)
    caps = [l for l in ls if re.match(r"^图\d+\.\d+", l["t"])]
    # 图注名：同一行右侧的小字
    for c in caps:
        tail = [l for l in ls if l is not c and abs(l["y0"] - c["y0"]) < 4 and 0 <= l["x0"] - c["x1"] < 30 and l["h"] < 12]
        c["name"] = c["t"] + "".join(l["t"] for l in sorted(tail, key=lambda l: l["x0"]))
        c["box"] = [c["x0"], c["y0"], max([c["x1"]] + [l["x1"] for l in tail]), c["y1"]]
    erase = [l for l in ls if l["h"] >= 11.0 and not re.match(r"^图\d", l["t"])] + [c for c in caps]
    for c in caps:
        erase.append(dict(x0=c["box"][0], y0=c["box"][1], x1=c["box"][2], y1=c["box"][3]))
    for l in erase:
        for y in range(max(0, int(l["y0"]) - 1), min(H, int(l["y1"]) + 2)):
            for x in range(max(0, int(l["x0"]) - 1), min(W, int(l["x1"]) + 2)):
                ink[y * W + x] = 0
    for y in range(int(H_PT - 62), H):       # 页脚
        for x in range(W):
            ink[y * W + x] = 0
    # 膨胀后连通
    D = 5
    seen = bytearray(W * H)
    comps = []
    pts = [i for i in range(W * H) if ink[i]]
    inkset = ink
    for s in pts:
        if seen[s]:
            continue
        q = deque([s]); seen[s] = 1
        x0 = x1 = s % W; y0 = y1 = s // W; n = 0
        while q:
            i = q.popleft(); n += 1
            x, y = i % W, i // W
            x0, x1, y0, y1 = min(x0, x), max(x1, x), min(y0, y), max(y1, y)
            for yy in range(max(0, y - D), min(H, y + D + 1)):
                base = yy * W
                for xx in range(max(0, x - D), min(W, x + D + 1)):
                    j = base + xx
                    if inkset[j] and not seen[j]:
                        seen[j] = 1; q.append(j)
        if (x1 - x0) >= 30 and (y1 - y0) >= 25 and n > 150:
            comps.append([x0, y0, x1, y1])
    used = set()
    for c in caps:
        cx = (c["box"][0] + c["box"][2]) / 2
        best, bd = None, 1e9
        for k, b in enumerate(comps):
            # 图在图注上方（最常见）或左右
            if b[1] < c["box"][1] - 10 and b[3] <= c["box"][3] + 30 and b[0] - 20 < cx < b[2] + 20:
                d = max(0, c["box"][1] - b[3])
            elif b[1] - 10 < c["box"][1] < b[3] + 10 and (abs(b[2] - c["box"][0]) < 40 or abs(c["box"][2] - b[0]) < 40):
                d = 15
            else:
                continue
            if d < bd:
                best, bd = k, d
        if best is None or bd > 80:
            rows.append(f"{book}\t{c['name']}\t（没配上）\t")
            continue
        used.add(best)
        x0, y0, x1, y1 = comps[best]
        crop = img.crop((int(x0 * R) - 6, int(y0 * R) - 6, int(x1 * R) + 6, int(y1 * R) + 6))
        m = re.match(r"^图(\d+\.\d+)(.*)$", c["name"])
        name = re.sub(r'[\\/:*?"<>|\s]', "", m.group(2))[:40] or "无名"
        fn = f"p{book:03d}_{m.group(1)}_{name}.png"
        crop.save(OUT / fn)
        rows.append(f"{book}\t{c['name']}\t{fn}\t{int(x1-x0)}x{int(y1-y0)}")
    for k, b in enumerate(comps):
        if k in used:
            continue
        x0, y0, x1, y1 = b
        crop = img.crop((int(x0 * R) - 6, int(y0 * R) - 6, int(x1 * R) + 6, int(y1 * R) + 6))
        fn = f"p{book:03d}_块{k}.png"
        crop.save(OUT / fn)
        rows.append(f"{book}\t（无图注）\t{fn}\t{x1-x0}x{y1-y0}")
    print(pi, len(caps), len(comps), flush=True)
(OUT / "清单.tsv").write_text("课本页\t图注\t文件\t尺寸pt\n" + "\n".join(rows), encoding="utf-8")
