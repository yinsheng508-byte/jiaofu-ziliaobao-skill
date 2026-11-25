"""PDF → 逐页预览图 + 拼版总览，供 Agent 用 Read 逐页检查排版。

  python3 render.py <文件.pdf> <输出目录> [--dpi 70] [--pages 2-4]

输出：p-01.png …（单页）和 总览-1.png …（每张 6 页，3×2）。
"""
import argparse
from pathlib import Path

from export_pdf import pdf_to_png


def render(pdf, out, dpi=70, pages=None):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    for f in out.glob("*.png"):
        f.unlink()
    a = b = None
    if pages:
        a, _, b = pages.partition("-")
        b = b or a
    pdf_to_png(pdf, out / "p", dpi, a, b)
    files = sorted(out.glob("p-*.png"))
    try:
        from PIL import Image
    except ImportError:
        return files, []
    sheets = []
    ims = [Image.open(f) for f in files]
    if not ims:
        return files, sheets
    w, h = ims[0].size
    for k in range(0, len(ims), 6):
        sheet = Image.new("RGB", (w * 3, h * 2), "white")
        for i, im in enumerate(ims[k:k + 6]):
            sheet.paste(im, ((i % 3) * w, (i // 3) * h))
        path = out / f"总览-{k // 6 + 1}.png"
        sheet.save(path)
        sheets.append(path)
    return files, sheets


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out")
    ap.add_argument("--dpi", type=int, default=70)
    ap.add_argument("--pages")
    a = ap.parse_args()
    files, sheets = render(a.pdf, a.out, a.dpi, a.pages)
    print(f"{len(files)} 页")
    for s in sheets:
        print("  ", s)
