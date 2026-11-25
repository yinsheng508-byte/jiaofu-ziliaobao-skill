"""手写笔记合订：按顺序出每一章（页码全书连续、答案放章末），再合成一本可打印的小册子。

  python3 run.py notebook <项目>/手写笔记 --name <书名> [--all]

- 手写稿按文件名排序：00_封面目录.md、01_….md、02_….md；文件名带「_答案」的不算
- 每章头部自动写「起始页码」和「答案: 章末」，前面的章页数变了后面自动顺延
- 只重出改过的章（稿子比 PDF 新、或起始页码变了）；--all 全部重出
- 成品：3_成品/手写笔记/<稿名>/（每章 PNG + PDF + 检查报告）和 3_成品/手写笔记/<书名>_合订本.pdf
- 页数变了要同步改封面目录稿里的目录页码
"""
import argparse
import logging
import re
import subprocess
import sys
from pathlib import Path

logging.getLogger("pypdf").setLevel(logging.ERROR)

HERE = Path(__file__).resolve().parent
NOTE = HERE / "handnote.py"


def chapters(d):
    return sorted(p for p in Path(d).glob("[0-9][0-9]_*.md") if "_答案" not in p.stem)


def set_meta(md, start):
    t = md.read_text(encoding="utf-8")
    m = re.match(r"^(---\n.*?\n---\n)(.*)$", t, re.S)
    if not m:
        raise SystemExit(f"{md.name} 没有 --- 文件头")
    head, body = m.groups()
    old = head
    head = re.sub(r"^显示答案:.*\n", "", head, flags=re.M)
    for k, v in (("起始页码", str(start)), ("答案", "章末")):
        if re.search(rf"^{k}:", head, re.M):
            head = re.sub(rf"^{k}:.*$", f"{k}: {v}", head, flags=re.M)
        else:
            head = head[:-4] + f"{k}: {v}\n---\n"
    if head != old:
        md.write_text(head + body, encoding="utf-8")
        return True
    return False


def render(md, out_root):
    out = out_root / md.stem
    for old in out.glob("p-*.png"):          # 页数变少时别留下旧页
        old.unlink()
    r = subprocess.run([sys.executable, str(NOTE), str(md), "--out", str(out)], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode:
        print(r.stdout[-2000:], r.stderr[-2000:])
        raise SystemExit(f"出图失败：{md.name}")
    rep = (out / "检查报告.md").read_text(encoding="utf-8")
    fill = re.search(r"## 各页填充率\n(.+)", rep)
    print(f"  {md.stem}：{len(list(out.glob('p-*.png')))} 页　填充率 {fill.group(1) if fill else '—'}")


def merge(pdfs, book):
    try:
        from pypdf import PdfWriter, PdfReader
    except ImportError:
        venv = HERE.parent / ".venv" / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
        code = ("import sys\nfrom pypdf import PdfWriter, PdfReader\nw = PdfWriter()\n"
                "[w.append(PdfReader(p)) for p in sys.argv[2:]]\nw.write(sys.argv[1])")
        subprocess.run([str(venv), "-c", code, str(book), *map(str, pdfs)], check=True)
        return
    w = PdfWriter()
    for p in pdfs:
        w.append(PdfReader(str(p)))
    w.write(str(book))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dir")
    ap.add_argument("--name", required=True)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--out", help="成品放哪（默认：新项目放 3_成品/手写笔记/，老项目放 手写笔记/成品/）")
    a = ap.parse_args()
    d = Path(a.dir)
    sys.path.insert(0, str(HERE))
    import paths
    P = paths.of(d)
    out_root = Path(a.out) if a.out else (P.手写成品 if not P.legacy and d.resolve() == P.手写笔记.resolve() else d / "成品")
    start, pdfs = 1, []
    for md in chapters(d):
        changed = set_meta(md, start)
        pdf = out_root / md.stem / f"{md.stem}.pdf"
        if a.all or changed or not pdf.exists() or md.stat().st_mtime > pdf.stat().st_mtime:
            render(md, out_root)
        else:
            print(f"  {md.stem}：没改，跳过")
        pdfs.append(pdf)
        start += len(list((out_root / md.stem).glob("p-*.png")))
    book = out_root / f"{a.name}_合订本.pdf"
    merge(pdfs, book)
    print(f"合订本：{book}（共 {start - 1} 页）")


if __name__ == "__main__":
    main()
