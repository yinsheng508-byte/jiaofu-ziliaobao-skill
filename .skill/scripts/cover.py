"""封面：项目里的 封面/封面.html 用 Chrome 无界面导出成 A4 PDF，拼到完整版 PDF 最前面。
页码标签：封面标「封面」，正文从 1 起，阅读器里显示的页码和目录、页脚一致。
用法：.venv/bin/python cover.py <封面.html> <完整版.pdf>   （需要 pypdf，装在 .skill/.venv）
"""
import os, shutil, subprocess, sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def find_chrome():   # 统一用 export_pdf.find_chrome（Mac / Windows，Chrome 没有用 Edge）
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import export_pdf
    return export_pdf.find_chrome() or ("chrome" if sys.platform == "win32" else "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")


CHROME = find_chrome()


def render(html: Path) -> Path:
    out = html.with_suffix(".pdf")
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={out}", html.resolve().as_uri()],
                   check=True, capture_output=True)
    if len(PdfReader(out).pages) != 1:
        raise SystemExit(f"封面导出成了 {len(PdfReader(out).pages)} 页，检查 封面.html 的尺寸")
    return out


def attach(cover: Path, book: Path):
    w = PdfWriter()
    w.append(PdfReader(cover))
    w.append(PdfReader(book))
    n = len(w.pages)
    w.set_page_label(0, 0, prefix="封面")
    w.set_page_label(1, n - 1, style="/D", start=1)
    tmp = book.with_suffix(".tmp.pdf")
    with open(tmp, "wb") as f:
        w.write(f)
    tmp.replace(book)
    return n


if __name__ == "__main__":
    html, book = Path(sys.argv[1]), Path(sys.argv[2])
    print(f"  封面已加到 {book.name}，共 {attach(render(html), book)} 页")
