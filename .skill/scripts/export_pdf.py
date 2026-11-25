"""Word 母稿 → PDF，并提供目录页码查找。

  python3 export_pdf.py <文件.docx> [--engine word|wps|lo]

engine：
  word  用 Microsoft Word 导出（AppleScript）。成品首选：和用户在 Word 里看到的一致
  lo    用 LibreOffice 无界面导出。只做预览和目录页码定位；headless 读不到 macOS 系统中文字体，
        所以用独立配置目录，把排版参数.json里的字体文件复制进 user/fonts
"""
import argparse
import glob
import json
import os
import re
import sys
import unicodedata
import shutil
import subprocess
import logging
from pathlib import Path

logging.getLogger("pypdf").setLevel(logging.ERROR)   # 不刷「Ignoring wrong pointing object」这类无害警告

SKILL = Path(__file__).resolve().parent.parent
STYLE = json.loads((SKILL / "config" / "排版参数.json").read_text(encoding="utf-8"))
IS_WIN = sys.platform == "win32"
LO_PROFILE = (Path(os.environ.get("LOCALAPPDATA", Path.home())) / "教辅资料包_lo_profile") if IS_WIN else (Path.home() / "Library/Caches/教辅资料包_lo_profile")
FONT_DIRS = (
    [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
     Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Windows/Fonts"]
    if IS_WIN else
    [Path.home() / "Library/Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
     Path("/System/Library/Fonts")]
)
FONT_KEYS = {
    "Songti": ["Songti", "simsun", "stsong"],
    "SimSun": ["simsun", "stsong"],
    "Kaiti": ["Kaiti", "simkai", "stkaiti"],
    "KaiTi": ["simkai", "stkaiti"],
    "Source Han Serif": ["SourceHanSerif", "思源宋体", "NotoSerifSC"],
    "Noto Serif": ["NotoSerifCJK", "NotoSerifSC"],
    "Times New Roman": ["Times New Roman", "times"],
}


def find_soffice():
    if shutil.which("soffice"):
        return shutil.which("soffice")
    if IS_WIN:
        for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if Path(p).exists():
                return p
    return None


def find_word():
    if IS_WIN:
        for p in (r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                  r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                  r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"):
            if Path(p).exists():
                return p
        return None
    return "/Applications/Microsoft Word.app" if Path("/Applications/Microsoft Word.app").exists() else None


def find_wps():
    if IS_WIN:
        for p in (Path(os.environ.get("LOCALAPPDATA", "")) / "Kingsoft/WPS Office/ksolaunch.exe",
                  Path(r"C:\Program Files (x86)\Kingsoft\WPS Office\ksolaunch.exe")):
            if p.exists():
                return str(p)
        return None
    return "/Applications/wpsoffice.app" if Path("/Applications/wpsoffice.app").exists() else None


def find_chrome():
    if IS_WIN:
        for p in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                  r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                  str(Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe"),
                  r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                  r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
            if Path(p).exists():
                return p
        return shutil.which("chrome") or shutil.which("msedge")
    mac = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    return mac if Path(mac).exists() else shutil.which("google-chrome")


def font_files():
    want = []
    for fam in STYLE["fonts"].values():
        if not isinstance(fam, str):
            continue
        keys = next((v for k, v in FONT_KEYS.items() if fam.startswith(k)), [fam.split()[0]])
        want += keys
    dirs = list(FONT_DIRS)
    if not IS_WIN:
        dirs += [Path(p) for p in glob.glob("/System/Library/AssetsV2/com_apple_MobileAsset_Font*/*/AssetData")]
    found = []
    for d in dirs:
        if not d.exists():
            continue
        for f in d.iterdir():
            if f.suffix.lower() in (".ttc", ".ttf", ".otf") and any(k.lower() in f.name.lower() for k in want):
                found.append(f)
    return found


def lo_profile():
    fonts = LO_PROFILE / "user/fonts"
    fonts.mkdir(parents=True, exist_ok=True)
    for f in font_files():
        dst = fonts / f.name
        if not dst.exists():
            shutil.copyfile(f, dst)
    return LO_PROFILE


def export_lo(docx):
    docx = Path(docx).resolve()
    soffice = find_soffice()
    if not soffice:
        if IS_WIN and find_word():
            return export_word(docx)
        raise FileNotFoundError("未找到 LibreOffice (soffice)")
    subprocess.run([soffice, f"-env:UserInstallation={lo_profile().as_uri()}", "--headless", "--convert-to", "pdf",
                    "--outdir", str(docx.parent), str(docx)], check=True, capture_output=True, timeout=300)
    return docx.with_suffix(".pdf")


def export_word(docx):
    """macOS 走 AppleScript，Windows 走 PowerShell Word COM。失败时抛异常，由调用方决定退回。"""
    docx = Path(docx).resolve()
    out = docx.with_suffix(".pdf")
    if IS_WIN:
        out.unlink(missing_ok=True)
        src_s = str(docx).replace("'", "''")
        out_s = str(out).replace("'", "''")
        ps = (
            "$word = New-Object -ComObject Word.Application; "
            "$word.Visible = $false; $word.DisplayAlerts = 0; "
            "try { "
            f"$doc = $word.Documents.Open('{src_s}', $false, $true); "
            f"$doc.SaveAs([ref]'{out_s}', [ref]17); "
            "$doc.Close($false) "
            "} finally { $word.Quit() }"
        )
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True, timeout=240)
        return out
    box = Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents/教辅资料包导出"
    box.mkdir(parents=True, exist_ok=True)
    src, pdf = box / "in.docx", box / "out.pdf"
    shutil.copyfile(docx, src)
    pdf.unlink(missing_ok=True)
    script = f'''
with timeout of 180 seconds
tell application "Microsoft Word"
  open POSIX file "{src}"
  set d to active document
  save as d file name "{pdf}" file format format PDF
  close d saving no
end tell
end timeout'''
    subprocess.run(["osascript", "-e", script], check=True, capture_output=True, timeout=240)
    shutil.copyfile(pdf, out)
    return out


def export(docx, engine="lo"):
    if engine == "word":
        return export_word(docx)
    return export_lo(docx)


# Word / LibreOffice 导出的 PDF 抽字时，常把汉字抽成康熙部首（⼀）或部首补充区的简体字形（⻓）；NFKC 只管得了前者
RADICALS = dict(zip("⺰⻅⻆⻈⻉⻋⻌⻐⻓⻔⻙⻚⻛⻜⻠⻢⻥⻦⻧⻨⻩⻪⻬⻮⻰⻳⻘⻝⻊⻏",
                    "纟见角讠贝车辶钅长门韦页风飞饣马鱼鸟卤麦黄黾齐齿龙龟青食足阝"))


def normalize_cjk(s):
    """抽出来的文字归一：部首字形换回正常汉字，全半角统一（NFKC），去掉所有空白。"""
    s = "".join(RADICALS.get(c, c) for c in s)
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s))


def page_count(pdf):
    if shutil.which("pdfinfo"):
        out = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
        m = re.search(r"Pages:\s+(\d+)", out)
        if m:
            return int(m[1])
    from pypdf import PdfReader
    return len(PdfReader(str(pdf)).pages)


def extract_page_texts(pdf):
    """逐页抽原始文本：优先 pdftotext，缺失时回落 pypdf。"""
    if shutil.which("pdftotext"):
        n = page_count(pdf)
        return [
            subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), str(pdf), "-"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
            for i in range(1, n + 1)
        ]
    from pypdf import PdfReader
    return [(p.extract_text() or "") for p in PdfReader(str(pdf)).pages]


def page_texts(pdf):
    return [
        re.sub(r"\s+", "", re.sub(r"[\u2e80-\u2fdf]", lambda m: unicodedata.normalize("NFKC", m.group()), t))
        for t in extract_page_texts(pdf)
    ]


def find_pages(pdf, entries):
    """按目录顺序往后找每个标题所在页；跳过目录页（带点线的页），避免答案区同名篇目抢先。"""
    texts = page_texts(pdf)
    start = 0
    for i, t in enumerate(texts[:5]):
        if len(re.findall(r"[.．·…]{8,}", t)) >= 3:  # 目录点线（第一遍还没页码，不要求后面跟数字）
            start = i + 1
    pages, cur = {}, start
    for _, title in entries:
        # 居中标题常被 pdftotext 拆开（「练习一」和「基础填空」分在两处），只用全角空格前的部分匹配
        key = re.sub(r"\s+", "", title.split("\u3000")[0])[:10]
        for i in range(cur, len(texts)):
            if key in texts[i]:
                pages[title] = i + 1
                cur = i
                break
    return pages


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--engine", default="lo", choices=["word", "lo"])
    a = ap.parse_args()
    print("->", export(a.docx, a.engine))


def pdf_to_png(pdf, prefix, dpi=100, first=None, last=None):
    """PDF 转逐页 PNG，文件名和 pdftoppm 一致（<prefix>-01.png …）。
    有 pdftoppm（poppler）就用它；没有（常见于 Windows）回落 pypdfium2：pip install pypdfium2。"""
    pdf, prefix = Path(pdf), Path(prefix)
    if shutil.which("pdftoppm"):
        cmd = ["pdftoppm", "-r", str(dpi), "-png"]
        if first:
            cmd += ["-f", str(first), "-l", str(last or first)]
        subprocess.run(cmd + [str(pdf), str(prefix)], check=True, stderr=subprocess.DEVNULL)   # 不刷「Bad bounding box」这类无害警告
        return
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise SystemExit("PDF 转图片需要 poppler（pdftoppm）或 pypdfium2：pip install pypdfium2")
    doc = pdfium.PdfDocument(str(pdf))
    n = len(doc)
    width = len(str(n))
    a, b = int(first or 1), int(last or first or n)
    for i in range(a - 1, min(b, n)):
        doc[i].render(scale=dpi / 72).to_pil().save(f"{prefix}-{i + 1:0{width}d}.png")
    doc.close()
