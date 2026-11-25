"""原作者信息清查：成品里不能留原素材的作者名、网名、机构/店铺名、系列品牌、水印、联系方式和引流语。

三处用到：
  1. 入库（ingest）后自动扫素材，把可疑字样写进 <项目>/原作者信息.txt 的「候选」区，AI 看过后把确认的名字写进「屏蔽词」区
  2. 内容稿检查（check）和手写笔记检查：屏蔽词 + 引流词命中算 error，平台名算 warn
  3. 交付检查（run.py audit <项目>）：逐个扫 成品/ 和 手写笔记/成品/ 里的 Word、PDF 正文、页眉页脚、批注、
     文件属性（作者、最后修改人、标题等）、文件名，生成《交付检查报告.md》；嵌图只能看图核对，报告里列出张数

  python3 author.py suggest <项目>
  python3 author.py audit <项目>
  python3 author.py scrub <项目>     清空 成品/ 里 Word、PDF 的作者、最后修改人等属性（旧成品补救用；build 出的新成品本来就是空的）

词表在检查词表.json的 author_info；项目专属的名字写在 <项目>/原作者信息.txt（一行一个，# 开头是注释）。
"""
import html
import json
import logging
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

logging.getLogger("pypdf").setLevel(logging.ERROR)

SKILL = Path(__file__).resolve().parent.parent
CFG = json.loads((SKILL / "config" / "检查词表.json").read_text(encoding="utf-8")).get("author_info", {})
ERROR_RE = re.compile(CFG.get("error", "(?!)"), re.I)
WARN_RE = re.compile(CFG.get("warn", "(?!)"), re.I)
SUSPECT_RE = re.compile(CFG.get("suspect", "(?!)"), re.I)
HEADING_RE = re.compile(r"^(Unit|Module|Lesson|Part|Starter)\b|^第.{1,4}[单元章节课讲部分]|^[一二三四五六七八九十]+[、.．]|^[（(][一二三四五六七八九十\d]+[)）]|^\d+[、.．]", re.I)
TERMS_FILE = "原作者信息.txt"
TERMS_HEAD = """# 原作者信息（交付前必须清零）
# 「屏蔽词」区：一行一个，写原素材里出现的作者名、笔名/网名、老师称呼、机构/工作室/店铺名、系列品牌名、账号、水印字样及其变体。
#   check / audit 会在内容稿、手写稿和全部成品里检索，命中一个都不能交付。
#   有教学意义的人名（课文作者、题干人物、名言出处）不要写进来。
# 「候选」区：入库时脚本扫出来的可疑行，只是提示；看过后把确认的名字抄到屏蔽词区，候选区可以不管。

## 屏蔽词
"""


def _P(proj):
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import paths
    return paths.of(proj)


def project_terms(proj):
    f = _P(proj).原作者信息
    if not f.exists():
        return []
    terms, on = [], False
    for line in f.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if s.startswith("## "):
            on = s == "## 屏蔽词"
            continue
        if on and s and not s.startswith("#"):
            terms.append(s)
    return terms


def scan(text, terms=()):
    """返回 [(level, 命中词, 上下文)]。屏蔽词、引流词 → error；平台名 → warn。"""
    hits = []
    for t in terms:
        for m in re.finditer(re.escape(t), text, re.I):
            hits.append(("error", t, _ctx(text, m)))
    for m in ERROR_RE.finditer(text):
        hits.append(("error", m.group(), _ctx(text, m)))
    for m in WARN_RE.finditer(text):
        hits.append(("warn", m.group(), _ctx(text, m)))
    return hits


def _ctx(text, m, w=18):
    return re.sub(r"\s+", " ", text[max(0, m.start() - w):m.end() + w]).strip()


# ---------- 入库：扫候选 ----------

def suggest(proj):
    proj = Path(proj)
    found = {}
    for f in sorted(_P(proj).素材.glob("S[0-9][0-9]_*.txt")):
        sid = f.name[:3]
        text = f.read_text(encoding="utf-8", errors="ignore")
        for i, line in enumerate(text.splitlines(), 1):
            s = line.strip()
            if s and len(s) <= 80 and (SUSPECT_RE.search(s) or ERROR_RE.search(s)):
                found.setdefault(s, f"{sid} L{i}")
        pages = re.split(r"===== 第 \d+ 页 =====", text)
        if len(pages) >= 5:   # 每页都出现的短行多半是页眉页脚里的品牌字样
            cnt = {}
            for pg in pages:
                for s in {l.strip() for l in pg.splitlines() if 2 <= len(l.strip()) <= 30}:
                    if not re.fullmatch(r"[\d\s第页共/\-—.．、]+", s):
                        cnt[s] = cnt.get(s, 0) + 1
            for s, n in cnt.items():
                if HEADING_RE.match(s) or len(s) < 4:   # 每课都有的栏目标题（Unit 1、一、单词）不是品牌
                    continue
                if n >= max(3, len(pages) * 0.3):
                    found.setdefault(s, f"{sid} 重复出现 {n} 页（疑似页眉页脚）")
    f = _P(proj).原作者信息
    old = f.read_text(encoding="utf-8") if f.exists() else TERMS_HEAD + "\n"
    head = old.split("\n## 候选")[0].rstrip() + "\n"
    body = [f"# {s}　〔{where}〕" for s, where in list(found.items())[:200]]
    f.write_text(head + "\n## 候选（脚本扫出的可疑行，仅供参考）\n" + ("\n".join(body) if body else "# 未扫出可疑行") + "\n",
                 encoding="utf-8")
    return f, len(found)


# ---------- 交付检查 ----------

def docx_parts(path):
    """Word 里所有文字：正文、表格、页眉页脚、脚注、批注、文本框；外加文件属性和嵌图数量。"""
    texts, props, images = [], {}, 0
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("word/media/"):
                images += 1
            elif n.startswith("word/") and n.endswith(".xml"):
                x = z.read(n).decode("utf-8", "ignore")
                x = re.sub(r"</w:p>", "\n", x)
                t = html.unescape(re.sub(r"<[^>]+>", "", x))
                if t.strip():
                    texts.append((n.split("/")[-1].replace(".xml", ""), t))
            elif n in ("docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"):
                x = z.read(n).decode("utf-8", "ignore")
                for k, v in re.findall(r"<(?:\w+:)?(\w+)[^>/]*>([^<]+)</", x):
                    props[k] = html.unescape(v)
    return texts, props, images


def pdf_parts(path):
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import export_pdf
    texts = [(f"第{i}页", t) for i, t in enumerate(export_pdf.extract_page_texts(path), 1)]
    props = {}
    if shutil.which("pdfinfo"):
        out = subprocess.run(["pdfinfo", str(path)], capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
        for k in ("Title", "Subject", "Keywords", "Author", "Creator", "Producer"):
            m = re.search(rf"^{k}:[ \t]*(.*)$", out, re.M)   # 只在本行取值：属性为空时不能吃到下一行
            if m and m.group(1).strip():
                props[k] = m.group(1).strip()
    else:
        try:
            from pypdf import PdfReader
            for k, v in (PdfReader(str(path)).metadata or {}).items():
                props[k.lstrip("/")] = str(v)
        except Exception:
            pass
    return texts, props


def audit(proj, roots=None, out=None):
    """roots：要查的文件夹（默认 成品 + 手写笔记成品）；out：报告写到哪（默认 3_成品/检查/交付检查报告.md）。"""
    P = _P(proj)
    proj = P.root
    terms = project_terms(proj)
    roots = roots or [P.成品, P.手写成品]
    files = [f for r in roots if r.exists() for f in sorted(r.rglob("*"))
             if f.is_file() and not f.name.startswith(("~$", ".")) and P.检查 not in f.parents and f.suffix != ".md"]
    files = sorted(set(files) | {p for p in (P.交付说明, P.封面 / "封面.html") if p.exists()}) + sorted(proj.glob("*.zip"))
    rows, notes, n_err, n_warn = [], [], 0, 0

    def add(where, text):
        nonlocal n_err, n_warn
        for lv, word, ctx in scan(text, terms):
            rows.append((lv, where, word, ctx))
            n_err += lv == "error"
            n_warn += lv == "warn"

    for f in files:
        rel = f.relative_to(proj)
        add(f"{rel}（文件名）", f.name)
        suf = f.suffix.lower()
        try:
            if suf == ".docx":
                texts, props, images = docx_parts(f)
                for part, t in texts:
                    add(f"{rel} · {part}", t)
                for k, v in props.items():
                    add(f"{rel} · 属性 {k}", v)
                bad = {k: v for k, v in props.items() if k in ("creator", "lastModifiedBy", "Company", "Manager") and v.strip()}
                if bad:
                    rows.append(("error", f"{rel} · 文件属性", "、".join(f"{k}={v}" for k, v in bad.items()), "作者 / 最后修改人 / 单位应为空"))
                    n_err += 1
                if images:
                    notes.append(f"{rel}：嵌了 {images} 张图，逐张看图确认没有水印、署名、二维码")
            elif suf == ".pdf":
                texts, props = pdf_parts(f)
                for part, t in texts:
                    add(f"{rel} · {part}", t)
                for k, v in props.items():
                    add(f"{rel} · 属性 {k}", v)
                if props.get("Author", "").strip():
                    rows.append(("error", f"{rel} · 文件属性", f"Author={props['Author']}", "PDF 作者属性应为空"))
                    n_err += 1
            elif suf in (".md", ".txt", ".html"):
                add(str(rel), f.read_text(encoding="utf-8", errors="ignore"))
            elif suf == ".zip":
                with zipfile.ZipFile(f) as z:
                    for n in z.namelist():
                        add(f"{rel} → {n}（文件名）", Path(n).name)
            elif suf in (".png", ".jpg", ".jpeg"):
                pass   # 图片文字查不到：手写笔记 PNG 的文字在手写稿 .md 里已查；其余看图核对
        except Exception as ex:
            notes.append(f"{rel}：读取失败（{type(ex).__name__}: {ex}），手动核对")

    pngs = sum(1 for f in files if f.suffix.lower() in (".png", ".jpg", ".jpeg"))
    if pngs:
        notes.append(f"成品里有 {pngs} 张图片（手写笔记页、预览图等）：文字已通过手写稿 / PDF 检索，图里的配图、水印仍需看图核对")
    if not terms:
        notes.insert(0, f"⚠️ {TERMS_FILE} 的「屏蔽词」区是空的：只查了通用引流词，原作者的具体名字没查。入库后先把原作者名、机构名、品牌名填进去")

    ok = n_err == 0
    L = [f"# 交付检查报告：{proj.name}", "",
         f"**结论：{'✅ 未发现原作者信息，可以交付' if ok else '❌ 有原作者信息残留，改完重新导出再查'}**", "",
         f"检查了 {len(files)} 个文件；屏蔽词 {len(terms)} 个；error {n_err} / warn {n_warn}", ""]
    L.append("## 一、命中（error 必须清掉；warn 逐条判断，正常教学内容如「网络问政、公众号提建议」可以保留）")
    L += [f"- [{lv}] {where}：**{word}** —— {ctx}" for lv, where, word, ctx in rows[:300]] or ["- 无"]
    L += ["", "## 二、需要人工看的"]
    L += [f"- {n}" for n in notes] or ["- 无"]
    out = Path(out) if out else P.交付检查报告
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(L) + "\n", encoding="utf-8")
    return ok, n_err, n_warn, out


# ---------- 清文件属性 ----------

DOCX_PROP_RE = re.compile(r"(<(dc:creator|cp:lastModifiedBy|dc:description|cp:keywords|dc:subject|cp:category)\b[^>]*>)[^<]*(</\2>)")
APP_PROP_RE = re.compile(r"(<(Company|Manager)\b[^>]*>)[^<]*(</\2>)")
PDF_KEYS = ("/Author", "/Subject", "/Keywords")


def scrub_docx(path):
    """清空 Word 的作者、最后修改人、备注、关键词、主题、类别、单位、经理；正文一个字不动。返回是否改过。"""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    changed, out = False, []
    for info, data in items:
        if info.filename in ("docProps/core.xml", "docProps/app.xml"):
            x = data.decode("utf-8")
            y = (DOCX_PROP_RE if info.filename.endswith("core.xml") else APP_PROP_RE).sub(r"\1\3", x)
            if y != x:
                changed, data = True, y.encode("utf-8")
        out.append((info, data))
    if changed:
        tmp = path.with_name(path.name + ".tmp")
        with zipfile.ZipFile(tmp, "w") as z:
            for info, data in out:
                z.writestr(info, data)
        tmp.replace(path)
    return changed


def scrub_pdf(path):
    """去掉 PDF 的 Author / Subject / Keywords（页面内容原样复制）。返回是否改过。"""
    from pypdf import PdfReader, PdfWriter
    path = Path(path)
    r = PdfReader(str(path))
    md = r.metadata or {}
    if not any(str(md.get(k, "")).strip() for k in PDF_KEYS):
        return False
    w = PdfWriter(clone_from=r)
    w.add_metadata({k: "" for k in PDF_KEYS})
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "wb") as f:
        w.write(f)
    tmp.replace(path)
    return True


def _scrub_pdf_any(f):
    """系统 Python 没装 pypdf 时（装不进系统 Python 时，setup 会把它装在 .venv），借 .venv 里的 Python 跑。"""
    try:
        import pypdf  # noqa: F401
    except ImportError:
        venv = SKILL / ".venv" / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
        if not venv.exists():
            print(f"  ⚠️ 没装 pypdf，{Path(f).name} 的 PDF 属性没清：pip install pypdf")
            return False
        return subprocess.run([str(venv), __file__, "scrub-pdf", str(f)]).returncode == 10
    return scrub_pdf(f)


def scrub(proj):
    """把 成品/ 和 手写笔记/成品/ 里所有 Word、PDF 的作者类属性清空（不碰 素材/）。返回改过的文件列表。"""
    P = _P(proj)
    proj = P.root
    done = []
    for r in (P.成品, P.手写成品, P.样章):
        for f in sorted(r.rglob("*")) if r.exists() else []:
            if not f.is_file() or f.name.startswith(("~$", ".")):
                continue
            try:
                if f.suffix.lower() == ".docx" and scrub_docx(f) or f.suffix.lower() == ".pdf" and _scrub_pdf_any(f):
                    done.append(f)
            except Exception as ex:
                print(f"  ⚠️ {f.relative_to(proj)} 没清成（{type(ex).__name__}: {ex}），在 Word / PDF 软件里手动删属性")
    return done


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "scrub-pdf":   # 内部用：.venv 里清单个 PDF，改过返回 10
        sys.exit(10 if scrub_pdf(sys.argv[2]) else 0)
    if len(sys.argv) != 3 or sys.argv[1] not in ("suggest", "audit", "scrub"):
        sys.exit(__doc__)
    if sys.argv[1] == "suggest":
        f, n = suggest(sys.argv[2])
        print(f"候选 {n} 条 -> {f}")
    elif sys.argv[1] == "scrub":
        for f in scrub(sys.argv[2]):
            print("  已清属性", f.name)
    else:
        ok, e, w, out = audit(sys.argv[2])
        print(f"{'✅' if ok else '❌'} 原作者信息 error {e} / warn {w}\n-> {out}")
        sys.exit(0 if ok else 1)
