"""成品闸门：build 之后自动跑，也可以单独跑  python3 run.py gates <项目>

  拦截（不通过不能进确认点）：重开校验（Word / PDF 打得开）、版本页数一致（背诵 / 默写、学生 / 答案……）
  警告（逐条判断，结论写进进度记录）：填充率（正文页 <80%；章末页不限）、目录页码对不上、出处核对（出处页里找不到关键词和数字）

结果写进 <项目>/闸门报告.md。填充率用渲染图量：每页最低一行墨迹到版心底部的比例，Mac / Windows 通用。
"""
import logging
import re
import sys
import unicodedata
from pathlib import Path

logging.getLogger("pypdf").setLevel(logging.ERROR)

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import export_pdf  # noqa: E402
import paths  # noqa: E402

PAIRS = [("背诵版", "默写版"), ("学生版", "答案版"), ("学生版", "教师版"), ("原卷版", "解析版"), ("题目版", "答案版")]
BODY_MIN = 0.80   # 正文页；章末页（下一页是新的章或节）、首页、章首页不限


def _norm(s):
    return export_pdf.normalize_cjk(s)


# ---------- 重开校验 ----------

def reopen(files):
    bad = []
    for f in files:
        try:
            if f.suffix == ".docx":
                import docx
                docx.Document(str(f))
            elif f.suffix == ".pdf" and export_pdf.page_count(f) < 1:
                bad.append(f"{f.name}：0 页")
        except Exception as ex:
            bad.append(f"{f.name}：打不开（{type(ex).__name__}）")
    return bad


# ---------- 版本一致 ----------

def versions(pdfs):
    bad, ok = [], []
    by = {f.stem: f for f in pdfs}
    for stem, f in by.items():
        for a, b in PAIRS:
            if a in stem:
                other = by.get(stem.replace(a, b))
                if other:
                    na, nb = export_pdf.page_count(f), export_pdf.page_count(other)
                    (ok if na == nb else bad).append(f"{f.name} {na} 页 / {other.name} {nb} 页")
    return bad, ok


# ---------- 填充率 ----------

def _chapter_titles(proj):
    """章（# ）总是另起页；节（## ）只有样式.json 里没关掉「细节.节.另起页」时才算分界。"""
    P = paths.of(proj)
    c = P.内容稿
    if not c.exists():
        return []
    style = P.排版 / "样式.json"
    sec_break = True
    if style.exists():
        import json
        sec_break = json.loads(style.read_text(encoding="utf-8")).get("细节", {}).get("节", {}).get("另起页", True)
    heads = ("# ", "## ") if sec_break else ("# ",)
    return [_norm(re.sub(r"〔[^〕]*〕", "", l.lstrip("#").strip())) for l in c.read_text(encoding="utf-8").splitlines() if l.startswith(heads)]


def fill_rates(pdf, chapters, dpi=40):
    """返回 [(页码, 填充率, 类型)]，类型：首页 / 章首 / 章末 / 正文。"""
    from PIL import Image
    import tempfile
    texts = [_norm(t) for t in export_pdf.extract_page_texts(pdf)]
    n = len(texts)
    starts = set()
    for i, t in enumerate(texts):
        head = t[:120]
        if any(c and c[:12] in head for c in chapters):
            starts.add(i)
    out = []
    with tempfile.TemporaryDirectory() as d:
        export_pdf.pdf_to_png(pdf, str(Path(d) / "p"), dpi=dpi)
        pngs = sorted(Path(d).glob("p-*.png"))
        for i, png in enumerate(pngs):
            im = Image.open(png).convert("L")
            w, h = im.size
            top, bottom = int(h * 0.06), int(h * 0.93)       # 去掉页眉、页脚区
            px = im.load()
            low = top
            for y in range(bottom, top, -1):
                if any(px[x, y] < 200 for x in range(int(w * 0.08), int(w * 0.92), 2)):
                    low = y
                    break
            rate = (low - top) / (bottom - top)
            kind = "首页" if i == 0 else "章首" if i in starts else "章末" if (i + 1 in starts or i == n - 1) else "正文"
            out.append((i + 1, rate, kind))
    return out


# ---------- 出处核对 ----------

def source_check(proj, limit=60):
    P = paths.of(proj)
    c = P.内容稿
    if not c.exists():
        return []
    src = {}
    for f in P.素材.glob("S[0-9][0-9]_*.txt"):
        pages = re.split(r"===== 第 (\d+) 页 =====", f.read_text(encoding="utf-8", errors="ignore"))
        src[f.name[:3]] = {int(pages[k]): _norm(pages[k + 1]) for k in range(1, len(pages) - 1, 2)}
        src[f.name[:3]][0] = _norm("".join(pages))
    warn = []
    for line in c.read_text(encoding="utf-8").splitlines():
        refs = re.findall(r"〔([^〕]+)〕", line)
        if not refs or line.lstrip().startswith(("#", "!", ">", "【")):   # 标题、标签、导语是新拟的，不核
            continue
        keyw = re.findall(r"\{\{(.+?)\}\}", line)
        body = re.sub(r"〔[^〕]*〕|\{\{|\}\}|\*\*", "", line)
        body = re.sub(r"^\s*(?:\d+|[（(]\d+[)）])\s*[.．、]?\s*", "", body)   # 去掉题号、条目号（成品里会重新编号）
        nums = [n for n in re.findall(r"\d+(?:\.\d+)?", body) if len(n) >= 2]
        phrases = re.findall(r"[\u4e00-\u9fff]{6,}", body)[:3]
        if not (nums or keyw or phrases):
            continue
        pool = ""
        for ref in refs:
            for part in re.split(r"[；;]", ref):
                m = re.match(r"\s*(S\d\d)\s*(?:p(\d+))?", part)
                if m and m.group(1) in src:
                    pg = int(m.group(2)) if m.group(2) else 0
                    pool += src[m.group(1)].get(pg, "") + src[m.group(1)].get(pg - 1, "") + src[m.group(1)].get(pg + 1, "")
        if not pool:
            continue

        def found(ph):       # 整合升级会改写表述：长句只要大部分 4 字片段能在素材里找到就算
            grams = [ph[i:i + 4] for i in range(0, len(ph) - 3, 2)]
            return sum(g in pool for g in grams) >= max(1, len(grams) * 0.5)
        miss = [k for k in nums + keyw if _norm(k) not in pool] + [k for k in phrases if not found(_norm(k))]
        keys = nums + keyw + phrases
        if len(miss) > len(keys) / 2:
            warn.append(f"〔{'；'.join(refs)}〕 找不到：{'、'.join(miss[:4])} —— {body.strip()[:40]}")
        if len(warn) >= limit:
            break
    return warn


# ---------- 目录核对 ----------

def toc_check(pdf, titles, toc_pages=3):
    """目录页（前几页）里「标题 … 页码」和标题在正文实际所在页比：封面会让页码整体偏移，取最常见的偏移量，只报和它不一样的。"""
    from collections import Counter
    texts = [_norm(t) for t in export_pdf.extract_page_texts(pdf)]
    pairs = []
    for t in titles:
        if len(t) < 4:
            continue
        key = re.escape(t[:16])
        for k in range(min(toc_pages, len(texts))):
            if "目录" not in texts[k]:                   # 只核写着「目录」的页（打卡表等不是目录）
                continue
            m = re.search(key + r"[^\d]{0,40}?(\d{1,3})", texts[k])
            if m:
                body = next((j for j in range(k + 1, len(texts)) if t[:16] in texts[j]), None)
                if body is not None:
                    pairs.append((t, int(m.group(1)), body + 1))
                break
    if len(pairs) < 3:
        return []
    off = Counter(real - shown for t, shown, real in pairs).most_common(1)[0][0]
    return [f"{pdf.name}：目录写第 {shown} 页，正文在第 {real - off} 页 —— {t[:20]}" for t, shown, real in pairs if real - shown != off]


def run(proj, out_dir=None):
    """out_dir：查哪个文件夹的成品（默认 3_成品；出样章时是 2_样张/样章）。"""
    P = paths.of(proj)
    proj = P.root
    out_dir = Path(out_dir) if out_dir else P.成品
    files = sorted(f for f in out_dir.glob("*") if f.suffix in (".docx", ".pdf") and not f.name.startswith("~$")) if out_dir.exists() else []
    pdfs = [f for f in files if f.suffix == ".pdf"]
    block, warn, notes = [], [], []
    block += [f"重开校验：{b}" for b in reopen(files)]
    vb, vok = versions(pdfs)
    block += [f"版本页数不一致：{b}" for b in vb]
    notes += [f"版本页数一致：{o}" for o in vok]
    chapters = _chapter_titles(proj)
    lows = []
    for pdf in pdfs:
        for pg, rate, kind in fill_rates(pdf, chapters):
            need = BODY_MIN if kind == "正文" else None
            if need and rate < need:
                lows.append(f"{pdf.name} 第 {pg} 页（{kind}）只排了 {rate:.0%}，要 ≥{need:.0%}")
    warn += lows
    for pdf in pdfs:
        warn += [f"目录页码：{w}" for w in toc_check(pdf, chapters)]
    warn += [f"出处核对：{w}" for w in source_check(proj)]
    ok = not block
    L = [f"# 闸门报告：{proj.name}", "", f"**结论：{'✅ 拦截项全部通过' if ok else '❌ 有拦截项，改完重新 build'}**　拦截 {len(block)}　警告 {len(warn)}", "",
         "## 拦截（必须清零）"] + ([f"- {b}" for b in block] or ["- 无"]) + \
        ["", "## 警告（逐条判断，结论写进进度记录）"] + ([f"- {w}" for w in warn[:200]] or ["- 无"]) + \
        ["", "## 通过项"] + ([f"- {n}" for n in notes] or ["- 无"]) + \
        ["", "填充率：正文页 ≥80%；章末页（下一页是新的章或节）、首页、章首页不限。填不满补素材里的细节，不放大字号。",
         "出处核对只是提示：找不到的多半是改写过表述或抽字有误，逐条回素材看。"]
    rep = P.闸门报告 if out_dir == P.成品 else out_dir.parent / "样章闸门报告.md"
    rep.parent.mkdir(parents=True, exist_ok=True)
    rep.write_text("\n".join(L) + "\n", encoding="utf-8")
    return ok, len(block), len(warn), rep


if __name__ == "__main__":
    ok, b, w, rep = run(sys.argv[1])
    print(f"{'✅' if ok else '❌'} 闸门 拦截 {b} / 警告 {w} -> {rep}")
    sys.exit(0 if ok else 1)
