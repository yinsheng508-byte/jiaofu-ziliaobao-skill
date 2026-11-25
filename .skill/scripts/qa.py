"""取题和自查的小工具（check 查不到的地方）。用法：python3 qa.py <子命令> ...

  stems   <项目> <Sxx> <起行> <止行>   列素材某段的题号行、选项和【答案】（跳过解析），挑题用
  scan    <项目>                      扫内容稿练习：空数和答案份数、答案在不在选项里、解析「故选」是否一致、
                                      看图/划线依赖、全书重复题
  sources <项目> [讲解.md ...]         讲解里的英文例句能不能在素材里原样找到（找不到的逐条人工核：多半是原题填了答案）
  grep    <项目> <Sxx,Syy> <正则>...   在素材里按关键词看上下文，核讲解的中文说法
  pages   <项目> [--pdf 文件]          成品 PDF：字很少的孤页、目录页码和页脚是否一致
"""
import argparse, collections, glob, re, subprocess, sys, unicodedata
from pathlib import Path

SRC = re.compile(r"〔[^〕]*〕")
SKIP = ("【解析】", "【详解】", "【分析】", "【点睛】")


def _P(proj):
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import paths
    return paths.of(proj)


def src_file(proj, sid):
    return next(_P(proj).素材.glob(f"{sid}_*.txt"))


def stems(a):
    L = src_file(a.project, a.sid).read_text(encoding="utf-8").split("\n")
    skip = False
    for i in range(a.start - 1, min(a.end, len(L))):
        s = L[i].strip()
        if not s or re.match(r"^\d+\s*/\s*\d+$", s):
            continue
        if s.startswith(SKIP):
            skip = True; continue
        if s.startswith("【答案】"):
            print(f"   =>{s[4:][:100]}"); skip = True; continue
        if re.match(r"^\d+\s*[．.](?!\d)", s) or re.match(r"^[一二三四五六七八九十]+、", s) or s.startswith("题型"):
            skip = False
        if not skip:
            print(f"{i + 1}:{s[:120]}")


def parse_draft(path):
    """内容稿 → 练习题和答案，都按 (分组标题, 题号) 编号。兼容章内答案（[答案]）和书末答案（# 参考答案）。"""
    ex, an = collections.OrderedDict(), collections.OrderedDict()
    zone, grp, cur = "explain", None, None
    for l in Path(path).read_text(encoding="utf-8").split("\n"):
        if l.startswith("[答案]") or l.strip() == "# 参考答案":
            zone, grp, cur = "answer", None, None; continue
        if l.startswith("# "):
            zone, grp, cur = "explain", None, None; continue
        if l.startswith("### "):
            grp, cur = l[4:].strip(), None
            if zone != "answer":
                zone = "exercise"
            continue
        if l.startswith("## ") or zone == "explain" or not grp:
            continue
        box = an if zone == "answer" else ex
        m = re.match(r"^(\d+)\.(.*)", l)
        if m:   # 短文选择题「1.A．…｜B．…」本身就是一道题
            cur = (grp, int(m.group(1))); box[cur] = [l]; continue
        if re.match(r"^[一二三四五六七八九十]+、", l):
            cur = None; continue
        if cur and l.strip():
            box[cur].append(l)
    return ex, an


def scan(a):
    ex, an = parse_draft(_P(a.project).内容稿)
    probs, seen = [], {}
    for k, v in ex.items():
        txt = "\n".join(v); body = SRC.sub("", txt)
        if k not in an:
            probs.append((k, "没有答案")); continue
        ans = SRC.sub("", an[k][0]).split("参考答案：", 1)[-1].strip()
        why = " ".join(an[k][1:])
        opts = [l for l in v if re.match(r"^(\d+\.)?A[．.]", l)]
        blanks = len(re.findall(r"＿+", body))
        if opts:
            letters = re.findall(r"(?:^|｜)\s*(?:\d+\.)?([A-D])[．.]", " ｜".join(opts))
            if ans not in letters:
                probs.append((k, f"答案 {ans} 不在选项 {letters} 里"))
            m = re.search(r"故选\s*([A-D])", why)
            if m and m.group(1) != ans:
                probs.append((k, f"解析「故选{m.group(1)}」和答案 {ans} 不一致"))
        elif "[横线" not in txt:
            parts = [p for p in re.split(r"[；;]", ans) if p.strip()]
            if blanks == 0:
                probs.append((k, "填空题没有空"))
            elif blanks > 1 and blanks != len(parts):
                probs.append((k, f"{blanks} 个空、答案 {len(parts)} 份：{ans}（一条长空填几个词时用 ansgroup 分组）"))
        m = re.search(r".{0,10}(图片|如图|看图|划线|画线|下划线).{0,10}", why)
        if m:
            probs.append((k, "解析提到图/划线，确认题干看得出：" + m.group()))
        norm = re.sub(r"[^a-z]", "", SRC.sub("", v[0]).lower().split(".", 1)[-1])[:40]
        if len(norm) > 15:
            if norm in seen:
                probs.append((k, f"和 {seen[norm]} 开头相同，确认是不是重复题"))
            else:
                seen[norm] = k
    for k, m in probs:
        print(k, m)
    print(f"练习 {len(ex)} 题，答案 {len(an)} 条，提示 {len(probs)} 条")


def sources(a):
    S = re.sub(r"[^a-z0-9]", "", "".join(p.read_text(encoding="utf-8") for p in sorted(_P(a.project).素材.glob("S*.txt"))).lower().replace("’", "'"))
    files = a.files or sorted(glob.glob(str(_P(a.project).选题配置 / "*_讲解.md")))
    n = 0
    for f in files:
        for i, l in enumerate(Path(f).read_text(encoding="utf-8").split("\n"), 1):
            l2 = re.sub(r"〔[^〕]*〕|\{\{|\}\}", "", l)
            for seg in re.findall(r"[A-Za-z][A-Za-z0-9'’ ,.!?\-—/()]{14,}", l2):
                for piece in re.split(r"\s/\s|[()]|—", seg):
                    if len(re.findall(r"[A-Za-z']+", piece)) < 4:
                        continue
                    if re.sub(r"[^a-z0-9]", "", piece.lower().replace("’", "'")) not in S:
                        print(f"{Path(f).name}:{i}: {piece.strip()}"); n += 1
    print(f"素材里找不到原样的英文片段 {n} 处（逐条回素材核：原题填答案、去掉了括号提示等算正常）")


def grep(a):
    for pat in a.patterns:
        print("==", pat)
        k = 0
        for sid in a.sids.split(","):
            for f in _P(a.project).素材.glob(f"{sid}_*.txt"):
                for i, l in enumerate(f.read_text(encoding="utf-8").split("\n"), 1):
                    for m in re.finditer(pat, l):
                        if k < 4:
                            print(f"  {sid} L{i}: …{l[max(0, m.start() - 50):m.end() + 70]}…")
                        k += 1
        if k > 4:
            print(f"  （共 {k} 处）")


def pages(a):
    import export_pdf
    proj = Path(a.project)
    pdf = Path(a.pdf) if a.pdf else next(iter(sorted(_P(proj).成品.glob("*（完整版）.pdf"))))
    P = [re.sub(r"[\u2e80-\u2fdf]", lambda m: unicodedata.normalize("NFKC", m.group()), t)
         for t in export_pdf.extract_page_texts(pdf)]
    n = len(P)
    is_toc = lambda pg: len(re.findall(r"\.{8,}", pg)) >= 3
    foot = lambda pg: (lambda m: int(m.group(1)) if m else None)(re.search(r"第\s*(\d+)\s*页", pg))
    short = [(i + 1, len(re.sub(r"\s", "", pg))) for i, pg in enumerate(P)
             if foot(pg) and not is_toc(pg) and len(re.sub(r"\s", "", pg)) < a.min_chars]
    for i, c in short:
        print(f"孤页？PDF 第 {i} 页（页脚第 {foot(P[i - 1])} 页）只有 {c} 字：{re.sub(r'\s+', ' ', P[i - 1])[:40]}")
    toc = [(m.group(1).strip(), int(m.group(2))) for pg in P if is_toc(pg) for l in pg.split("\n")
           for m in [re.match(r"^(.*?)\.{4,}\s*(\d+)$", l.strip())] if m]
    bad = []
    for t, p in toc:
        real = next((foot(pg) for pg in P if not is_toc(pg) and t in pg.replace("\n", " ")), None)
        if real != p:
            bad.append((t, p, real))
    for t, p, r in bad:
        print(f"目录不对：{t} 目录写 {p}，正文在第 {r} 页")
    print(f"共 {n} 页；孤页 {len(short)} 处；目录 {len(toc)} 条，不对 {len(bad)} 条")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("stems"); p.add_argument("project"); p.add_argument("sid")
    p.add_argument("start", type=int); p.add_argument("end", type=int)
    p = sub.add_parser("scan"); p.add_argument("project")
    p = sub.add_parser("sources"); p.add_argument("project"); p.add_argument("files", nargs="*")
    p = sub.add_parser("grep"); p.add_argument("project"); p.add_argument("sids"); p.add_argument("patterns", nargs="+")
    p = sub.add_parser("pages"); p.add_argument("project"); p.add_argument("--pdf")
    p.add_argument("--min-chars", type=int, default=150)
    a = ap.parse_args()
    {"stems": stems, "scan": scan, "sources": sources, "grep": grep, "pages": pages}[a.cmd](a)


if __name__ == "__main__":
    main()
