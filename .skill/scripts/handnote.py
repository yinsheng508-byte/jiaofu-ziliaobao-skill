"""手写笔记排版：手写稿.md → A4 图片 + 可打印 PDF（尺寸: 小红书 时出 3:4 图）（规范见 ../templates/手写笔记/手写笔记规范.md）。

  python3 handnote.py <手写稿.md>... [--out 目录] [--font 09] [--scale 2] [--pages 1-3] [--no-pdf]
  python3 handnote.py init <项目或目录> <样张编号 01–23> [文件名]      复制一份样张当起点

手写稿 = 头部 --- 键: 值 --- + 正文；正文用 === 单独一行分页。
产物：<out>/p-01.png …（A4 2160×3054）、<稿名>.pdf（A4 打印版）、总览.png、检查报告.md、_html/（中间 HTML）。
"""
import argparse
import html
import json
import os
import re
import shutil
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
NOTE = HERE.parent / "assets" / "手写笔记"   # 样式、字体、样张


def _find_chrome():   # 统一用 export_pdf.find_chrome（Mac / Windows，Chrome 没有用 Edge）
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import export_pdf
    return export_pdf.find_chrome() or ("chrome" if sys.platform == "win32" else "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")


CHROME = _find_chrome()
FONTS = json.loads((NOTE / "字体表.json").read_text(encoding="utf-8"))
CARD_COLORS = ["#f7c9c9", "#c9dcf5", "#fbe7a6", "#c8e8cf", "#ddd0f1", "#fbd2b0", "#c6ecec"]
COLOR_TAGS = {"红": "r", "蓝": "b", "绿": "g", "紫": "v", "橙": "o", "灰": "gray"}

# 版式默认值：纸张、标题样式、页码样式、默认字体
LAYOUTS = {
    "横线本": dict(纸张="横线本", 标题样式="色块", 页码="-n-", 字体="09"),
    "表格框": dict(纸张="表格框", 标题样式="色块", 页码="无", 字体="09"),
    "纯色":   dict(纸张="纯色", 标题样式="马克笔", 页码="n / N", 字体="11"),
    "白底":   dict(纸张="白底", 标题样式="红字", 页码="无", 字体="14"),
    "导图":   dict(纸张="导图", 标题样式="红字", 页码="无", 字体="16"),
    "方格":   dict(纸张="方格", 标题样式="色块", 页码="-n-", 字体="09"),
    "点阵":   dict(纸张="点阵", 标题样式="马克笔", 页码="n / N", 字体="01"),
    "网格":   dict(纸张="网格", 标题样式="色块", 页码="-n-", 字体="16"),
    "牛皮纸": dict(纸张="牛皮纸", 标题样式="马克笔", 页码="n / N", 字体="11"),
    "便签":   dict(纸张="便签", 标题样式="马克笔", 页码="无", 字体="01"),
    "康奈尔": dict(纸张="康奈尔", 标题样式="色块", 页码="-n-", 字体="09"),
    "四线格": dict(纸张="四线格", 标题样式="马克笔", 页码="-n-", 字体="01"),
    "田字格": dict(纸张="田字格", 标题样式="红字", 页码="-n-", 字体="26"),
}


# ---------------- 行内 ----------------
def inline(s, meta):
    codes = []

    def keep(m):
        codes.append(html.escape(m.group(1), quote=False))
        return f"\x00{len(codes) - 1}\x00"
    s = re.sub(r"`([^`]+)`", keep, s)
    s = re.sub(r"\s*〔S\d+[^〕]*〕", "", s)               # 出处只留在稿子里，不上图
    s = html.escape(s, quote=False)
    show = meta.get("显示答案", "否") in ("是", "true", "1")

    def brace(m):
        tag, val = m.group(1), m.group(2)
        if tag in COLOR_TAGS:
            return f'<span class="{COLOR_TAGS[tag]}">{val}</span>'
        if tag == "印":
            return f'<span class="prt">{val}</span>'
        if tag == "分":
            a, _, b = val.partition("|")
            return f'<span class="frac"><span>{a}</span><span>{b}</span></span>'
        if tag == "章":
            return f'<span class="stamp">{val}</span>'
        if tag == "贴":
            return f'<span class="sticker">{val}</span>'
        if tag == "框":
            return f'<span class="box">{val}</span>'
        if tag == "粉":
            return f'<span class="hl-pink">{val}</span>'
        if tag in ("注", "拼"):
            a, _, b = val.partition("|")
            return f'<ruby class="{"py" if tag == "拼" else "ann"}">{a}<rt>{b}</rt></ruby>'
        if tag == "着":
            return f'<span class="emph">{val}</span>'
        if tag == "圈":
            return "".join(f'<span class="circ">{c}</span>' for c in val)
        if tag == "删":
            return f'<span class="del">{val}</span>'
        if tag == "双":
            return f'<span class="dbl">{val}</span>'
        if tag == "描":
            return f'<span class="trace">{val}</span>'
        if tag == "标":
            return f'<span class="cap">{val}</span>'
        if tag == "签":
            return f'<span class="tag">{val}</span>'
        if tag.startswith("空#"):                    # 章末答案模式：题里只留编号 + 空，答案集中到章末
            n = max(2, len(re.sub(r"<[^>]+>", "", val)))
            return f'<span class="qn">({tag[2:]})</span><span class="blank" style="--n:{n}">&nbsp;</span>'
        if tag == "空":
            n = max(2, len(re.sub(r"<[^>]+>", "", val)))
            return f'<span class="blank" style="--n:{n}">{val if show else "&nbsp;"}</span>'
        return m.group(0)

    prev = None
    while prev != s:                       # 由内向外，支持嵌套
        prev = s
        s = re.sub(r"\{([^{}|]+)\|([^{}]*)\}", brace, s)
    s = s.replace("{点}", '<span class="dot blue"></span>').replace("{红点}", '<span class="dot red"></span>')
    s = s.replace("{箭}", '<span class="arrow">→</span>')
    s = re.sub(r"\[\[(.+?)\]\]", r'<span class="box">\1</span>', s)
    s = re.sub(r"\*\*(.+?)\*\*", r'<span class="r">\1</span>', s)
    s = re.sub(r"==(.+?)==", r'<span class="hl">\1</span>', s)
    s = re.sub(r"~~(.+?)~~", r'<span class="wave">\1</span>', s)
    s = re.sub(r"__(.+?)__", r'<span class="ul">\1</span>', s)
    return re.sub("\x00(\\d+)\x00", lambda m: f'<span class="cap">{codes[int(m.group(1))]}</span>', s)


# ---------------- 行块 ----------------
def line_block(ln, meta):
    st = ln.rstrip()
    if not st.strip():
        return ""
    if st.strip() == "[空行]":
        return '<div class="row"></div>'
    if st.startswith("#### "):
        t = inline(st[5:], meta)
        t = re.sub(r"^(\(\d+\)|（\d+）)", r'<span class="n">\1</span>', t)
        return f'<div class="row h4">{t}</div>'
    if st.startswith("### "):
        return f'<div class="h3"><span>{inline(st[4:], meta)}</span></div>'
    if st.startswith("## "):
        return f'<div class="h2"><span>{inline(st[3:], meta)}</span></div>'
    if st.startswith("#> "):
        return f'<div class="sec"><i class="l"><b></b><b></b></i><span>{inline(st[3:], meta)}</span><i class="r"><b></b><b></b></i></div>'
    if st.startswith("# "):
        t = inline(st[2:], meta)
        sty = meta["标题样式"]
        if sty == "色块":
            return f'<div class="h1"><span>{t}</span></div>'
        return f'<div class="title{" red" if sty == "红字" else ""}"><span class="mk">{t}</span></div>'
    if st.startswith("> "):
        return f'<div class="row sub">{inline(st[2:], meta)}</div>'
    m = re.match(r"^\[( |x|X|√)\]\s*(.*)", st)
    if m:
        on = m.group(1) != " "
        return f'<div class="row chk{" on" if on else ""}"><span class="cb">{"✓" if on else ""}</span>{inline(m.group(2), meta)}</div>'
    m = re.match(r"^(\d{1,2})[.．]\s+(.*)", st)
    if m:
        return f'<div class="row ol"><span class="on">{m.group(1)}</span><span class="ot">{inline(m.group(2), meta)}</span></div>'
    if st.startswith("* "):
        return f'<div class="row ul-dot"><span class="bd"></span>{inline(st[2:], meta)}</div>'
    if st.startswith("- "):
        return f'<div class="row dash"><span class="d">-</span>{inline(st[2:], meta)}</div>'
    m = re.match(r"^(例\d*[:：])(.*)", st)
    if m:
        return f'<div class="row ex"><span class="dot blue"></span><span class="lab">{m.group(1)}</span><span class="b">{inline(m.group(2), meta)}</span></div>'
    m = re.match(r"^((?:注|注意|反例|易错)[:：])(.*)", st)
    if m:
        return f'<div class="row note"><span class="dot red"></span><span class="lab">{m.group(1)}</span>{inline(m.group(2), meta)}</div>'
    if st.startswith("  "):
        return f'<div class="row indent">{inline(st.strip(), meta)}</div>'
    if st.startswith("字 "):                        # 田字格一行：一个字一格，全角空格留空格
        return f'<div class="row zi">{inline(st[2:], meta)}</div>'
    if st.startswith("en "):                        # 四线格一行英文
        t = st[3:].strip()
        return f'<div class="row en">{"&nbsp;" if t == "[空行]" else inline(t, meta)}</div>'
    if st.startswith("^ "):
        return f'<div class="row center">{inline(st[2:], meta)}</div>'
    return f'<div class="row">{inline(st, meta)}</div>'


def flow(lines, meta):
    return "".join(line_block(ln, meta) for ln in lines)


def blk_box(lines, meta, arg):
    dot = " has-dot" if lines and re.match(r"^\s*(注|注意|反例|易错)", lines[0]) else ""
    inner = "".join(f'<div class="row">{inline(ln.strip(), meta)}</div>' if ln.strip() else "" for ln in lines)
    return f'<div class="boxed{dot}">{inner}</div>'


def blk_print(lines, meta, arg):
    return '<div class="pblock">' + "".join(f"<div>{inline(ln.strip(), meta)}</div>" for ln in lines if ln.strip()) + "</div>"


def blk_question(lines, meta, arg):
    opts, sol = [], []
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        m = re.match(r"^(\*?)([A-H][.．、].*)", s)
        if m:
            opts.append(f'<div><span class="{"ok" if m.group(1) else ""}">{inline(m.group(2), meta)}</span></div>')
        else:
            sol.append(f"<div>{inline(s, meta)}</div>")
    return f'<div class="exq"><div class="opts">{"".join(opts)}</div><div class="sol">{"".join(sol)}</div></div>'


def blk_pre(lines, meta, arg):
    return "".join(f'<div class="row pre">{inline(ln, meta) or "&nbsp;"}</div>' for ln in lines)


def split_cells(r):
    """按 | 分列，{分|a|b} 这类花括号里的 | 不算。"""
    out, cur, depth = [], "", 0
    for ch in r:
        depth += (ch == "{") - (ch == "}")
        if ch == "|" and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    return out + [cur]


def blk_table(lines, meta, arg):
    rows = [ln.strip() for ln in lines if ln.strip()]
    out, head = [], len(rows) > 1 and re.fullmatch(r"\|?[\s:\-|]+\|?", rows[1])
    for i, r in enumerate(rows):
        if head and i == 1:
            continue
        cells = [c.strip() for c in split_cells(r.strip().strip("|"))]
        tag = "th" if head and i == 0 else "td"
        tds = []
        for c in cells:
            short = len(re.sub(r"\{[^|{}]+\|", "", c)) <= 8
            tds.append(f'<{tag}{" class=c" if short else ""}>{inline(c, meta)}</{tag}>')
        out.append("<tr>" + "".join(tds) + "</tr>")
    cls = {"浅": "ghost", "对比": "vs", "首列": "firstcol", "无框": "bare", "问答": "qa"}
    extra = "".join(" " + cls[w] for w in arg.split() if w in cls)
    return f'<table class="t{extra}">{"".join(out)}</table>'


def blk_tree(lines, meta, arg):
    items = []
    for ln in lines:
        if not ln.strip():
            continue
        d = (len(ln) - len(ln.lstrip(" "))) // 2
        items.append((d, ln.strip()))

    def build(i, depth):
        out = []
        while i < len(items) and items[i][0] >= depth:
            d, t = items[i]
            kids_start = i + 1
            j = kids_start
            while j < len(items) and items[j][0] > d:
                j += 1
            if t.startswith(">"):
                out.append(f'<div class="tbox">{inline(t[1:].strip(), meta)}</div>')
            elif j > kids_start:
                inner, _ = build(kids_start, d + 1)
                out.append(f'<div class="node lv{min(d, 3)}"><div class="lab">{inline(t, meta)}</div><div class="kids">{inner}</div></div>')
            else:
                t2 = re.sub(r"^([^：:]{1,6}[：:])", r'<span class="k">\1</span>', inline(t, meta))
                out.append(f'<div class="leaf">{t2}</div>')
            i = j
        return "".join(out), i

    inner, _ = build(0, 0)
    return f'<div class="tree">{inner}</div>'


def blk_card(lines, meta, arg):
    meta["_card"] = meta.get("_card", 0) + 1
    n = meta["_card"]
    c = CARD_COLORS[(n - 1) % len(CARD_COLORS)]
    lines = [ln.strip() for ln in lines if ln.strip()]
    head, rest = lines[0], lines[1:]
    tx, tip = [], ""
    for ln in rest:
        m = re.match(r"^(启示|点拨|用法|总结)[:：](.*)", ln)
        if m:
            tip = f'<div class="tip"><b>{m.group(1)}：</b><span>{inline(m.group(2), meta)}</span></div>'
        else:
            tx.append(f'<p class="tx">{inline(ln, meta)}</p>')
    return f'<div class="card" style="--c:{c}"><div class="hd"><div class="num">{n}</div><h3>{inline(head, meta)}</h3></div>{"".join(tx)}{tip}</div>'


def blk_toc(lines, meta, arg):
    out, n = [], 0
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        if s.startswith("说明:") or s.startswith("说明："):
            out.append(f'<div class="meta">{inline(s[3:], meta)}</div>')
            continue
        n += 1
        t, _, p = s.rpartition("|") if "|" in s else (s, "", "")
        c = CARD_COLORS[(n - 1) % len(CARD_COLORS)]
        out.append(f'<div class="it" style="--c:{c}"><div class="num">{n}</div><div>{inline(t.strip(), meta)}</div><div class="lead"></div><div>{p.strip()}</div></div>')
    return f'<div class="toc">{"".join(out)}</div>'


def wy_sentence(s, meta):
    out, k = [], 0
    for tok in re.findall(r"\[[^\]]+\]\([^)]*\)|\||[^\[|]", s):
        if tok == "|":
            out.append('<span class="bar"></span>')
        elif tok.startswith("["):
            w, a = re.match(r"\[([^\]]+)\]\(([^)]*)\)", tok).groups()
            c = "var(--red)" if k % 2 == 0 else "var(--orange)"
            k += 1
            lo = " lo" if k % 2 == 0 else ""
            out.append(f'<span class="w{lo}" style="--c:{c}"><span>{html.escape(w)}</span><span class="br"></span><span class="an">{inline(a, meta)}</span></span>')
        else:
            out.append(html.escape(tok))
    return "".join(out)


def blk_wenyan(lines, meta, arg):
    groups, cur = [], []
    for ln in lines + [""]:
        if ln.strip():
            cur.append(ln.strip())
        elif cur:
            groups.append(cur)
            cur = []
    nums = "①②③④⑤⑥⑦⑧⑨⑩"
    out = []
    for g in groups:
        meta["_wy"] = meta.get("_wy", 0) + 1
        sen, src, tr = g[0], "", ""
        for ln in g[1:]:
            if ln.startswith("——"):
                src = ln
            elif re.match(r"^(简译|译文|翻译)[:：]", ln):
                lab, body = re.split(r"[:：]", ln, maxsplit=1)
                tr = f'<div class="tr">{lab}：{inline(body, meta)}</div>'
        no = nums[(meta["_wy"] - 1) % 10]
        out.append(f'<div class="wy"><div class="sen"><span class="no">{no}</span><span class="line">{wy_sentence(sen, meta)}</span>'
                   f'{f"<span class=src>{html.escape(src)}</span>" if src else ""}</div>{tr}</div>')
    return "".join(out)


def blk_cover(lines, meta, arg):
    kv = {}
    for ln in lines:
        if re.match(r"^\s*[^：:]+[：:]", ln):
            k, v = re.split(r"[：:]", ln.strip(), maxsplit=1)
            kv[k.strip()] = v.strip()
    tags = "".join(f'<span style="--c:{CARD_COLORS[i % 7]}">{inline(t.strip(), meta)}</span>' for i, t in enumerate(kv.get("标签", "").split("|")) if t.strip())
    big = "<br>".join(f'<span class="mk">{inline(x, meta)}</span>' for x in kv.get("标题", "").split("/"))
    return (f'<div class="cover"><div class="kicker">{inline(kv.get("小字", ""), meta)}</div><div class="big">{big}</div>'
            f'<div class="subt">{inline(kv.get("副标题", ""), meta)}</div><div class="tags">{tags}</div></div>'
            f'<div class="by">{inline(kv.get("署名", ""), meta)}</div>'
            '<div class="doodle" style="left:90px;top:120px">✶</div><div class="doodle" style="right:110px;top:220px">☁</div>'
            '<div class="doodle" style="right:90px;bottom:200px">✎</div><div class="doodle" style="left:120px;bottom:260px">♡</div>')


SHOW = ("是", "true", "1")
BRANCH = ["#e2372e", "#3340c4", "#2f7d4f", "#e5892b", "#8b67c9", "#1f9aa0"]


def show_ans(meta):
    return meta.get("显示答案", "否") in SHOW


def paren(ans, meta):
    a = ans if show_ans(meta) and ans else "&nbsp;&nbsp;&nbsp;"
    return f'<span class="paren">（<b>{a}</b>）</span>'


def blk_choice(lines, meta, arg):
    """:::选择 [印]  1. 题干（ ）  /  A. … *B. …"""
    qs, cur = [], None
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        m = re.match(r"^(\*?)([A-H])[.．、]\s*(.*)", s)
        if m and cur is not None:
            cur["opts"].append((m.group(1) == "*", m.group(2), m.group(3)))
        else:
            m2 = re.match(r"^(\d{1,2})[.．、]\s*(.*)", s)
            cur = {"n": m2.group(1) if m2 else "", "stem": m2.group(2) if m2 else s, "opts": []}
            qs.append(cur)
    out = []
    for q in qs:
        ans = "".join(L for ok, L, _ in q["opts"] if ok)
        stem = inline(q["stem"], meta)
        p = paren(ans, meta)
        stem = re.sub(r"（\s*）|\(\s*\)", p, stem, count=1) if re.search(r"（\s*）|\(\s*\)", stem) else stem + p
        longest = max((len(t) for _, _, t in q["opts"]), default=0)
        cols = 4 if longest <= 6 else 2 if longest <= 14 else 1
        opts = "".join(f'<div class="op{" ok" if ok and show_ans(meta) else ""}"><span class="ol">{L}.</span>{inline(t, meta)}</div>'
                       for ok, L, t in q["opts"])
        num = f'<span class="qn">{q["n"]}.</span>' if q["n"] else ""
        out.append(f'<div class="qz{" prt-q" if arg == "印" else ""}"><div class="q">{num}{stem}</div><div class="opts c{cols}">{opts}</div></div>')
    return "".join(out)


def blk_tf(lines, meta, arg):
    out = []
    for i, ln in enumerate(x.strip() for x in lines if x.strip()):
        m = re.match(r"^(.*?)\s*(√|×|✓|✗|对|错)?$", ln)
        body, a = m.group(1), m.group(2) or ""
        a = {"对": "√", "✓": "√", "错": "×", "✗": "×"}.get(a, a)
        body = re.sub(r"^\d{1,2}[.．、]\s*", "", body)
        out.append(f'<div class="row tf"><span class="qn">{i + 1}.</span><span class="tx">{inline(body, meta)}</span>{paren(a, meta)}</div>')
    return "".join(out)


def blk_match(lines, meta, arg):
    pairs, ans = [], {}
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        m = re.match(r"^答案[:：]\s*(.*)", s)
        if m:
            for a, b in re.findall(r"(\d+)\s*[-—]\s*([A-Za-z])", m.group(1)):
                ans[int(a)] = ord(b.upper()) - 64
            continue
        l, _, r = s.partition("|")
        pairs.append((l.strip(), r.strip()))
    n, rh = len(pairs), 66
    lines_svg = ""
    if show_ans(meta):
        for a, b in ans.items():
            y1, y2 = (a - .5) * rh, (b - .5) * rh
            lines_svg += f'<line x1="41%" y1="{y1}" x2="59%" y2="{y2}" />'
    left = "".join(f'<div class="mi"><span>{i + 1}. {inline(l, meta)}</span><i></i></div>' for i, (l, _) in enumerate(pairs))
    right = "".join(f'<div class="mi"><i></i><span>{chr(65 + i)}. {inline(r, meta)}</span></div>' for i, (_, r) in enumerate(pairs))
    return (f'<div class="match" style="height:{n * rh}px;--rh:{rh}px"><svg width="100%" height="{n * rh}">{lines_svg}</svg>'
            f'<div class="ml">{left}</div><div class="mr">{right}</div></div>')


def blk_short(lines, meta, arg):
    n = int(arg) if arg.isdigit() else 2
    q, a = [], []
    for ln in lines:
        s = ln.strip()
        if re.match(r"^(答|参考答案)[:：]", s):
            a.append(re.split(r"[:：]", s, maxsplit=1)[1])
        elif s:
            q.append(s)
    qh = "".join(line_block(x, meta) for x in q)
    body = inline("".join(a), meta) if show_ans(meta) else ""
    return f'{qh}<div class="ans" style="height:calc(var(--lh) * {n})"><span class="lab">答：</span>{body}</div>'


def blk_flow(lines, meta, arg):
    st = [x.strip() for x in lines if x.strip()]
    out = []
    for i, s in enumerate(st):
        t, _, d = s.partition("|")
        c = BRANCH[i % len(BRANCH)]
        out.append(f'<div class="st" style="--c:{c}"><span class="sn">{i + 1}</span><div><div class="stt">{inline(t.strip(), meta)}</div>'
                   f'{f"<div class=std>{inline(d.strip(), meta)}</div>" if d else ""}</div></div>')
        if i < len(st) - 1:
            out.append('<span class="far">→</span>' if arg != "竖" else '<span class="far">↓</span>')
    return f'<div class="flow{" v" if arg == "竖" else ""}">{"".join(out)}</div>'


def blk_timeline(lines, meta, arg):
    out = []
    for i, s in enumerate(x.strip() for x in lines if x.strip()):
        t, _, d = s.partition("|")
        out.append(f'<div class="ev" style="--c:{BRANCH[i % len(BRANCH)]}"><div class="yr">{inline(t.strip(), meta)}</div><div class="et">{inline(d.strip(), meta)}</div></div>')
    return f'<div class="tl">{"".join(out)}</div>'


def blk_mind(lines, meta, arg):
    br, cur = [], None
    for ln in lines:
        if not ln.strip():
            continue
        if ln.startswith(" "):
            if cur is not None:
                cur[1].append(ln.strip())
        else:
            cur = (ln.strip(), [])
            br.append(cur)
    half = (len(br) + 1) // 2

    def b(i, t, kids):
        c = BRANCH[i % len(BRANCH)]
        items = "".join(f"<div>{inline(k, meta)}</div>" for k in kids)
        return f'<div class="mm-b" style="--c:{c}"><div class="bt">{inline(t, meta)}</div><div class="bi">{items}</div></div>'
    L = "".join(b(i, t, k) for i, (t, k) in enumerate(br[:half]))
    R = "".join(b(i + half, t, k) for i, (t, k) in enumerate(br[half:]))
    return f'<div class="mm"><div class="mm-col l">{L}</div><div class="mm-c">{inline(arg, meta)}</div><div class="mm-col r">{R}</div></div>'


def blk_memo(lines, meta, arg):
    inner = "".join(f"<div>{inline(x.strip(), meta)}</div>" for x in lines if x.strip())
    return f'<div class="memo{" " + arg if arg in ("粉", "蓝", "绿") else ""}">{inner}</div>'


def blk_quote(lines, meta, arg):
    inner = "".join(f"<div>{inline(x.strip(), meta)}</div>" for x in lines if x.strip())
    return f'<div class="quote"><div class="qt">{inner}</div>{f"<div class=qs>{inline(arg, meta)}</div>" if arg else ""}</div>'


def blk_concept(lines, meta, arg):
    meta["_cc"] = meta.get("_cc", 0) + 1
    c = CARD_COLORS[(meta["_cc"] - 1) % len(CARD_COLORS)]
    inner = "".join(f'<div class="row">{inline(x.strip(), meta)}</div>' for x in lines if x.strip())
    return f'<div class="concept" style="--c:{c}"><span class="ctag">{inline(arg, meta)}</span>{inner}</div>'


def blk_vocab(lines, meta, arg):
    out = []
    for s in (x.strip() for x in lines if x.strip()):
        f = [c.strip() for c in split_cells(s)] + [""] * 5
        w, ph, pos, mean, ex = f[:5]
        exh = f'<div class="vx">{inline(ex, meta)}</div>' if ex else ""
        out.append(f'<div class="vw"><div class="vh"><span class="wd">{html.escape(w)}</span><span class="ph">{html.escape(ph)}</span>'
                   f'{f"<span class=pos>{html.escape(pos)}</span>" if pos else ""}<span class="mn">{inline(mean, meta)}</span></div>{exh}</div>')
    return f'<div class="vocab">{"".join(out)}</div>'


def blk_vs(lines, meta, arg):
    """:::辨析 A | B   每行 左 | 右"""
    a, _, b = arg.partition("|")
    rows = "".join(f'<div class="vr"><div>{inline(l.strip(), meta)}</div><div>{inline(r.strip(), meta)}</div></div>'
                   for l, _, r in (x.strip().partition("|") for x in lines if x.strip()))
    return f'<div class="vsx"><div class="vr vhd"><div>{inline(a.strip(), meta)}</div><div>{inline(b.strip(), meta)}</div></div>{rows}</div>'


def find_fig(name, meta):
    for d in (Path(meta.get("_dir", ".")) / "图", NOTE / "插图"):
        for f in [d / f"{name}.svg", *d.glob(f"**/{name}.svg")] if d.exists() else []:
            if f.exists():
                return f
    raise SystemExit(f"找不到插图 {name}.svg（放在手写稿同目录的 图/ 或 assets/手写笔记/插图/ 下）")


def blk_fig(lines, meta, arg):
    """:::图 名字 [右|左|整] [宽=60] [挖空]   块内：「图名：…」一行 + 其余行是图旁文字（左图右文）。
    SVG 里 class="lab" 的 <text> 是标注：用手写字体；挖空时换成 ①② 编号（显示答案: 是 时不挖）。"""
    w = arg.split()
    name = w[0]
    side = "右" if "右" in w else ("左" if "左" in w else "整")
    pct = next((int(x[2:]) for x in w if x.startswith("宽=")), 100 if side == "整" else 56)
    blank = "挖空" in w and not show_ans(meta)
    svg = find_fig(name, meta).read_text(encoding="utf-8")
    svg = re.sub(r"<\?xml.*?\?>", "", svg).strip()
    vb = re.search(r'viewBox="([\d.\s-]+)"', svg)
    vw, vh = [float(x) for x in vb.group(1).split()[2:4]]
    if blank:
        n = [0]

        def num(m):
            n[0] += 1
            return f'{m.group(1)}<tspan class="no">{"①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"[n[0] - 1]}</tspan><tspan class="ln">＿＿＿＿</tspan></text>'
        svg = re.sub(r'(<text[^>]*class="lab[^"]*"[^>]*>)[^<]*</text>', num, svg)
        if re.search(r'class="lab[^"]*"[^>]*text-anchor="end"|text-anchor="end"[^>]*class="lab', svg):   # 左侧标注换成「①＿＿」会变长，左右各留出空白
            pad = vw * .14
            vb0 = [float(x) for x in vb.group(1).split()]
            svg = svg.replace(vb.group(0), f'viewBox="{vb0[0] - pad:.0f} {vb0[1]:.0f} {vw + 2 * pad:.0f} {vh:.0f}"', 1)
            vw += 2 * pad
    LH, BW = 56, 960
    fw = BW * pct / 100
    k = vw / fw                                                  # 标注字号固定为 26px，不随图缩放
    svg = re.sub(r"<svg\b", f'<svg class="figsvg" preserveAspectRatio="xMidYMid meet" style="--labfs:{26 * k:.1f}px"', svg, count=1)
    cap = next((l.strip()[3:] for l in lines if re.match(r"^\s*图名[:：]", l)), "")
    rest = [l for l in lines if not re.match(r"^\s*图名[:：]", l)]
    rows = -(-(fw * vh / vw + (34 if cap else 0)) // LH)       # 吸附到整行
    caph = f'<div class="figcap">{inline(cap, meta)}</div>' if cap else ""
    fig = f'<div class="figbox" style="width:{pct}%;height:calc(var(--lh) * {int(rows)})">{svg}{caph}</div>'
    if side == "整" or not any(l.strip() for l in rest):
        return f'<div class="fig {"c" if side == "整" else side}">{fig}</div>'
    text = f'<div class="figtext">{flow(rest, meta)}</div>'
    return f'<div class="fig2 {side}">{fig if side == "左" else text}{text if side == "左" else fig}</div>'


BLOCKS = {"框": blk_box, "印": blk_print, "例题": blk_question, "竖式": blk_pre, "表": blk_table, "树": blk_tree,
          "卡片": blk_card, "目录": blk_toc, "文言": blk_wenyan, "封面": blk_cover,
          "选择": blk_choice, "判断": blk_tf, "连线": blk_match, "简答": blk_short, "流程": blk_flow, "时间轴": blk_timeline,
          "发散": blk_mind, "便签": blk_memo, "金句": blk_quote, "概念": blk_concept, "单词": blk_vocab, "辨析": blk_vs, "图": blk_fig}



def page_opts(lines):
    """页首的 @纸张: 网格、@留白 这类行：只对这一页生效。"""
    opts, rest = {}, []
    for ln in lines:
        m = re.match(r"^@(\S+?)(?:[:：]\s*(.*))?$", ln.strip())
        if m and not rest and m.group(1) in ("纸张", "留白", "底色"):
            opts[m.group(1)] = (m.group(2) or "是").strip()
        elif ln.strip() or rest:
            rest.append(ln)
    return opts, rest


def page_units(lines, meta):
    """把一页源码拆成「单元」：一个块或一行；返回 [(html, 是否标题)]，自动分页按单元装页。"""
    lines = [re.sub(r"\s*〔S\d+[^〕]*〕", "", ln) for ln in lines]
    out, i = [], 0
    while i < len(lines):
        ln = lines[i]
        m = re.match(r"^:::(\S+)\s*(.*)$", ln.strip())
        if m and m.group(1) in BLOCKS:
            j = i + 1
            while j < len(lines) and lines[j].strip() != ":::":
                j += 1
            out.append((BLOCKS[m.group(1)](lines[i + 1:j], meta, m.group(2).strip()), False))
            i = j + 1
            continue
        h = line_block(ln, meta)
        if h:
            out.append((h, bool(re.match(r"^#{1,4}\s|^#>\s", ln.strip()))))
        i += 1
    return out


def wrap(units):
    return "".join(f'<div class="u{" kh" if hd else ""}">{h}</div>' for h, hd in units)


def page_html(lines, meta):
    return wrap(page_units(lines, meta))


def paginate(sections, meta, tmp):
    """sections = [(opts, units)]。量出每个单元的高度，按页面可用高度贪心装页；标题不落页尾。"""
    W, H = SIZES.get(meta.get("尺寸", "A4"), SIZES["A4"])
    probe = [(i + 1, wrap(u), o) for i, (o, u) in enumerate(sections)]
    f = tmp / "测量.html"
    f.write_text(doc(meta, probe, len(probe)).replace("</body>", MEASURE_JS + "</body>"), encoding="utf-8")
    dom = chrome(["--virtual-time-budget=15000", "--dump-dom", f.as_uri()]).stdout
    m = re.search(r'data-measure="([^"]*)"', dom)
    data = json.loads(html.unescape(m.group(1)))
    pages = []
    for (o, units), (cap, hs) in zip(sections, data):
        cur = []                                     # [(unit, h)]
        for u, h in zip(units, hs):
            room = cap - sum(x[1] for x in cur)
            if cur and h > room + 1 and 'class="fig' in u[0]:
                s = shrink_fig(u[0], room)              # 插图缩到 ≥65% 能放下，就缩一点放在这页
                if __import__("os").environ.get("HN_DEBUG"): print("shrink?", int(h), int(room), bool(s), file=sys.stderr)
                if s:
                    u, h = (s[0], u[1]), s[1]
            if cur and sum(x[1] for x in cur) + h > cap + 1:
                carry = []
                while cur and cur[-1][0][1]:             # 页尾的标题跟着下一块走
                    carry.insert(0, cur.pop())
                if cur:
                    pages.append((o, [x[0] for x in cur]))
                cur = carry
            cur.append((u, h))
        if cur:
            pages.append((o, [x[0] for x in cur]))
    return pages


def shrink_fig(h, room, LH=56, BW=960):
    m = re.search(r'class="figbox" style="width:(\d+(?:\.\d+)?)%;height:calc\(var\(--lh\) \* (\d+)\)"', h)
    vb = re.search(r'viewBox="[\d.\s-]+?\s([\d.]+)\s([\d.]+)"', h)
    if not m or not vb or 'class="fig2' in h:
        return None
    pct, rows = float(m.group(1)), int(m.group(2))
    vw, vh = float(vb.group(1)), float(vb.group(2))
    capx = 34 if 'class="figcap"' in h else 0
    nrows = int((room - LH) // LH)                       # 上下各半行
    if nrows < 2:
        return None
    npct = (nrows * LH - capx) * vw / vh / BW * 100
    if npct >= pct or npct < pct * .75:
        return None
    k = vw / (BW * npct / 100)
    h = h.replace(m.group(0), f'class="figbox" style="width:{npct:.1f}%;height:calc(var(--lh) * {nrows})"')
    h = re.sub(r"--labfs:[\d.]+px", f"--labfs:{26 * k:.1f}px", h)
    return h, (nrows + 1) * LH


MEASURE_JS = ("<script>document.fonts.ready.then(()=>{setTimeout(()=>{const r=[];document.querySelectorAll('.page').forEach(p=>{const b=p.querySelector('.body');"
              "r.push([b.clientHeight,[...b.querySelectorAll(':scope>.u')].map(u=>u.getBoundingClientRect().height)]);});"
              "document.body.setAttribute('data-measure',JSON.stringify(r));},300)});</script>")


# ---------------- 页面 ----------------
def parse(md):
    text = re.sub(r"<!--.*?-->", "", Path(md).read_text(encoding="utf-8"), flags=re.S)
    meta = {}
    m = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    if m:
        for ln in m.group(1).splitlines():
            if re.match(r"^\s*[^#][^:：]*[:：]", ln):
                k, v = re.split(r"[:：]", ln, maxsplit=1)
                meta[k.strip()] = v.strip()
        text = text[m.end():]
    base = LAYOUTS.get(meta.get("版式", "横线本"), LAYOUTS["横线本"])
    for k, v in base.items():
        meta.setdefault(k, v)
    meta["_dir"] = Path(md).resolve().parent
    if meta.get("答案") == "章末":
        text = answers_to_end(text, meta)
    pages = [p.split("\n") for p in re.split(r"\n===\s*\n", "\n" + text + "\n")]
    pages = [p for p in pages if any(x.strip() for x in p)]
    return meta, pages


def answers_to_end(text, meta):
    """头部 `答案: 章末`：{空|答案} 换成带编号的空，挖空图的标注也收集起来，章末自动加「参考答案」。"""
    groups, order, head, n = {}, [], "", [0]
    out = []

    def add(key, item):
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(item)

    for ln in text.split("\n"):
        h = re.match(r"^#{1,3}\s+(.+)$", ln)
        if h:
            head = re.sub(r"〔[^〕]*〕", "", h.group(1)).strip()
            head = re.sub(r"^[一二三四五六七八九十]+、|^\d+[、.]\s*", "", head)
            head = re.sub(r"^常考[：:]", "", head).rstrip("？?")

        def sub(m):
            n[0] += 1
            add(head, f"({n[0]}) {m.group(1)}")
            return "{空#" + str(n[0]) + "|" + m.group(1) + "}"
        ln = re.sub(r"\{空\|([^{}]*)\}", sub, ln)
        f = re.match(r"^:::图\s+(\S+)(.*)$", ln)
        if f and "挖空" in f.group(2):
            svg = find_fig(f.group(1), meta).read_text(encoding="utf-8")
            labs = re.findall(r'<text[^>]*class="lab[^"]*"[^>]*>([^<]*)</text>', svg)
            add(f"识图：{f.group(1)}", "　".join(f"{'①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'[i]}{t}" for i, t in enumerate(labs)))
        out.append(ln)
    if order:
        out += ["", "# 参考答案"]
        out += [f"- **{k}**　" + "　".join(groups[k]) for k in order]
    return "\n".join(out)


def font_of(key):
    if key in FONTS:
        return FONTS[key]
    for k, v in FONTS.items():
        if not k.startswith("_") and v["name"] == key:
            return v
    raise SystemExit(f"字体 {key} 不在 assets/手写笔记/字体表.json里，可选：" + "、".join(f"{k} {v['name']}" for k, v in FONTS.items() if not k.startswith("_")))


SIZES = {"A4": (1080, 1527), "小红书": (1080, 1440)}       # A4 = 210×297 等比；打印 PDF 时整体缩到 A4
TINTS = {"奶油": "#fffdf5", "纯白": "#ffffff", "米黄": "#f8f4ea", "薄荷": "#f1faf4", "樱粉": "#fff5f6",
         "天蓝": "#f2f7fe", "牛皮": "#efe4cf", "浅灰": "#f6f6f4"}


def deco_html(meta, paper, n):
    head = meta.get("页眉", "")
    out = ""
    if paper == "表格框":
        out += f'<div class="frame"></div><div class="tab">{html.escape(head)}</div>'
    elif head:
        l, _, r = head.partition("|")
        out += f'<div class="header"><span>{html.escape(l.strip())}</span><span>{html.escape(r.strip())}</span></div>'
    if paper == "康奈尔":
        out += f'<div class="cue"><span>{html.escape(meta.get("线索栏", "关键词 / 问题"))}</span></div><div class="summary"><span>总结</span></div>'
    for d in re.split(r"[、,，\s]+", meta.get("装饰", "")):
        if d == "活页孔":
            out += '<div class="holes">' + "<i></i>" * 11 + "</div>"
        elif d == "胶带":
            out += '<div class="tape l"></div><div class="tape r"></div>'
        elif d == "页角":
            out += '<div class="corner"></div>'
    return out


def doc(meta, pages_html, total, print_mode=False):
    f = font_of(meta["字体"])
    fp = (NOTE / f["file"]).resolve().as_uri()
    css = (NOTE / "样式.css").resolve().as_uri()
    W, H = SIZES.get(meta.get("尺寸", "A4"), SIZES["A4"])
    extra = [f"--W:{W}px", f"--H:{H}px"]
    if meta.get("正文字号"):
        extra.append(f"--fs:{meta['正文字号']}px")
    if meta.get("行高"):
        extra.append(f"--lh:{meta['行高']}px")
    paper = meta["纸张"]
    tint = TINTS.get(meta.get("底色", ""), "")
    pstyle = f' style="--paper:{tint}"' if tint else ""
    out = []
    for n, body, *po in pages_html:
        po = po[0] if po else {}
        paper = po.get("纸张", meta["纸张"])
        tint = TINTS.get(po.get("底色", meta.get("底色", "")), "")
        pstyle = (f' style="--paper:{tint}"' if tint else "") + (" data-open" if "留白" in po else "")
        pm = meta["页码"]
        foot = ""
        if pm == "-n-":
            foot = f'<div class="footer">-{n}-</div>'
        elif pm == "n / N":
            foot = f'<div class="footer right">{n:02d} / {total:02d}</div>'
        out.append(f'<section class="page paper-{paper}"{pstyle}>{deco_html(meta, paper, n)}<div class="body">{body}</div>{foot}</section>')
    zoom = 793.5 / W
    pr = (f"<style>@page{{size:A4;margin:0}}html,body{{background:none}}</style>"
          f"<script>document.fonts.ready.then(()=>{{document.documentElement.style.zoom={zoom:.5f}}})</script>") if print_mode else ""
    return (f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><link rel="stylesheet" href="{css}">'
            f'<style>@font-face{{font-family:"HW";src:url("{fp}");size-adjust:{f["sa"]}%}}'
            f'body{{--ls:{f["ls"]}px;--sk:{f["sk"]}px;{";".join(extra)}}}</style></head><body>{"".join(out)}'
            '<script>' + MM_JS + CHECK_JS + 'document.fonts.ready.then(()=>{mm();check();});</script>' + pr +
            "</body></html>")


# 自动检查：超出、标题落页尾、字体是否加载、每页强调 / 大块 / 提示数量 → body[data-check]
CHECK_JS = ("function check(){const r=[];const hw=[...document.fonts].some(f=>f.family.replace(/\"/g,'')=='HW'&&f.status=='loaded');"
            "document.querySelectorAll('.page').forEach((p,i)=>{const b=p.querySelector('.body');if(!b)return;"
            "const q=s=>b.querySelectorAll(s).length;const last=b.lastElementChild;const lastIn=last&&last.classList.contains('u')?last.lastElementChild:last;"
            "const big=[...b.querySelectorAll('table.t')].filter(t=>t.rows.length>=4).length+q('.tree,.mm,.flow,.tl,.vocab');"
            "const bt=b.getBoundingClientRect().top;const lb=last?last.getBoundingClientRect().bottom:bt;r.push({p:i+1,fill:Math.round(100*(lb-bt)/b.clientHeight),open:p.hasAttribute('data-open'),over:Math.max(0,b.scrollHeight-b.clientHeight-4),tail:!!(last&&(last.classList.contains('kh')||/\\bh[234]\\b/.test(last.className))),"
            "red:q('.r'),wave:q('.wave'),hl:q('.hl'),big:big,hint:q('.memo,.quote,.concept,.boxed,.vsx,.sticker,.stamp')});});"
            "document.body.setAttribute('data-check',JSON.stringify({hw:hw,pages:r}));}")
LIMITS = {"red": (8, "红字"), "wave": (5, "波浪线"), "hl": (4, "荧光"), "hint": (2, "提示类组件（框/便签/金句/概念/贴纸/印章）"), "big": (1, "大块组件（大表/导图/流程/时间轴/词表）")}


MM_JS = ("function mm(){const NS='http://www.w3.org/2000/svg';document.querySelectorAll('.mm').forEach(m=>{"
         "const r=m.getBoundingClientRect(),c=m.querySelector('.mm-c').getBoundingClientRect();const svg=document.createElementNS(NS,'svg');"
         "svg.setAttribute('class','mm-svg');svg.setAttribute('width',r.width);svg.setAttribute('height',r.height);"
         "m.querySelectorAll('.mm-b').forEach(b=>{const q=b.querySelector('.bt').getBoundingClientRect();const L=q.right<=c.left+2;"
         "const x1=(L?c.left:c.right)-r.left,y1=c.top+c.height/2-r.top,x2=(L?q.right:q.left)-r.left,y2=q.top+q.height/2-r.top,mx=(x1+x2)/2;"
         "const p=document.createElementNS(NS,'path');p.setAttribute('d',`M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`);"
         "p.setAttribute('stroke',getComputedStyle(b).getPropertyValue('--c'));p.setAttribute('stroke-width','3');p.setAttribute('fill','none');"
         "p.setAttribute('stroke-linecap','round');svg.appendChild(p);});m.prepend(svg);});}")


def chrome(args):
    return subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--allow-file-access-from-files", *args],
                          capture_output=True, text=True, encoding="utf-8", errors="replace")


def source_gaps(pages):
    """项目模式（稿子里出现过〔Sxx〕）下，例题、练习、材料块要标出处。"""
    text = "\n".join("\n".join(p) for p in pages)
    if "〔S" not in text:
        return []
    gaps = []
    for pi, p in enumerate(pages, 1):
        i = 0
        while i < len(p):
            ln = p[i].strip()
            m = re.match(r"^:::(\S+)", ln)
            if m and m.group(1) in ("例题", "选择", "判断", "连线", "简答", "印", "卡片", "文言", "单词", "金句"):
                j = i + 1
                while j < len(p) and p[j].strip() != ":::":
                    j += 1
                blk = "\n".join(p[max(0, i - 1):j + 1])
                if "〔S" not in blk and "【缺】" not in blk:
                    gaps.append(f"第 {pi} 页 :::{m.group(1)} 块")
                i = j + 1
                continue
            if re.match(r"^例\d*[:：]", ln) and "〔S" not in ln and "【缺】" not in ln:
                gaps.append(f"第 {pi} 页「{ln[:16]}」")
            i += 1
    return gaps


def render(md, out=None, font=None, scale=2, pick=None, pdf=True):
    meta, pages = parse(md)
    meta["_dir"] = str(Path(md).resolve().parent)
    if font:
        meta["字体"] = font
    out = (Path(out) if out else Path(md).with_name(Path(md).stem + "_图")).resolve()
    (out / "_html").mkdir(parents=True, exist_ok=True)
    start = int(meta.get("起始页码", 1))
    bodies = []
    if meta.get("自动分页", "否") == "是":
        secs = []
        for p in pages:
            o, rest = page_opts(p)
            secs.append((o, page_units(rest, meta)))
        for i, (o, us) in enumerate(paginate(secs, meta, out / "_html")):
            bodies.append((i + start, wrap(us), o))
        pages = [p for p in pages]            # 出处检查仍按源码分段
    else:
        for i, p in enumerate(pages):
            o, rest = page_opts(p)
            bodies.append((i + start, page_html(rest, meta), o))
    total = len(bodies)
    W, H = SIZES.get(meta.get("尺寸", "A4"), SIZES["A4"])
    full = out / "_html" / "全部.html"
    full.write_text(doc(meta, bodies, total), encoding="utf-8")
    dom = chrome(["--virtual-time-budget=15000", "--dump-dom", full.as_uri()]).stdout
    m = re.search(r"data-check=\"([^\"]*)\"", dom)
    chk = json.loads(html.unescape(m.group(1))) if m else {"hw": False, "pages": []}
    want = set(range(start, total + start))
    if pick:
        a, _, b = pick.partition("-")
        want = set(range(int(a), int(b or a) + 1))

    def shot(item):
        n = item[0]
        h = out / "_html" / f"p-{n:02d}.html"
        h.write_text(doc(meta, [item], total), encoding="utf-8")
        png = out / f"p-{n:02d}.png"
        chrome([f"--force-device-scale-factor={scale}", f"--window-size={W},{H}", "--virtual-time-budget=10000", f"--screenshot={png}", h.as_uri()])
        return png

    with ThreadPoolExecutor(4) as ex:
        files = list(ex.map(shot, [b for b in bodies if b[0] in want]))
    sheet = overview(files, out / "总览.png", W / H)
    pdf_path = None
    if pdf and meta.get("尺寸", "A4") == "A4" and not pick:
        ph = out / "_html" / "打印.html"
        ph.write_text(doc(meta, bodies, total, print_mode=True), encoding="utf-8")
        pdf_path = out / f"{Path(md).stem}.pdf"
        chrome(["--no-pdf-header-footer", "--virtual-time-budget=15000", f"--print-to-pdf={pdf_path}", ph.as_uri()])
    f = font_of(meta["字体"])

    errs, warns = [], []
    if not chk["hw"]:
        errs.append(f"手写字体 {f['name']} 没加载上（页面回落成了系统字体）：字体文件要先用 fontTools 修，见templates/手写笔记/手写笔记规范.md第一节")
    for pg in chk["pages"]:
        if pg["over"]:
            errs.append(f"第 {pg['p']} 页内容超出 {pg['over']}px：把后半段挪到下一页（加 ===），不缩字号")
        if pg.get("fill", 100) < 80 and not pg.get("open") and not pg["over"]:
            warns.append(f"第 {pg['p']} 页只排了 {pg['fill']}%：下面空得多，把下一块内容挪上来或补素材细节（章首、章末页写 @留白 可豁免）")
        if pg["tail"]:
            warns.append(f"第 {pg['p']} 页最后一行是标题：标题后面没内容，把它挪到下一页")
        for k, (lim, name) in LIMITS.items():
            if pg[k] > lim:
                warns.append(f"第 {pg['p']} 页{name} {pg[k]} 处，超过 {lim}")
    for g in source_gaps(pages):
        warns.append(f"{g}没标出处〔Sxx pN〕（素材里没有就写【缺】）")
    sys.path.insert(0, str(HERE))
    import author   # 原作者信息：手写稿全文 + 稿名都查
    mdp = Path(md).resolve()
    import paths
    proj = paths.find_root(mdp)
    terms = author.project_terms(proj)
    for lv, word, ctx in author.scan(mdp.name + "\n" + re.sub(r"〔[^〕]*〕", "", mdp.read_text(encoding="utf-8")), terms):
        (errs if lv == "error" else warns).append(f"原作者信息「{word}」：{ctx}" + ("" if lv == "error" else "（正常教学内容可保留）"))
    if not terms and paths.P(proj).素材.exists():
        warns.append(f"{author.TERMS_FILE} 的屏蔽词区是空的：原作者的具体名字没查")
    rep = [f"# 手写笔记检查报告：{Path(md).name}", "",
           f"- {total} 页 · 版式 {meta.get('版式', '横线本')} · 纸张 {meta['纸张']} · 字体 {f['name']} · 尺寸 {meta.get('尺寸', 'A4')}",
           f"- 结论：{'❌ 有错误要改' if errs else ('⚠️ 可以交付，提醒项逐条看' if warns else '✅ 可以交付')}", ""]
    rep += ["## 错误", *[f"- {e}" for e in errs], ""] if errs else []
    rep += ["## 提醒", *[f"- {w}" for w in warns], ""] if warns else []
    rep += ["## 各页填充率", " · ".join(f"{pg['p']}:{pg.get('fill', 0)}%" for pg in chk["pages"]), ""]
    rep += ["## 还要人工看的（templates/手写笔记/手写笔记规范.md第八节）", "- 每页逐张看图：字体、横线对齐、颜色意思、例题答案和素材一致、文言注释不压字"]
    (out / "检查报告.md").write_text("\n".join(rep) + "\n", encoding="utf-8")

    print(f"{len(files)} 页 · 版式 {meta.get('版式', '横线本')} · 字体 {f['name']} -> {out}")
    if sheet:
        print("  总览", sheet)
    if pdf_path:
        print("  打印 PDF（A4）", pdf_path)
    for e in errs:
        print("  ❌", e)
    for w in warns:
        print("  ⚠️", w)
    print("  检查报告", out / "检查报告.md")
    return files, errs


def overview(files, dest, ratio=0.75, cols=4):
    try:
        from PIL import Image
    except ImportError:
        return None
    if not files:
        return None
    w = 360
    h = int(w / ratio)
    rows = (len(files) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (w + 16) + 16, rows * (h + 16) + 16), "white")
    for i, f in enumerate(files):
        im = Image.open(f).convert("RGB").resize((w, h))
        sheet.paste(im, (16 + (i % cols) * (w + 16), 16 + (i // cols) * (h + 16)))
    sheet.save(dest)
    return dest


def init(dest, kind, name=None):
    """复制一份版式样张当起点：init <项目或目录> <样张编号 01–23> [文件名]"""
    src = next(iter(sorted((NOTE / "样张").glob(f"{int(kind):02d}_*.md"))), None)
    if not src:
        raise SystemExit("样张编号 01–23，见 assets/手写笔记/样张/")
    sys.path.insert(0, str(HERE))
    import paths
    d = Path(dest)
    P = paths.P(d)
    if P.内容稿.exists() or P.素材清单.exists() or (d / "1_规划").exists():   # 传的是资料包项目 → 放进 2_样张/手写笔记/
        d = P.手写笔记
    d.mkdir(parents=True, exist_ok=True)
    n = len(list(d.glob("*.md"))) + 1
    t = d / f"{n:02d}_{name or src.stem.split('_', 1)[1]}.md"
    body = src.read_text(encoding="utf-8")
    body = re.sub(r"<!--.*?-->", "<!-- 从样张 " + src.name + " 复制。内容只取素材，每个例题 / 练习 / 材料块标〔Sxx pN〕，没有的写【缺】 -->", body, count=1, flags=re.S)
    t.write_text(body, encoding="utf-8")
    print(t)
    return t


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "init":
        ap = argparse.ArgumentParser(prog="handnote.py init")
        ap.add_argument("dest")
        ap.add_argument("kind")
        ap.add_argument("name", nargs="?")
        a = ap.parse_args(sys.argv[2:])
        init(a.dest, a.kind, a.name)
        sys.exit(0)
    ap = argparse.ArgumentParser()
    ap.add_argument("md", nargs="+")
    ap.add_argument("--out")
    ap.add_argument("--font")
    ap.add_argument("--scale", type=float, default=2)
    ap.add_argument("--pages")
    ap.add_argument("--no-pdf", action="store_true")
    a = ap.parse_args()
    bad = 0
    for md in a.md:          # 可一次传多份手写稿；多份时 --out 是父目录
        out = (Path(a.out) / Path(md).stem) if a.out and len(a.md) > 1 else a.out
        _, errs = render(md, out, a.font, a.scale, a.pages, not a.no_pdf)
        bad += bool(errs)
    sys.exit(1 if bad else 0)
