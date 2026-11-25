"""内容稿.md → Word 母稿（.docx）。格式全部写进 Word 样式，参数来自排版参数.json，标记见templates/考点答题模板/标记写法.md。

  python3 build_docx.py <内容稿.md> <输出.docx> [--variant 完整版] [--pages pages.json] [--show-sources] [--allow-gaps]

variant：完整版 / 考点模板 / 练习题 / 参考答案 / 背诵默写
pages.json：目录项 → 页码（run.py 第一遍导出 PDF 后回填）
"""
import argparse
import json
import math
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

SKILL = Path(__file__).resolve().parent.parent
STYLE = json.loads((SKILL / "config" / "排版参数.json").read_text(encoding="utf-8"))
C = dict(STYLE["colors"])
F = dict(STYLE["fonts"])
if sys.platform == "win32":   # Windows 没有 Songti SC / Kaiti SC，换系统自带的宋体、楷体（排版参数.json fonts.windows）
    F.update(F.get("windows", {}))
E = dict(STYLE["elements"])
THEMES = {k: v for k, v in STYLE.get("themes", {}).items() if not k.startswith("_")}
THEME = {}
# 新元素在没有主题时也要有样式（用知识清单的取值兜底）
for _t in THEMES.values():
    for _k, _v in _t.get("elements", {}).items():
        E.setdefault(_k, _v)
    for _k, _v in _t.get("colors", {}).items():
        C.setdefault(_k, _v)
BASE_C, BASE_E = dict(C), dict(E)


def apply_theme(kind):
    """按内容稿「类型」套主题：覆盖颜色和元素样式（见排版参数.json themes）。"""
    THEME.clear()
    C.clear(); C.update(BASE_C)
    E.clear(); E.update(BASE_E)
    t = THEMES.get(kind or "")
    if t:
        THEME.update(t)
        C.update(t.get("colors", {}))
        E.update(t.get("elements", {}))
PG = STYLE["page"]
PAGE_W = Cm(21.0)
TEXT_W = Emu(PAGE_W - Cm(PG["margin_left_cm"]) - Cm(PG["margin_right_cm"]))
COMPACT_ANS = False
VARIANTS = ["完整版", "考点模板", "练习题", "参考答案", "背诵默写"]

# 栏目标签 → (标签底色, 后续行样式)
LABELS = {
    "常见问法": ("label_问法", "问法"),
    "答题思路": ("label_思路", "思路"),
    "答题模板": ("label_模板", "模板"),
    "解析": ("label_解析", "讲解正文"),
    "知识点": ("label_问法", "讲解正文"),
    "易错点": ("label_思路", "讲解正文"),
    "方法": ("label_模板", "思路"),
    # 知识清单类
    "课本原句": ("label_问法", "讲解正文"),
    "口诀": ("label_模板", "模板"),
    "拓展": ("label_解析", "讲解正文"),
    "考点": ("label_模板", "讲解正文"),   # 中考怎么考：!小标签：内容
    "考法": ("label_模板", "讲解正文"),   # 中考总复习：## 考点 下面的「中考怎么考」，避免和 ## 考点 重名
}
DEFAULT_LABEL = ("label_解析", "讲解正文")

# ---------------- 解析 ----------------
SRC_RE = re.compile(r"〔[^〕]*〕")


def parse(path):
    text = Path(path).read_text(encoding="utf-8")
    text = re.sub(r"<!--.*?-->", "", text, flags=re.S)
    meta = {}
    if text.startswith("---"):
        head, _, text = text[3:].partition("\n---")
        for line in head.strip().splitlines():
            if ":" in line:
                k, v = line.split(":", 1)
                meta[k.strip()] = v.strip()
    blocks, zone, label = [], "front", None
    table = None
    for raw in text.splitlines():
        line = raw.rstrip()
        s = line.strip()
        if table is not None:
            if s == "[/表]":
                blocks.append({"kind": "table", "text": "", "zone": zone, "label": label, **table})
                table = None
            elif s:
                table["rows"].append([c.strip() for c in re.split(r"\s*[|｜]\s*", s)])
            continue
        if not s:
            continue

        def add(kind, txt=s, **kw):
            blocks.append({"kind": kind, "text": txt, "zone": zone, "label": label, **kw})

        if re.fullmatch(r"\[表(?: (?:无表头|首列))*\]", s):   # [表] / [表 无表头] / [表 首列]
            table = {"rows": [], "head": "无表头" not in s, "first": "首列" in s}
        elif s.startswith("#### ") and zone != "answer":
            add("sub", s[5:].strip())
        elif s == "[分页]":
            add("pagebreak", "")
        elif s.startswith("[答案]"):   # 章内答案：本章练习后面换页接答案，到下一个 # 单元为止
            zone, label = "answer", None
            add("ans_head", s[4:].strip())
        elif s.startswith("[卷头]"):
            zone = "explain"
            add("exam_head", s[4:].strip())
        elif s.startswith("# "):
            title = s[2:].strip()
            zone = "answer" if SRC_RE.sub("", title).strip() == "参考答案" else "explain"
            label = None
            add("answer_start" if zone == "answer" else "unit", title)
        elif s.startswith("## "):
            if zone == "answer":
                add("ans_unit", s[3:].strip())
            else:
                zone, label = "explain", None
                add("kd", s[3:].strip())
        elif s.startswith("### "):
            if zone == "answer":
                add("ans_title", s[4:].strip())
            else:
                zone, label = "exercise", None
                add("ex_title", s[4:].strip())
        elif zone == "answer":
            if re.match(r"^(\d+[.．]|（\d+）).{0,8}考点[:：]", s):
                add("ans_q")
            elif s.startswith(("详细答题思路", "答题思路", "解析：")):
                add("ans_why")
            elif re.match(r"^(\d+[.．]|（\d+）)?\s*参考答案[:：]", s):
                add("ans_ans")  # 「1.参考答案：C」这种只有答案没有解析的行也算
            else:
                add("ans_text")
        elif s.startswith("> ") and zone == "explain":
            add("yaosu", s[2:].strip())
        elif re.fullmatch(r"【[^】]{1,8}】", s) and not s.startswith("【缺"):
            label = s[1:-1]
            add("label", label)
        elif s.startswith("|") or s.startswith("｜"):
            add("ex_para" if zone == "exercise" else "passage", s[1:].strip())
        elif s.startswith("——") and "《" in s:
            add("source")
        elif s in ("（有改动）", "(有改动)"):
            add("source")
        elif re.fullmatch(r"\[横线(\d*)\]", s):
            n = re.fullmatch(r"\[横线(\d*)\]", s).group(1)
            add("lines", "", n=int(n or 1))
        elif zone == "exercise":
            if re.match(r"^[一二三四五六七八九十]+、", s):
                add("ex_sec")
            elif re.match(r"^[（(][一二三四五六七八九十]+[)）]|^任务[一二三四五六七八九十]+[:：]", s):
                add("ex_sub")
            elif re.match(r"^[A-D][.．]", s):
                add("opts")
            else:
                add("q")
        else:  # 讲解区
            if s.startswith("例：") or s.startswith("例:"):
                add("li")
            elif s.startswith("- "):
                add("eg", s[2:].strip())
            elif s.startswith("答：") or s.startswith("答:"):
                add("da")
            elif s.startswith("!") or s.startswith("！"):
                add("pair", s[1:].strip())
            elif re.match(r"^[A-D][.．]", s):
                add("opts")
            else:
                add("body")
    return meta, blocks


def select(blocks, variant):
    if variant == "完整版":
        return blocks
    out = []
    for b in blocks:
        z, k = b["zone"], b["kind"]
        if variant == "考点模板" and (z == "explain") and k != "exam_head":
            out.append(b)
        elif variant == "练习题" and (z == "exercise" or k == "unit"):
            out.append(b)
        elif variant == "参考答案" and z == "answer":
            out.append(b)
        elif variant == "背诵默写" and z == "explain" and k in ("unit", "yaosu", "kd", "label", "body", "pair",
                                                               "sub", "eg", "table"):
            out.append(b)
    # 背诵默写版不收例文，「课本原句」标签也一并去掉
    if variant == "背诵默写":
        out = [b for b in out if not (b["kind"] == "label" and b["text"] == "课本原句")]
    # 去掉没有内容的单元条（例如练习题版里没有练习的单元）
    cleaned = []
    for i, b in enumerate(out):
        if b["kind"] == "unit":
            nxt = next((x for x in out[i + 1:]), None)
            if nxt is None or nxt["kind"] == "unit":
                continue
        cleaned.append(b)
    return cleaned


# ---------------- XML 工具 ----------------
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
             "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]


def _reorder(el, order):
    rank = {qn("w:" + n): i for i, n in enumerate(order)}
    kids = list(el)
    for k in kids:
        el.remove(k)
    for k in sorted(kids, key=lambda k: rank.get(k.tag, 999)):
        el.append(k)


def normalize(doc):
    """按 OOXML 规定顺序重排 rPr/pPr，避免 Word 报“内容有问题”。"""
    roots = [doc.element, doc.styles.element] + [s.footer._element for s in doc.sections]
    for root in roots:
        for el in root.iter(qn("w:rPr")):
            _reorder(el, RPR_ORDER)
        for el in root.iter(qn("w:pPr")):
            _reorder(el, PPR_ORDER)


def set_fonts(rpr, east, latin=None):
    latin = latin or F["latin"]
    f = rpr.find(qn("w:rFonts"))
    if f is None:
        f = OxmlElement("w:rFonts")
        rpr.append(f)
    for a in list(f.attrib):
        del f.attrib[a]
    for k, v in (("ascii", latin), ("hAnsi", latin), ("eastAsia", east), ("cs", latin)):
        f.set(qn("w:" + k), v)


def add_shd(parent, fill):
    s = OxmlElement("w:shd")
    for k, v in (("val", "clear"), ("color", "auto"), ("fill", fill)):
        s.set(qn("w:" + k), v)
    parent.append(s)


def add_run_border(rpr, color="000000", sz=6):
    b = OxmlElement("w:bdr")
    for k, v in (("val", "single"), ("sz", str(sz)), ("space", "0"), ("color", color)):
        b.set(qn("w:" + k), v)
    rpr.append(b)


def add_para_border(ppr, side, color, sz=8, space=1):
    bdr = ppr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        ppr.append(bdr)
    e = OxmlElement("w:" + side)
    for k, v in (("val", "single"), ("sz", str(sz)), ("space", str(space)), ("color", color)):
        e.set(qn("w:" + k), v)
    bdr.append(e)


def rgb(key):
    return RGBColor.from_string(C.get(key, key))


# ---------------- 样式 ----------------
STYLE_NAME = {"单元条": "Heading 1", "考点": "Heading 2", "页脚": "Footer"}


def sname(el):
    return STYLE_NAME.get(el, "资料-" + el)


def build_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.size = Pt(12)
    set_fonts(normal.element.get_or_add_rPr(), F["serif"])
    normal.paragraph_format.widow_control = True
    for el, spec in E.items():
        name = sname(el)
        try:
            st = doc.styles[name]
        except KeyError:
            st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            st.base_style = normal
        f = st.font
        f.size, f.bold, f.italic = Pt(spec["size"]), spec.get("bold", False), False
        f.color.rgb = rgb(spec.get("color", "black"))
        set_fonts(st.element.get_or_add_rPr(), F[spec.get("font", "serif")])
        pf = st.paragraph_format
        pf.space_before, pf.space_after = Pt(spec.get("before", 0)), Pt(spec.get("after", 0))
        pf.line_spacing = Pt(spec["line"]) if spec["line"] else 1.0  # line 为 0：单倍行距（带左色条的标题用，色条才贴字高）
        pf.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(
            spec.get("align"), WD_ALIGN_PARAGRAPH.JUSTIFY if el in ("例文", "练习正文") else None)
        if spec.get("indent_cm"):
            pf.left_indent = Cm(spec["indent_cm"])
        if spec.get("first_line_chars"):
            ind = OxmlElement("w:ind")
            ind.set(qn("w:firstLineChars"), str(spec["first_line_chars"] * 100))
            st.element.get_or_add_pPr().append(ind)

    for lvl in (1, 2, 3):
        doc.styles[sname(f"目录{lvl}")].paragraph_format.tab_stops.add_tab_stop(
            TEXT_W, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    h1 = doc.styles["Heading 1"]
    ppr = h1.element.get_or_add_pPr()
    add_shd(ppr, C["unit_fill"])
    add_para_border(ppr, "bottom", C["unit_line"], 8)
    h2 = doc.styles["Heading 2"]
    rpr = h2.element.get_or_add_rPr()
    set_fonts(rpr, F["serif"], F["serif"])  # 中西文同字体，标签框不在数字处断开
    if THEME.get("kd_style") == "bar":  # 通栏浅底 + 左侧粗色条 + 下边线
        ppr = h2.element.get_or_add_pPr()
        add_shd(ppr, C["theme_light"])
        add_para_border(ppr, "left", C["theme"], 36, 6)
        add_para_border(ppr, "bottom", C["theme"], 6, 2)
    else:
        add_run_border(rpr)
        add_shd(rpr, C["kd_fill"])
    for el, fill, line in (("课本句", "quote_fill", "quote_line"), ("口诀框", "rhyme_fill", "rhyme_line")):
        ppr = doc.styles[sname(el)].element.get_or_add_pPr()
        add_shd(ppr, C[fill])
        add_para_border(ppr, "left", C[line], 24 if el == "课本句" else 8, 6)
        if el == "口诀框":
            for side in ("top", "bottom", "right"):
                add_para_border(ppr, side, C[line], 8, 4)
    doc.styles[sname("知识小标题")].paragraph_format.keep_with_next = True
    doc.styles[sname("课本句")].paragraph_format.keep_with_next = True   # 课本原句块不和后面的来源行、知识点拆开
    doc.styles[sname("口诀框")].paragraph_format.keep_with_next = True   # 口诀框整块不跨页（八上 Unit 1 口诀被拆到两页）
    if THEME.get("kd_style") == "bar":
        ppr = doc.styles[sname("练习标题")].element.get_or_add_pPr()
        add_para_border(ppr, "top", C["theme"], 6, 3)
        add_para_border(ppr, "bottom", C["theme"], 6, 3)
        add_shd(ppr, C["theme_light"])
        ex_body = doc.styles[sname("练习正文")]
        ex_body_ppr = ex_body.element.get_or_add_pPr()
        add_shd(ex_body_ppr, "F7FAFD")
        add_para_border(ex_body_ppr, "left", C["table_line"], 12, 6)
        ex_body.paragraph_format.left_indent = Cm(0.35)
        ex_body.paragraph_format.right_indent = Cm(0.25)
    for el in ("单元条", "单元要素", "考点", "栏目", "例题", "例文", "练习标题", "练习大题", "练习小标题", "答案单元", "答案篇名", "答案题"):
        doc.styles[sname(el)].paragraph_format.keep_with_next = True
    ans = doc.styles[sname("作答线")]
    ans.paragraph_format.left_indent = Cm(0.6)
    ans.paragraph_format.tab_stops.add_tab_stop(TEXT_W, WD_TAB_ALIGNMENT.RIGHT)
    # 字符样式：只放颜色/底色/下划线，加粗一律在 run 上直接设（加粗是开关属性，放样式里会和段落样式抵消）
    for name, (fill_key, _) in list(LABELS.items()) + [("其他", DEFAULT_LABEL)]:
        st = doc.styles.add_style("标签-" + name, WD_STYLE_TYPE.CHARACTER)
        st.font.color.rgb = rgb("black")
        add_shd(st.element.get_or_add_rPr(), C[fill_key])
    doc.styles.add_style("强调-红", WD_STYLE_TYPE.CHARACTER).font.color.rgb = rgb("red")
    st = doc.styles.add_style("答案-蓝", WD_STYLE_TYPE.CHARACTER)
    st.font.color.rgb = rgb("blue")
    st.font.underline = True


# ---------------- 渲染 ----------------
class Renderer:
    def __init__(self, doc, mode="normal", show_sources=False, allow_gaps=False):
        self.doc, self.mode, self.show_sources, self.allow_gaps = doc, mode, show_sources, allow_gaps

    cur = None
    pending_break = False

    def para(self, el, text=None, **kw):
        self.cur = el
        p = self.doc.add_paragraph(style=sname(el))
        if self.pending_break:   # 分页做成「段前分页」：上一页正好排满时不会多出一张空白页
            p.paragraph_format.page_break_before = True
            self.pending_break = False
        if text:
            self.runs(p, text, **kw)
        return p

    def runs(self, p, text, bold=None, color=None, style=None):
        """行内标记：__下划线__ [[着重号]] {{默写词}} 〔出处〕 【缺】"""
        if not self.show_sources:
            text = SRC_RE.sub("", text)
        text = text.replace("（　　）", "\u2060（\u00a0\u00a0\u00a0\u00a0）")   # 选择括号不在中间或括号前孤立断行
        for part in re.split(r"(__.+?__|\[\[.+?\]\]|\{\{.+?\}\}|【缺】[^【〔]*|＿{2,})", text):
            if not part:
                continue
            r = None
            if part.startswith("＿"):  # 句中/句尾填空：西文下划线连成一条线，行尾不会被吞掉
                r = p.add_run("_" * max(6, len(part) * 2), style=style)
                set_fonts(r._r.get_or_add_rPr(), F["latin"], F["latin"])
            elif part.startswith("__") and part.endswith("__"):
                r = p.add_run(part[2:-2], style=style)
                r.font.underline = True
            elif part.startswith("[[") and part.endswith("]]"):
                r = p.add_run(part[2:-2], style=style)
                em = OxmlElement("w:em")
                em.set(qn("w:val"), "dot")
                r._r.get_or_add_rPr().append(em)
            elif part.startswith("{{") and part.endswith("}}"):
                word = part[2:-2]
                if self.mode == "默写":  # 西文下划线字符连成一条线；空格加下划线在行尾、表格里会被吞掉
                    n = max(6, sum(2 if ord(ch) > 0x2E80 else 1 for ch in word) + 2)
                    r = p.add_run("_" * n, style=style)
                    set_fonts(r._r.get_or_add_rPr(), F["latin"], F["latin"])
                else:
                    r = p.add_run(word, style=style)
                    if self.mode == "背诵":
                        r.font.underline = True
                    elif THEME.get("key_bold"):
                        r.bold = True
                        if self.cur in ("讲解正文", "易错条", "例句", "知识小标题") and not color:
                            r.font.color.rgb = rgb("theme")
            elif part.startswith("【缺】"):
                if not self.allow_gaps:
                    raise SystemExit(f"内容稿还有【缺】没补：{part[:40]}（补齐，或加 --allow-gaps 出预览版）")
                r = p.add_run(part, style=style)
                r.font.color.rgb = rgb("red")
                r.bold = True
                continue
            else:
                r = p.add_run(part, style=style)
            if bold is not None:
                r.bold = bold
            if color:
                r.font.color.rgb = rgb(color)

    def page_break(self):
        self.pending_break = True

    def answer_lines(self, n):
        for _ in range(n):
            self.para("作答线").add_run("\t").font.underline = True

    def options(self, text, keep_last=False):
        if not self.show_sources:
            text = SRC_RE.sub("", text)  # 出处标记不参与选项宽度估算
        opts = [o.strip() for o in re.split(r"\s{2,}|｜|\|", text) if o.strip()]
        width_chars = TEXT_W.pt / E["选项"]["size"] - 2

        def w(s):
            return sum(1 if ord(ch) > 0x2E80 else 0.55 for ch in s)

        per_line = len(opts)
        while per_line > 1 and max(w(o) for o in opts) > width_chars / per_line - 1.5:
            per_line = (2 if len(opts) in (4, 6) else 1) if per_line > 2 else 1
        for i in range(0, len(opts), per_line):
            p = self.para("选项")
            indent = Cm(E["选项"].get("indent_cm", 0))
            slot = (TEXT_W - indent) // per_line
            for j in range(1, per_line):
                p.paragraph_format.tab_stops.add_tab_stop(Emu(indent + slot * j))
            self.runs(p, "\t".join(opts[i:i + per_line]))
            has_more = i + per_line < len(opts)
            if has_more:
                p.paragraph_format.space_after = Pt(1.5)
            p.paragraph_format.keep_with_next = has_more or keep_last   # 选项不拆页；例题选项和「答：」同页

    def example(self, text):
        """例句：英文常规黑字，「｜」后的中文译文灰色小一号。"""
        p = self.para("例句")
        en, _, cn = text.partition("｜")
        r0 = p.add_run("▸ ")
        r0.font.color.rgb = rgb("theme")
        r0.font.size = Pt(E["例句"]["size"] - 2)
        self.runs(p, en.rstrip())
        if cn.strip():
            p.add_run("　")
            before = len(p.runs)
            self.runs(p, cn.strip())
            for r in p.runs[before:]:
                r.font.color.rgb = rgb("eg_cn")
                r.font.size = Pt(E["例句"]["size"] - 1)

    def table(self, rows, head=True, first=False):
        """[表]…[/表]：浅蓝表头、细灰蓝框线，列宽按内容长短分配。"""
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        tb = self.doc.add_table(rows=len(rows), cols=ncol)
        tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement("w:" + side)
            for k2, v in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", C["table_line"])):
                e.set(qn("w:" + k2), v)
            borders.append(e)
        tb._tbl.tblPr.append(borders)
        mar = OxmlElement("w:tblCellMar")
        for side, v in (("top", 28), ("bottom", 28), ("left", 80), ("right", 80)):
            e = OxmlElement("w:" + side)
            e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa")
            mar.append(e)
        tb._tbl.tblPr.append(mar)

        def w(s):
            s = re.sub(r"\{\{|\}\}|〔[^〕]*〕", "", s)
            return sum(1 if ord(ch) > 0x2E80 or ch in "…—“”‘’" else 0.55 for ch in s)  # 省略号、破折号在中文字体里占全宽
        center = [max(w(r[j]) for r in rows[1 if head else 0:]) <= 9 for j in range(ncol)]
        # 列宽：短列（不超过 8 字）按内容给足、不折行；剩下的宽度按内容长短分给长列
        avail = TEXT_W.pt / E["表格文字"]["size"]      # 整行能放下的字数
        nat = [max(1.0, max(w(r[j]) for r in rows)) + 1.6 for j in range(ncol)]   # +1.6：单元格左右内边距和粗体余量
        short = [n <= 9.5 for n in nat]
        fixed = sum(n for n, sh in zip(nat, short) if sh)
        longs = [n for n, sh in zip(nat, short) if not sh]
        if not longs or sum(nat) <= avail:
            grow = (avail - sum(nat)) / ncol if sum(nat) <= avail else 0
            cols = [n + grow for n in nat] if sum(nat) <= avail else [n * avail / sum(nat) for n in nat]
        else:
            rest = max(avail - fixed, avail * 0.35)
            lt = sum(min(n, 70) for n in longs)
            cols = [n if sh else rest * min(n, 70) / lt for n, sh in zip(nat, short)]
        tot = sum(cols)
        widths = [Emu(int(TEXT_W * c / tot)) for c in cols]
        tb.autofit = False
        for col, wd in zip(tb.columns, widths):
            col.width = wd
        for i, row in enumerate(rows):
            for j, txt in enumerate(row):
                c = tb.rows[i].cells[j]
                c.width = widths[j]
                is_head = head and i == 0
                para = c.paragraphs[0]
                para.style = sname("表头文字" if is_head else "表格文字")
                if is_head:
                    add_shd(c._tc.get_or_add_tcPr(), C["table_head"])
                self.cur = "表格文字"
                self.runs(para, txt)
                if center[j] or is_head:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tcpr = c._tc.get_or_add_tcPr()
                va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "center"); tcpr.append(va)
                if first and j == 0 and not is_head:
                    for r in para.runs:
                        r.bold = True
        keep_whole = len(tb.rows) <= 6
        if not keep_whole and head:  # 长表允许跨页接续，避免半页空白，跨页时自动重复表头
            tb.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        for i, row in enumerate(tb.rows):  # 单行不断开；小表（≤6行）整表不跨页，长表保表头+前两行、末两行同页
            trpr = row._tr.get_or_add_trPr()
            trpr.append(OxmlElement("w:cantSplit"))
            if (keep_whole and i < len(tb.rows) - 1) or (not keep_whole and (i < 2 or i == len(tb.rows) - 2)):
                for c in row.cells:
                    for para in c.paragraphs:
                        para.paragraph_format.keep_with_next = True
        gap = self.doc.add_paragraph(style=sname("表格文字")).paragraph_format  # 表后留极小空，紧凑衔接
        gap.line_spacing, gap.space_after = Pt(3), Pt(0)

    def template_blank(self, text):
        """默写版：答题模板整句换成空白横线，条数按原句长度估算。"""
        clean = re.sub(r"\{\{|\}\}|__|\[\[|\]\]", "", SRC_RE.sub("", text))
        per_line = TEXT_W.pt / E["模板"]["size"]
        self.answer_lines(max(1, math.ceil(len(clean) / per_line)))

    def blocks(self, blocks):
        first_unit = True
        merged = set()
        for i, b in enumerate(blocks):
            if i in merged:
                continue
            k, t = b["kind"], b["text"]
            nxt = blocks[i + 1]["kind"] if i + 1 < len(blocks) else None
            lab = LABELS.get(b["label"], DEFAULT_LABEL) if b["label"] else None
            if k == "pagebreak":
                self.page_break()
            elif k in ("unit", "answer_start"):
                if not first_unit or self.mode != "首单元不分页":
                    self.page_break()
                first_unit = False
                self.para("单元条", t)
            elif k == "yaosu":
                self.para("单元要素", t)
            elif k == "kd":
                self.para("考点", t)
            elif k == "label":
                key = t if t in LABELS else "其他"
                p = self.para("栏目")
                self.runs(p, f"{t}：", style="标签-" + key, bold=True)
            elif k == "sub":
                self.para("知识小标题", t)
            elif k == "eg":
                self.example(t)
            elif k == "table":
                self.table(b["rows"], b["head"], b["first"])
            elif k in ("body", "pair"):
                el = lab[1] if lab else "讲解正文"
                if THEME and b["label"] in ("易错点", "考点", "考法"):
                    el = "易错条"
                if THEME and b["label"] == "口诀":
                    el = "口诀框"
                if self.mode == "默写" and b["label"] == "答题模板":
                    self.template_blank(t)
                    continue
                p = self.para(el)
                if SRC_RE.sub("", t).rstrip().endswith(("：", ":")) or nxt == "eg":   # 引出下文的句子、带例句的易错条不留在页尾
                    p.paragraph_format.keep_with_next = True
                if k == "pair" and re.search(r"[:：]", t):
                    head, sep, tail = re.split(r"([:：])", t, maxsplit=1)
                    self.runs(p, head + sep, style="强调-红", bold=True if THEME else None)
                    self.runs(p, tail)
                else:
                    self.runs(p, t)
            elif k == "li":
                self.para("例题", t)
            elif k == "passage":
                self.para("课本句" if THEME and b["label"] == "课本原句" else "例文", t)
            elif k == "source":
                p = self.para("来源", t)
                if nxt in ("q", "opts", "lines", "ex_sec", "da", "label", "sub"):
                    p.paragraph_format.keep_with_next = True
            elif k == "da":
                p = self.para("示范答案")
                body = re.sub(r"^答[:：]", "", t)
                self.runs(p, "答：")
                key, gap, why = re.split(r"(　+|\s{2,})", body, maxsplit=1) if THEME and re.search(r"　|\s{2,}", body) else (body, "", "")
                self.runs(p, key, style="答案-蓝", bold=True)
                if why:
                    self.runs(p, "　" + why)
            elif k == "ex_title":
                self.para("练习标题", t)
            elif k == "ex_para":
                prev_k = blocks[i - 1]["kind"] if i > 0 else None
                p = self.para("练习正文", t)
                if prev_k != "ex_para":
                    p.paragraph_format.space_before = Pt(3)
                if nxt != "ex_para":
                    p.paragraph_format.space_after = Pt(6)
            elif k == "ex_sec":
                self.para("练习大题", t)
            elif k == "ex_sub":
                self.para("练习小标题", t)
            elif k == "q" and re.match(r"^\d+[.．]\s*A[.．]", SRC_RE.sub("", t)) and "｜" in t:
                self.options(t)
            elif k == "q":
                p = self.para("练习小题", t)
                cont = nxt == "q" and not re.match(r"^\d+[.．]", blocks[i + 1]["text"])  # 题干的第二、三行（对话）
                if nxt in ("opts", "lines") or cont:
                    p.paragraph_format.space_after = Pt(1.5)
                p.paragraph_format.keep_with_next = nxt in ("opts", "lines") or cont
            elif k == "opts":
                self.options(t, keep_last=nxt == "da")
            elif k == "lines":
                self.answer_lines(b["n"])
            elif k == "ans_head":   # 本章答案单独另起一页
                self.page_break()
                self.para("练习标题", t)
            elif k == "ans_unit":
                self.para("答案单元", t)
            elif k == "ans_title":
                self.para("答案篇名", t)
            elif k == "ans_q":
                p = self.para("答案题")
                m = re.match(r"^(.*?考点[:：])(.*)$", t)
                self.runs(p, m.group(1), style="强调-红", bold=True)
                self.runs(p, m.group(2))
            elif k == "ans_ans" and COMPACT_ANS and nxt == "ans_why":   # 紧凑排法：答案和解析排在同一段（内容稿文件头「答案排法: 紧凑」）
                p = self.para("答案题")
                head, sep, tail = re.split(r"([:：])", t, maxsplit=1)
                self.runs(p, re.sub(r"参考答案$", "", head) if re.match(r"^\d", head) else head + sep, bold=True)
                self.runs(p, tail, style="答案-蓝", bold=True)
                why = blocks[i + 1]["text"]
                wh, ws, wt = re.split(r"([:：])", why, maxsplit=1)
                self.runs(p, "　" + wh + ws, bold=True)
                self.runs(p, wt)
                p.paragraph_format.keep_with_next = False   # 「答案题」样式默认和下一段同页；合并后不再串成一整块
                p.paragraph_format.keep_together = True     # 单条答案与解析不跨页断开
                merged.add(i + 1)
            elif k in ("ans_why", "ans_ans"):
                p = self.para("答案题" if re.match(r"^(\d+[.．]|（\d+）)", t) else "答案行")
                head, sep, tail = re.split(r"([:：])", t, maxsplit=1) if re.search(r"[:：]", t) else (t, "", "")
                self.runs(p, head + sep, bold=True)
                if k == "ans_ans":
                    self.runs(p, tail, style="答案-蓝", bold=True)
                    if nxt == "ans_why":
                        p.paragraph_format.keep_with_next = True
                else:
                    self.runs(p, tail)
                    p.paragraph_format.keep_together = True
            elif k == "ans_text":
                self.para("答案行", t)
            elif k == "exam_head":
                self.para("卷头标题", t)
                self.para("卷头信息", "学校：＿＿＿＿＿＿　班级：＿＿＿＿＿＿　姓名：＿＿＿＿＿＿　得分：＿＿＿＿＿＿")


def field(p, instr):
    for kind in ("begin", "instr", "separate", "text", "end"):
        r = p.add_run()
        if kind == "instr":
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = f" {instr} "
            r._r.append(it)
        elif kind == "text":
            r.text = "1"
        else:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            r._r.append(fc)


def setup_page(doc):
    s = doc.sections[0]
    s.page_width, s.page_height = PAGE_W, Cm(29.7)
    s.top_margin, s.bottom_margin = Cm(PG["margin_top_cm"]), Cm(PG["margin_bottom_cm"])
    s.left_margin, s.right_margin = Cm(PG["margin_left_cm"]), Cm(PG["margin_right_cm"])
    s.footer_distance = Cm(PG["footer_distance_cm"])
    p = s.footer.paragraphs[0]
    p.style = doc.styles["Footer"]
    for part in re.split(r"(\{PAGE\}|\{NUMPAGES\})", PG["footer_text"]):
        if part in ("{PAGE}", "{NUMPAGES}"):
            field(p, part[1:-1])
        elif part:
            p.add_run(part)


def toc_entries(blocks, with_passages=True):
    out = []
    for b in blocks:
        t = SRC_RE.sub("", b["text"]).strip()
        if b["kind"] in ("unit", "answer_start"):
            out.append((1, t))
        elif b["kind"] == "kd":
            out.append((2, t))
        elif b["kind"] in ("ex_title", "ans_head") and with_passages:
            out.append((3, t))
    return out


def title_block(doc, meta, variant_label, entries, pages):
    p = doc.add_paragraph(style=sname("总标题"))
    for seg in meta.get("标题", meta.get("简称", "资料包")).split():
        text, _, color = seg.partition("|")
        r = p.add_run(text)
        r.font.color.rgb = rgb(color or "black")
    sub = meta.get("副标题", "")
    if variant_label:
        sub = f"{sub}（{variant_label}）" if sub else variant_label
    if sub:
        doc.add_paragraph(sub, style=sname("副标题"))
    for lvl, text in entries:
        doc.add_paragraph(f"{text}\t{pages.get(text, '')}", style=sname(f"目录{lvl}"))


def checklist(doc, meta, blocks):
    """背诵默写版第 1 页：考点打卡表（完成 √× + 家长签字）。"""
    p = doc.add_paragraph(style=sname("总标题"))
    p.add_run(meta.get("简称", "") + "背诵与默写")
    kds = [SRC_RE.sub("", b["text"]).strip() for b in blocks if b["kind"] == "kd"]
    table = doc.add_table(rows=len(kds) + 1, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Emu(int(TEXT_W * r)) for r in (0.66, 0.17, 0.17)]
    for j, head in enumerate(("考点背诵内容", "完成 √ ×", "家长签字")):
        c = table.rows[0].cells[j]
        c.text = ""
        c.paragraphs[0].style = sname("打卡表头")
        c.paragraphs[0].add_run(head)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, kd in enumerate(kds, 1):
        cells = table.rows[i].cells
        cells[0].text = kd
        cells[1].text = "□　□"
        cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False
    for col, w in zip(table.columns, widths):  # LibreOffice 看列网格宽度，Word 看单元格宽度，两个都设
        col.width = w
    for row in table.rows:
        for j, c in enumerate(row.cells):
            c.width = widths[j]


def build(content, out, variant="完整版", pages=None, show_sources=False, allow_gaps=False):
    global COMPACT_ANS
    meta, blocks = parse(content)
    COMPACT_ANS = meta.get("答案排法") == "紧凑"
    apply_theme(meta.get("类型"))
    doc = Document()
    build_styles(doc)
    setup_page(doc)
    pages = pages or {}
    sel = select(blocks, variant)
    label = "" if variant == "完整版" else variant
    if variant == "背诵默写":
        checklist(doc, meta, sel)
        r = Renderer(doc, "背诵", show_sources, allow_gaps)
        r.blocks(sel)
        r2 = Renderer(doc, "默写", show_sources, allow_gaps)
        r2.blocks(sel)
    else:
        entries = toc_entries(sel, meta.get("目录篇目", "是") != "否")
        title_block(doc, meta, label, entries, pages)
        Renderer(doc, "normal", show_sources, allow_gaps).blocks(sel)
    normalize(doc)
    cp = doc.core_properties   # 文件属性不留任何人名和工具名（python-docx 默认作者是 python-docx）
    cp.author = cp.last_modified_by = cp.comments = cp.keywords = cp.subject = cp.category = ""
    cp.title = re.sub(r"\s*\|[A-Za-z]+\s*", "", meta.get("标题", "")).strip()   # 去掉标题里的「|blue」配色标记
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return meta, sel


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("out")
    ap.add_argument("--variant", default="完整版", choices=VARIANTS)
    ap.add_argument("--pages")
    ap.add_argument("--show-sources", action="store_true")
    ap.add_argument("--allow-gaps", action="store_true")
    a = ap.parse_args()
    pages = json.loads(Path(a.pages).read_text(encoding="utf-8")) if a.pages else {}
    build(a.content, a.out, a.variant, pages, a.show_sources, a.allow_gaps)
    print("->", a.out)
