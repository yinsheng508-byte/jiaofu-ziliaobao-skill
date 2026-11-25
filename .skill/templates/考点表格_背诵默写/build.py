"""考点笔记（考点表格风，背诵版 / 默写版）：基于 知识手册_背诵默写 模板改，新增 @考点 / [点拨] / 红框表格 / 双线页眉。

知识手册（背诵版 / 默写版）：内容稿.md → 两个 Word（再用 Word 导出 PDF、回填目录页码、拼封面）

  python3 排版/build.py <项目> [--only 背诵|默写] [--pdf]

学科、系列之间的差别（主色、框的种类、标题样式、目录提示语……）全在 排版/样式.json，build.py 不用改。
内容稿标记见 项目规范.md（模板说明.md 里有完整列表）。
"""
import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def _skill_scripts():
    import os
    if os.environ.get("JIAOFU_SKILL"):                  # run.py build 会设置
        return Path(os.environ["JIAOFU_SKILL"]) / "scripts"
    for p in [HERE.parent.parent / "scripts", *[q / "scripts" for q in HERE.parents if (q / "SKILL.md").exists()]]:
        if (p / "docx_kit.py").exists():
            return p
    # 项目里的副本：沿着 输出项目/ 往上找 .skill
    for q in HERE.parents:
        if (q / ".skill" / "scripts" / "docx_kit.py").exists():
            return q / ".skill" / "scripts"
    raise SystemExit("找不到 skill 的 scripts/docx_kit.py：用 run.py build 跑，或设置环境变量 JIAOFU_SKILL=<skill 目录>")


sys.path.insert(0, str(_skill_scripts()))
import docx_kit as K  # noqa: E402
from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor, Emu  # noqa: E402
from PIL import Image  # noqa: E402

S = K.load_theme(HERE / "样式.json")          # 样式.json 里写 "主题": "<名>" 就先套 assets/主题/ 里的配色，再用本文件覆盖
COL = S["colors"]
SERIF, HEI, KAI, LATIN = K.FONTS["serif"], K.FONTS["hei"], K.FONTS["kai"], K.FONTS["latin"]
TITLE_FONT = K.FONTS["title"]
FONT = {"serif": SERIF, "hei": HEI, "kai": KAI, "title": TITLE_FONT}
BODY = S["正文"]["字号"]
LINE = S["正文"]["行距"]
MARGINS = S["页面"]["边距"]               # 上 下 左 右
PAGE_W = 21.0
TEXT_W = PAGE_W - 2 * MARGINS[2] if MARGINS[2] == MARGINS[3] else PAGE_W - MARGINS[2] - MARGINS[3]
IMG_DIRS = [HERE / d for d in S.get("插图目录", ["插图", "课本插图_全部"])]
STRIP = re.compile(S.get("去掉的出处括注", r"（(?:S0\d|课本[图表])[^）]*）"))
INLINE = re.compile(r"(\{\{.+?\}\}|\*\*.+?\*\*|\^\^.+?\^\^)")
D = S["细节"]


def rgb(k):
    return RGBColor.from_string(COL.get(k, k))


def set_font(run, east=SERIF, latin=LATIN, size=None, bold=None, color=None):
    K.set_font(run, east, latin, size, bold, COL.get(color, color) if color else None)


def para_fmt(p, before=0, after=0, line=LINE, **kw):
    K.para_fmt(p, before, after, line, **kw)


shd, p_border, borders, cell_border = K.shd, K.p_border, K.borders, K.cell_border
cell_margins, fixed_widths, tbl_indent, no_split = K.cell_margins, K.fixed_widths, K.tbl_indent, K.no_split
keep_rows, keep_rows_head, block_end = K.keep_rows, K.keep_rows_head, K.block_end


def find_img(name):
    name = name.strip()
    for d in IMG_DIRS:
        for p in sorted(d.glob("*.png")):
            if p.stem.split("_", 2)[-1] == name:
                return p
    raise SystemExit(f"找不到插图：{name}（放在 排版/插图/，文件名形如 M_p010_{name}.png）")


class Builder:
    def __init__(self, mode):
        self.mode = mode
        self.doc = Document()
        K.setup_page(self.doc, MARGINS, S["页面"].get("页眉距", 0.75), S["页面"].get("页脚距", 0.7), BODY)
        self.pages = {}
        self.pending_break = False
        self.blanks = 0
        self.noblank = False
        self.chapter_no = 0
        self.fresh = {}      # 单元格里还没用过的第一个空段落

    # ---------- 基础 ----------
    def runs(self, p, text, size=BODY, color="black", bold=False, east=SERIF):
        for part in INLINE.split(text):
            if not part:
                continue
            if part.startswith("{{") and self.noblank:
                r = p.add_run(part[2:-2]); set_font(r, east, LATIN, size, True, "red")
            elif part.startswith("{{"):
                w = part[2:-2]
                self.blanks += 1
                r = p.add_run(" " + w.replace(" ", " ") + " ")
                set_font(r, east, LATIN, size, True if self.mode == "背诵" else bold)
                r.font.underline = True
                if self.mode == "背诵":
                    r.font.color.rgb = rgb("red")
                    r.font.underline = False               # 考点表格风：浅粉底红字，不要下划线
                    shd(r._r, COL["hl"])
                else:   # 默写版：字涂白、留下划线，挖空和背诵版同宽，两版页码一致
                    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    r._r.get_or_add_rPr().find(qn("w:u")).set(qn("w:color"), "333333")
            elif part.startswith("^^"):          # 蓝色补充说明
                r = p.add_run(part[2:-2]); set_font(r, KAI, LATIN, size - 0.3, True, "blue")
            elif part.startswith("**"):
                r = p.add_run(part[2:-2]); set_font(r, east, LATIN, size, True, color)
            else:
                r = p.add_run(part); set_font(r, east, LATIN, size, bold, color)

    def para(self, cont=None, **fmt):
        cont = self.doc if cont is None else cont
        key = id(cont)
        if key in self.fresh:
            p = self.fresh.pop(key)
        else:
            p = cont.add_paragraph()
        para_fmt(p, **fmt)
        if self.pending_break and cont is self.doc:
            p.paragraph_format.page_break_before = True
            self.pending_break = False
        return p

    def new_table(self, cont, rows, cols):
        cont = self.doc if cont is None else cont
        if cont is self.doc and self.pending_break:
            self.para(line=2)
        if id(cont) in self.fresh:          # 单元格开头先放表：把预留空段落删掉
            p = self.fresh.pop(id(cont)); p._p.getparent().remove(p._p)
        t = cont.add_table(rows=rows, cols=cols)
        if cont is not self.doc:            # python-docx 在单元格里加表后会补一个空段落，压到极小
            last = cont.paragraphs[-1]
            para_fmt(last, line=3)
            self.fresh[id(cont)] = last
        return t

    def open_cell(self, cell):
        self.fresh[id(cell)] = cell.paragraphs[0]
        return cell

    def close_cell(self, cell):
        p = self.fresh.pop(id(cell), None)
        if p is not None and len(cell._tc.findall(qn("w:p"))) + len(cell._tc.findall(qn("w:tbl"))) > 1:
            p._p.getparent().remove(p._p)

    def spacer(self, cont, h=6):
        return self.para(cont, line=h)

    # ---------- 标题 ----------
    def chapter(self, text):
        """# 第一单元　少年有梦：居中大标题，另起页（参考：「第一章 职业理念」）"""
        tag, _, name = text.partition("　")
        c = D["章"]
        self.chapter_no += 1
        self.pending_break = True
        p = self.para(before=4, after=8, line=None, align=WD_ALIGN_PARAGRAPH.CENTER, keep_next=True)
        r = p.add_run(tag + "　" + name); set_font(r, TITLE_FONT, LATIN, c.get("字号", 22), False, "deep")
        self.just_chapter = True

    def section(self, text):
        """## 第一课　开启初中生活：左对齐粗体大节标题（参考：「第一节 教育观」），课接排不另起页"""
        tag, _, name = text.partition("　")
        c = D["节"]
        if c.get("另起页", False) and not getattr(self, "just_chapter", False):
            self.pending_break = True
        self.just_chapter = False
        self.after_section = True
        p = self.para(before=10, after=4, line=None, keep_next=True)
        r = p.add_run(tag + c["分隔"] + name); set_font(r, FONT[c["字体"]], LATIN, 14, True, "deep")

    def kp(self, text, cont=None):
        """### 第一节　奏响中学序曲：小一级粗体黑体标题，左侧主色短竖条"""
        p = self.para(cont, before=8, after=3, line=None, keep_next=True, left=0.25)
        tag, _, name = text.partition("　")
        r = p.add_run(tag + "　" + name if name else tag); set_font(r, HEI, LATIN, 11.5, True, "main")
        self.kd_no = 0

    kd_no = 0

    def kaodian(self, spec):
        """@考点 标题 | 标签1 ; 标签2：「考点 N」浅粉底红字小标签 + 考点名 + 右侧红色【】标签"""
        title, _, tags = [x.strip() for x in (spec.partition("|"))]
        self.kd_no += 1
        p = self.para(before=7, after=3, line=None, keep_next=True)
        r = p.add_run(f" 考点{self.kd_no} "); set_font(r, HEI, LATIN, 9.5, True, "red")
        shd(r._r, COL["hl"])
        bd = OxmlElement("w:bdr")
        for k, v in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", COL["rule"])):
            bd.set(qn("w:" + k), v)
        r._r.get_or_add_rPr().append(bd)
        r = p.add_run(" " + title); set_font(r, HEI, LATIN, 10.5, True, "deep")
        tg = [x.strip() for x in tags.split(";") if x.strip()]
        if tg:
            r = p.add_run("　"); set_font(r, SERIF, LATIN, 10.5)
            for x in tg:
                x = "★必背" if x == "必背" else x
                r = p.add_run("【" + x + "】"); set_font(r, HEI, LATIN, 8.5, True, "red")

    # ---------- 条目 ----------
    def item(self, line, cont=None, size=BODY, east=SERIF, line_h=LINE, nested=False):
        m = re.match(r"^(\d+\.\s?|\(\d+\)|（\d+）|[①-⑳])\s*(.*)$", line)
        lvl, mark, rest = 0, "", line
        if m:
            mark, rest = m.group(1), m.group(2)
            lvl = 1 if mark[0].isdigit() else 2 if mark[0] in "(（" else 3
        if lvl == 1:
            mark = mark.replace(" ", "").replace(".", ". ")
            left, first = 0.5, -0.5
        elif lvl == 2:
            mark = mark.replace("(", "（").replace(")", "）")
            left, first = 0.95, -0.95
        elif lvl == 3:                       # 有（1）（2）时，①②③ 缩进一层挂在下面
            left, first = (1.5, -0.5) if nested else (0.55, -0.55)
        else:
            left, first = 0, 0
        p = self.para(cont, left=left, first=first, line=line_h, after=2.5 if lvl == 1 else 1, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        lm = re.match(r"^([^：:{}，。（）、]{1,14}[：:])(.*)$", rest)
        if rest.rstrip().endswith("：") and not cont is None:
            p.paragraph_format.keep_with_next = True          # 「学会学习的要求：」这类引导行不落在页尾
        if mark:
            r = p.add_run(mark); set_font(r, SERIF, LATIN, size, lvl == 1, "main" if lvl == 1 else "black")
        if lm and lvl in (0, 1, 2):
            r = p.add_run(lm.group(1)); set_font(r, SERIF, LATIN, size, True, "main")
            self.runs(p, lm.group(2), size=size, east=east)
        else:
            self.runs(p, rest, size=size, east=east)
        return p

    def cell_text(self, p, txt, size):
        """单元格文字：<br> 换行；「原因：」「做法：」这类开头的小标签用蓝色加粗"""
        for k, seg in enumerate(txt.split("<br>")):
            if k:
                p.add_run().add_break()
            m = re.match(r"^(原因|做法|表现|影响)：(.*)$", seg)
            if m:
                r = p.add_run(m.group(1) + "："); set_font(r, KAI, LATIN, size, True, "blue")
                seg = m.group(2)
            self.runs(p, seg, size=size)

    # ---------- 表 ----------
    def table(self, rows, first_bold=False, cont=None, width=TEXT_W):
        c_ = D["表"]
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        t = self.new_table(cont, len(rows), ncol)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        dark = COL["main"]
        line = (c_["线粗"], COL[c_["线色"]])                  # 考点表格风：四周和内部都是细红框线
        borders(t, {"top": line, "bottom": line, "left": line, "right": line, "insideH": line, "insideV": line})
        cell_margins(t, 50, 50, 100, 100); tbl_indent(t)

        ONE = re.compile(r"\[\[图\s+(.+?)\s+高=([\d.]+)\]\]")
        IMGC = re.compile(r"^(?:\s*\[\[图\s+.+?\s+高=[\d.]+\]\]\s*)+$")

        def imgw(x):
            if not IMGC.match(x.strip()):
                return 0
            tot = 0
            for nm, h in ONE.findall(x):
                iw, ih = Image.open(find_img(nm)).size
                tot += float(h) * iw / ih + 0.15
            return tot + 0.2

        def clen(x):
            return 0 if IMGC.match(x.strip()) else len(re.sub(r"[{}*]", "", x))
        maxl = [max(clen(r[i]) for r in rows[1:]) if len(rows) > 1 else 0 for i in range(ncol)]
        maxl = [max(m, clen(rows[0][i])) for i, m in enumerate(maxl)]
        iws = [max(imgw(r[i]) for r in rows) for i in range(ncol)]
        cw = 0.36
        nat = [iws[i] + 0.2 if iws[i] else min(max(maxl[i] + 1.5, clen(rows[0][i]) + 1.5, 3), 10) * cw + 0.35 for i in range(ncol)]
        extra = width - sum(nat)
        wts = [0 if iws[i] else max(0, maxl[i] - 9) for i in range(ncol)]
        if extra > 0 and sum(wts):
            nat = [n + extra * w_ / sum(wts) for n, w_ in zip(nat, wts)]
        else:
            fixed = sum(n for n, w_ in zip(nat, iws) if w_)
            rest = sum(n for n, w_ in zip(nat, iws) if not w_)
            nat = [n if iws[k] else n * (width - fixed) / rest for k, n in enumerate(nat)]
        fixed_widths(t, nat)
        for i, r in enumerate(rows):
            no_split(t.rows[i])
            for j, txt in enumerate(r):
                c = t.rows[i].cells[j]
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = c.paragraphs[0]
                centre = i == 0 or maxl[j] <= 12                                   # 字少的列居中，长句列靠左（参考图：解读列靠左）
                para_fmt(p, line=16, align=WD_ALIGN_PARAGRAPH.CENTER if centre else WD_ALIGN_PARAGRAPH.LEFT)
                im = IMGC.match(txt.strip())
                if im and i > 0:
                    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for k_, (nm, h) in enumerate(ONE.findall(txt)):
                        if k_:
                            p.add_run("  ")
                        p.add_run().add_picture(str(find_img(nm)), height=Cm(float(h)))
                elif i == 0:
                    r_ = p.add_run(txt); set_font(r_, HEI, LATIN, 10, True, c_["表头字色"])
                    if c_.get("表头底纹"):
                        shd(c._tc, COL["head"])
                elif first_bold and j == 0:
                    shd(c._tc, COL["head"])
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    self.runs(p, txt, size=10, bold=True, color="deep", east=HEI)
                else:
                    self.cell_text(p, txt, 10)
        # 第一列空着的行并进上一行（如「第一课」下面两节）
        r0 = 1
        while r0 < len(rows):
            if rows[r0][0] == "":
                r1 = r0
                while r1 + 1 < len(rows) and rows[r1 + 1][0] == "":
                    r1 += 1
                top = t.cell(r0 - 1, 0)
                m = top.merge(t.cell(r1, 0))
                for q in m.paragraphs[1:]:
                    q._p.getparent().remove(q._p)
                r0 = r1 + 1
            else:
                r0 += 1
        if len(rows) <= 6:
            keep_rows(t)
        else:
            keep_rows_head(t)
        K.repeat_header(t)
        self.spacer(cont, 10)

    # ---------- 图 ----------
    def picture(self, cell, name, width=None, height=None, caption=True, size=8):
        path = find_img(name)
        p = self.para(cell, before=2, line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run()
        if height:
            r.add_picture(str(path), height=Cm(height))
        else:
            r.add_picture(str(path), width=Cm(width))
        p.paragraph_format.keep_together = True
        if caption:
            r.add_break()
            rr = p.add_run(path.stem.split("_", 2)[-1]); set_font(rr, KAI, LATIN, size, False, "grey")

    def fit(self, name, maxw, maxh):
        iw, ih = Image.open(find_img(name)).size
        return min(maxh, maxw * ih / iw)

    def figure_row(self, spec, cont=None, width=TEXT_W):
        files = spec["files"]; h = float(spec.get("高", 3.2)); n = len(files)
        ratios = []
        for f in files:
            iw, ih = Image.open(find_img(f)).size
            ratios.append(iw / ih)
        auto = (width * (0.72 if n == 1 else 0.96) - 0.5 * n) / sum(ratios)   # 图尽量放大，高度封顶
        if not spec.get("定"):
            h = max(h, min(auto, 7.8 if n == 1 else 5.6))
        nat = [min(r * h, width) + 0.4 for r in ratios]
        if sum(nat) > width:
            nat = [x * width / sum(nat) for x in nat]
        else:
            nat = [x + (width - sum(nat)) / n for x in nat]
        t = self.new_table(cont, 1, n)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders(t, {}); cell_margins(t, 0, 0, 30, 30); tbl_indent(t)
        fixed_widths(t, nat); no_split(t.rows[0])
        for i, f in enumerate(files):
            c = self.open_cell(t.rows[0].cells[i])
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
            hh = min(h, (nat[i] - 0.3) / ratios[i])
            self.picture(c, f, height=hh)
            self.close_cell(c)
        self.spacer(cont, 10)

    def figure_right(self, spec, body, cont=None, width=TEXT_W):
        w = float(spec.get("宽", 4.2)); files = spec["files"]
        t = self.new_table(cont, 1, 2)
        borders(t, {}); cell_margins(t, 0, 0, 0, 0); tbl_indent(t)
        lw = width - w - 0.3
        fixed_widths(t, [lw, w + 0.3])
        left, right = (self.open_cell(c) for c in t.rows[0].cells)
        self.render(body, left, lw - 0.1)
        self.close_cell(left)
        right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        txt = sum(len(re.sub(r"[{}*\[\]]", "", l)) for l in body if l.strip() and not l.strip().startswith("["))
        for f in files:
            hh = self.fit(f, w - 0.1, 7.0 / len(files))
            self.picture(right, f, height=hh)
        self.close_cell(right)
        if txt < 200:
            no_split(t.rows[0])
        self.spacer(cont, 10)

    # ---------- 答案框、点拨框 ----------
    def answer_box(self, lines, cont=None, width=TEXT_W):
        """没有表的考点：一个细红框，里面 ①②③ 分条（参考：「考点 2」那种框）"""
        t = self.new_table(cont, 1, 1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        line = (D["表"]["线粗"], COL[D["表"]["线色"]])
        borders(t, {"top": line, "bottom": line, "left": line, "right": line})
        cell_margins(t, 70, 70, 130, 130); tbl_indent(t)
        fixed_widths(t, [width])
        c = self.open_cell(t.rows[0].cells[0])
        if sum(len(l) for l in lines) < 220:
            no_split(t.rows[0])
        nested = any(re.match(r"^[（(]\d+[）)]", l) for l in lines)
        for l in lines:
            self.item(l, c, size=10.2, line_h=17.5, nested=nested)
        self.close_cell(c)
        self.spacer(cont, 8)

    def callout(self, kind, body, cont=None, width=TEXT_W):
        """[点拨 注意]：浅蓝底、蓝色虚线框，蓝色楷体小字"""
        t = self.new_table(cont, 1, 1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        tcPr = t.rows[0].cells[0]._tc.get_or_add_tcPr()
        tb = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            x = OxmlElement("w:" + side); x.set(qn("w:val"), "dashed"); x.set(qn("w:sz"), "4"); x.set(qn("w:color"), COL["blue"]); tb.append(x)
        tcPr.append(tb)
        borders(t, {})
        cell_margins(t, 60, 60, 130, 130); tbl_indent(t)
        fixed_widths(t, [width])
        c = self.open_cell(t.rows[0].cells[0])
        shd(c._tc, COL["box"])
        no_split(t.rows[0])
        p = self.para(c, line=16, after=0)
        r = p.add_run(kind + "："); set_font(r, HEI, LATIN, 9.5, True, "blue")
        txt = "".join(l.strip() for l in body)
        for part in INLINE.split(txt):
            if part.startswith("{{"):
                part = part[2:-2]
            r = p.add_run(part); set_font(r, KAI, LATIN, 9.8, False, "blue")
        self.close_cell(c)
        self.spacer(cont, 8)

    # ---------- 框 ----------
    def box(self, kind, title, body, cont=None, width=TEXT_W):
        b = D["框"]
        self.spacer(cont, 10).paragraph_format.keep_with_next = True
        t = self.new_table(cont, 1, 1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders(t, {}); cell_margins(t, 130, 140, 220, 220); tbl_indent(t)
        fixed_widths(t, [width])
        c = self.open_cell(t.rows[0].cells[0])
        shd(c._tc, COL["box"])
        if sum(len(l) for l in body) < 200:
            no_split(t.rows[0])
        p = self.para(c, line=None, after=3, keep_next=True)
        r = p.add_run(kind); set_font(r, HEI, LATIN, 10, True, "warn" if kind in b["警示"] else "main")
        if title:
            r = p.add_run(b["标题分隔"] + title); set_font(r, FONT[b["标题字体"]], LATIN, b["标题字号"], b["标题加粗"], "deep")
        self.noblank = True
        self.in_box = (KAI if kind in b["楷体"] else SERIF)
        self.render(body, c, width - 0.7, size=9.8, line_h=16.5)
        self.in_box = None
        self.noblank = False
        self.close_cell(c)
        self.spacer(cont, 12)

    in_box = None

    # ---------- 渲染 ----------
    def render(self, lines, cont=None, width=TEXT_W, size=BODY, line_h=LINE):
        i = 0
        grp = []

        def flush():
            if grp:
                self.answer_box(list(grp), cont, width)
                grp.clear()
        while i < len(lines):
            s = lines[i].strip(); i += 1
            if not s:
                continue
            if not s.startswith(("[", "@", "#")):
                grp.append(s)               # 普通条目先攒着，遇到表、点拨、下一个考点时合成一个框
                continue
            flush()
            if s.startswith("### "):
                self.kp(s[4:], cont)
            elif s.startswith("@考点"):
                self.kaodian(s[3:].lstrip("：: "))
            elif s.startswith("[点拨"):
                j = block_end(lines, i, "点拨")
                self.callout(s.strip("[]").split(None, 1)[1] if " " in s else "点拨", lines[i:j], cont, width)
                i = j + 1
            elif s.startswith("[表"):
                j = block_end(lines, i, "表")
                rows = [[c.strip() for c in re.split(r"\s*\|\s*", l.strip())] for l in lines[i:j] if l.strip()]
                i = j + 1
                self.table(rows, "首列" in s, cont, width)
            elif s.startswith("[图右"):
                j = block_end(lines, i, "图右")
                self.figure_right(self.spec(s), lines[i:j], cont, width)
                i = j + 1
            elif s.startswith("[图组"):
                self.figure_row(self.spec(s), cont, width)
            elif s.startswith("[框"):
                j = block_end(lines, i, "框")
                m = re.match(r"\[框\s+(\S+)\s*(.*)\]", s)
                self.box(m.group(1), m.group(2).strip(), lines[i:j], cont, width)
                i = j + 1
            else:
                self.item(s, cont, size=size, east=self.in_box or SERIF, line_h=line_h)
        flush()

    @staticmethod
    def spec(st):
        inner = st.strip("[]").split(None, 1)[1] if " " in st else ""
        d = {}
        for m in re.finditer(r"(宽|高)=([\d.]+)(!?)", inner):
            d[m.group(1)] = m.group(2)
            if m.group(3): d["定"] = True
        inner = re.sub(r"(宽|高)=[\d.]+!?", "", inner)
        d["files"] = [f.strip() for f in inner.split(";") if f.strip()]
        return d

    def build(self, src):
        text = Path(src).read_text(encoding="utf-8")
        text = re.sub(r"<!--.*?-->", "", text, flags=re.S)
        text = STRIP.sub("", text)
        text = re.sub(r"〔[^〕]*〕", "", text)          # 〔S03 p12〕出处只留在稿子里
        meta = {}
        if text.startswith("---"):
            head, _, text = text[3:].partition("\n---")
            for l in head.strip().splitlines():
                k, _, v = l.partition(":"); meta[k.strip()] = v.strip()
        lines = text.splitlines()
        self.meta = meta
        self.header_footer(meta)
        self.toc(lines)
        buf = []
        i = 0
        while i < len(lines):
            st = lines[i].strip(); i += 1
            if st.startswith("@单元 "):
                self.render(buf); buf = []; self.unit = st[4:].strip()
            elif st.startswith("# "):
                self.render(buf); buf = []; self.chapter(st[2:])
            elif st.startswith("## "):
                self.render(buf); buf = []; self.section(st[3:])
            elif st == "[分页]":
                self.render(buf); buf = []; self.pending_break = True
            else:
                buf.append(lines[i - 1])
        self.render(buf)

    def header_footer(self, meta):
        h = D["页眉"]
        sec = self.doc.sections[0]
        hp = sec.header.paragraphs[0]; para_fmt(hp, line=None)
        for pos in (4153, 8306, 4680, 9360):
            hp.paragraph_format.tab_stops.add_tab_stop(Emu(pos * 635), WD_TAB_ALIGNMENT.CLEAR)
        hp.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
        r = hp.add_run(meta.get("页眉", "")); set_font(r, KAI, LATIN, 8.5, False, "soft")
        r = hp.add_run("\t" + f"{self.mode}版"); set_font(r, FONT[h["版本字体"]], LATIN, 8.5, h["版本加粗"], "main")
        pPr = hp._p.get_or_add_pPr()                      # 页眉下面粗细双线（参考页眉）
        bd = OxmlElement("w:pBdr"); x = OxmlElement("w:bottom")
        for k, v in (("val", "thickThinSmallGap"), ("sz", "12"), ("space", "3"), ("color", COL["deep"])):
            x.set(qn("w:" + k), v)
        bd.append(x); pPr.append(bd)
        fp = sec.footer.paragraphs[0]; para_fmt(fp, line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = fp.add_run("第 "); set_font(r, SERIF, LATIN, 9, False, "grey")
        K.field(fp, "PAGE", SERIF, LATIN, 9, COL["deep"])
        r = fp.add_run(" 页 / 共 "); set_font(r, SERIF, LATIN, 9, False, "grey")
        K.field(fp, "NUMPAGES", SERIF, LATIN, 9, COL["deep"])
        r = fp.add_run(" 页"); set_font(r, SERIF, LATIN, 9, False, "grey")

    def toc(self, lines):
        t_ = D["目录"]
        ents = []
        n = 0
        for l in lines:
            l = l.strip()
            if l.startswith("@单元 "):
                ents.append(("big", l[4:].strip()))
            elif l.startswith("# "):
                ents.append(("unit", l[2:]))
            elif l.startswith("## "):
                ents.append(("lesson", l[3:]))
            elif l.startswith("### "):
                n = 0
                ents.append(("sec", re.sub(r"\s+$", "", l[4:])))
            elif l.startswith("@考点"):
                n += 1
                ents.append(("kd", f"考点{n}　" + l[3:].lstrip("：: ").partition("|")[0].strip()))
        self.toc_entries = ents
        p = self.para(before=4, after=10, line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run("目　　录"); set_font(r, TITLE_FONT, LATIN, 24, False, "deep")
        for kind, title in ents:
            pg = str(self.pages.get(title, ""))
            tag, _, name = title.partition("　")
            if kind == "big":
                p = self.para(before=16, after=2, line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
                r = p.add_run(tag + "　" + name); set_font(r, TITLE_FONT, LATIN, 14, False, "main")
                continue
            if kind == "unit":
                p = self.para(before=t_["章段前"], after=t_["章段后"], line=t_["章行距"])
                p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W - 0.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.SPACES)
                r = p.add_run(tag + "　" + name); set_font(r, HEI, LATIN, 12, True, "deep")
                r = p.add_run("\t" + pg); set_font(r, HEI, LATIN, 11, True, "deep")
            elif kind == "lesson":
                p = self.para(before=5, line=21, left=0.6)
                p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W - 0.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                r = p.add_run(tag + "　" + name); set_font(r, HEI, LATIN, 11, True, "black")
                r = p.add_run("\t" + pg); set_font(r, HEI, LATIN, 10.5, True, "black")
            elif kind == "sec":
                p = self.para(before=2, line=18.5, left=1.3)
                p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W - 0.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                r = p.add_run(tag + "　" + name); set_font(r, SERIF, LATIN, 10.5, True, "main")
                r = p.add_run("\t" + pg); set_font(r, SERIF, LATIN, 10.5, False, "black")
            else:
                p = self.para(line=16.5, left=2.0)
                p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W - 0.6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
                r = p.add_run(tag + " " + name); set_font(r, SERIF, LATIN, 9.5, False, "black")
                r = p.add_run("\t" + pg); set_font(r, SERIF, LATIN, 9.5, False, "black")
        p = self.para(before=6, line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run(S["目录提示"][self.mode]); set_font(r, KAI, LATIN, 9, False, "soft")
        self.pending_break = True


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("--only", choices=["背诵", "默写"])
    ap.add_argument("--pdf", action="store_true")
    ap.add_argument("--src", help="内容稿路径（默认 <项目>/内容稿.md）")
    ap.add_argument("--out", help="输出目录（默认 <项目>/成品）")
    a = ap.parse_args()
    import paths
    P = paths.of(a.project)
    proj = P.root
    src = Path(a.src) if a.src else P.内容稿
    out = Path(a.out) if a.out else P.成品
    out.mkdir(parents=True, exist_ok=True)
    name = S["书名"]
    nobreak = D["章"].get("不分页前缀", [])
    for mode in ("背诵", "默写"):
        if a.only and a.only != mode:
            continue
        f = out / f"{name}（{mode}版）.docx"
        b = Builder(mode); b.build(src); b.doc.save(f)
        print(mode, "关键词", b.blanks, "->", f)
        if not a.pdf:
            continue
        pdf, _ = K.export_best(f)
        strip = [b.meta.get("页眉", ""), f"{mode}版"]          # 页眉文字抽字时会夹进标题
        pages = K.find_pages(pdf, b.toc_entries, nobreak=nobreak, strip=strip)
        miss = [t for k, t in b.toc_entries if k != "big" and t not in pages]
        if miss:
            print("  目录没找到页码：", miss)
        b2 = Builder(mode); b2.pages = pages; b2.build(src); b2.doc.save(f)
        pdf, _ = K.export_best(f)
        if K.find_pages(pdf, b.toc_entries, nobreak=nobreak, strip=strip) != pages:
            print("  ⚠ 第二遍页码有变化，再跑一次")
        sys.path.insert(0, str(HERE))
        import 封面 as cover_mod
        cover = cover_mod.make(proj, mode, b.blanks) or P.封面 / f"封面_{mode}版.pdf"   # 有 封面模板.html 就重新生成
        if cover.exists():
            K.attach_cover(cover, pdf)
        print("  pdf:", pdf)


if __name__ == "__main__":
    main()
