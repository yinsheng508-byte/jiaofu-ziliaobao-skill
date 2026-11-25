"""Word 排版积木：所有 Word 类模板共用。模板的 build.py 这样用：

    import sys; sys.path.insert(0, str(<skill>/scripts))
    import docx_kit as K
    T = K.load_theme(<样式.json 或 主题名>)      # 颜色、字号、细节参数
    F = K.FONTS                                  # 按系统自动选好的字体：serif / hei / kai / latin / title

积木（都是 python-docx 的薄封装，产出的是真正的 Word 样式、表格、域，打开 .docx 都能改）：
  set_font(run, east, latin, size, bold, color_hex)     中西文字体分开设
  para_fmt(p, before, after, line, left, first, align, keep_next)   段落格式（line=None 用单倍行距）
  shd(el, fill)                  段落 / 单元格 / 文字底纹
  p_border(p, side, color, sz, space)   段落边线（标题下划线、页眉线）
  borders(tbl, spec)             表格边框：{"top": (8, "1F6A73"), "insideH": (2, "E2EAEB")}
  cell_border / cell_margins / fixed_widths / tbl_indent / no_split / keep_rows / keep_rows_head / repeat_header
  field(p, "PAGE" | "NUMPAGES", font_args)      页码域
  setup_page(doc, margins, header_dist, footer_dist, body_size)   A4 页面 + 正文默认字体
  block_end(lines, i, tag)       找 [tag] … [/tag] 的配对结束行
  find_pages(pdf, titles)        在导出的 PDF 里找每个标题所在页（回填目录用，Mac / Windows 通用）
  export_best(docx)              导出 PDF：优先 Word，失败退回 LibreOffice
  attach_cover(cover_pdf, pdf)   把封面 PDF 拼到最前面，页码标签设成「封面、1、2…」

字体按系统自动换：Mac 用 Songti SC / Heiti SC / Kaiti SC，Windows 用 SimSun / SimHei / KaiTi；
标题字体默认方正大标宋（FZDaBiaoSong-B06），没装时 Word 会自动换成系统宋体。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
THEMES = SKILL / "assets" / "主题"

if sys.platform == "win32":
    FONTS = {"serif": "SimSun", "hei": "SimHei", "kai": "KaiTi", "latin": "Times New Roman", "title": "FZDaBiaoSong-B06"}
else:
    FONTS = {"serif": "Songti SC", "hei": "Heiti SC", "kai": "Kaiti SC", "latin": "Times New Roman", "title": "FZDaBiaoSong-B06"}


def load_theme(src):
    """src：主题名（assets/主题/<名>.json）或 json 文件路径。样式.json 里可以写 "主题": "<名>"，其余键覆盖主题。"""
    p = Path(src)
    if not p.suffix:
        p = THEMES / f"{src}.json"
    data = json.loads(p.read_text(encoding="utf-8"))
    base = data.get("主题")
    if base:
        merged = load_theme(base)
        for k, v in data.items():
            if isinstance(v, dict) and isinstance(merged.get(k), dict):
                merged[k] = {**merged[k], **v}
            else:
                merged[k] = v
        data = merged
    return data


def hexcolor(theme_colors, key):
    return theme_colors.get(key, key)


def set_font(run, east, latin, size=None, bold=None, color=None):
    rpr = run._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), latin)
    rf.set(qn("w:eastAsia"), east)
    if size: run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color: run.font.color.rgb = RGBColor.from_string(color)


def shd(el, fill):
    tag = el.tag
    pr = el.get_or_add_tcPr() if tag.endswith("}tc") else (el.get_or_add_pPr() if tag.endswith("}p") else el.get_or_add_rPr())
    s = OxmlElement("w:shd"); s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto"); s.set(qn("w:fill"), fill)
    pr.append(s)


def para_fmt(p, before=0, after=0, line=None, left=0.0, first=0.0, align=None, keep_next=False):
    f = p.paragraph_format
    f.space_before, f.space_after = Pt(before), Pt(after)
    if line:
        f.line_spacing_rule = WD_LINE_SPACING.EXACTLY; f.line_spacing = Pt(line)
    f.left_indent = Cm(left); f.first_line_indent = Cm(first)
    if align is not None: p.alignment = align
    f.keep_with_next = keep_next
    f.widow_control = True
    sg = OxmlElement("w:snapToGrid"); sg.set(qn("w:val"), "0"); p._p.get_or_add_pPr().append(sg)


def p_border(p, side="bottom", color="C5D3CB", sz=4, space=3):
    pPr = p._p.get_or_add_pPr()
    bd = pPr.find(qn("w:pBdr"))
    if bd is None:
        bd = OxmlElement("w:pBdr"); pPr.append(bd)
    x = OxmlElement(f"w:{side}"); x.set(qn("w:val"), "single"); x.set(qn("w:sz"), str(sz))
    x.set(qn("w:space"), str(space)); x.set(qn("w:color"), color); bd.append(x)


def borders(tbl, spec):
    b = OxmlElement("w:tblBorders")
    for e in ("top", "left", "bottom", "right", "insideH", "insideV"):
        x = OxmlElement(f"w:{e}")
        v = spec.get(e)
        if v:
            x.set(qn("w:val"), "single"); x.set(qn("w:sz"), str(v[0])); x.set(qn("w:color"), v[1])
        else:
            x.set(qn("w:val"), "nil")
        b.append(x)
    tbl._tbl.tblPr.append(b)


def cell_border(cell, **sides):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for side, (sz, color) in sides.items():
        x = OxmlElement(f"w:{side}"); x.set(qn("w:val"), "single"); x.set(qn("w:sz"), str(sz)); x.set(qn("w:color"), color)
        b.append(x)
    tcPr.append(b)


def cell_margins(tbl, top=40, bottom=40, left=80, right=80):
    m = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        x = OxmlElement(f"w:{k}"); x.set(qn("w:w"), str(v)); x.set(qn("w:type"), "dxa"); m.append(x)
    tbl._tbl.tblPr.append(m)


def fixed_widths(tbl, widths_cm):
    tbl.autofit = False
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tbl._tbl.tblPr.append(lay)
    w = OxmlElement("w:tblW"); w.set(qn("w:w"), str(int(sum(widths_cm) * 567))); w.set(qn("w:type"), "dxa")
    old = tbl._tbl.tblPr.find(qn("w:tblW"))
    if old is not None: tbl._tbl.tblPr.remove(old)
    tbl._tbl.tblPr.append(w)
    for row in tbl.rows:
        for c, wd in zip(row.cells, widths_cm):
            c.width = Cm(wd)
    for gc, wd in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(wd * 567)))


def tbl_indent(tbl, w=0):
    x = OxmlElement("w:tblInd"); x.set(qn("w:w"), str(w)); x.set(qn("w:type"), "dxa"); tbl._tbl.tblPr.append(x)


def no_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def repeat_header(tbl):
    tbl.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def keep_rows_head(tbl):
    for c in tbl.rows[0].cells:
        for q in c.paragraphs: q.paragraph_format.keep_with_next = True


def keep_rows(tbl):
    for row in tbl.rows[:-1]:
        for c in row.cells:
            for q in c.paragraphs: q.paragraph_format.keep_with_next = True


def field(p, code, east, latin, size, color):
    r = p.add_run(); set_font(r, east, latin, size, False, color)
    a = OxmlElement("w:fldChar"); a.set(qn("w:fldCharType"), "begin"); r._r.append(a)
    r2 = p.add_run(); set_font(r2, east, latin, size, False, color)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = code; r2._r.append(it)
    r3 = p.add_run(); set_font(r3, east, latin, size, False, color)
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "end"); r3._r.append(b)


def setup_page(doc, margins=(1.55, 1.45, 1.6, 1.6), header_dist=0.75, footer_dist=0.7, body_size=10.5):
    """margins = (上, 下, 左, 右) cm。"""
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(margins[0]), Cm(margins[1])
    sec.left_margin, sec.right_margin = Cm(margins[2]), Cm(margins[3])
    sec.header_distance, sec.footer_distance = Cm(header_dist), Cm(footer_dist)
    st = doc.styles["Normal"]
    st.font.size = Pt(body_size)
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs"): rf.set(qn(a), FONTS["latin"])
    rf.set(qn("w:eastAsia"), FONTS["serif"]); st.element.get_or_add_rPr().append(rf)
    return sec


def block_end(lines, i, tag):
    """从 i 开始找与之配对的 [/tag]（允许同名嵌套）"""
    depth = 1
    while i < len(lines):
        s = lines[i].strip()
        if s.startswith(f"[{tag}") and not s.startswith(f"[{tag}]") or s == f"[{tag}]":
            depth += 1
        elif s == f"[/{tag}]":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    raise SystemExit(f"[{tag}] 没有配对的 [/{tag}]")


def _norm(s):
    sys.path.insert(0, str(HERE))
    import export_pdf
    return export_pdf.normalize_cjk(s)


def find_pages(pdf, entries, skip_first=("红字下划线", "横线处默写"), nobreak=(), strip=()):
    """entries = [(kind, title)]，kind 为 big 的跳过；返回 {title: 页码}。第一页（目录页）只看提示语之后的文字。
    strip：每页先去掉的文字（页眉、版本名），免得抽字时夹进标题中间。"""
    sys.path.insert(0, str(HERE))
    import export_pdf
    texts = [_norm(t) for t in export_pdf.extract_page_texts(pdf)]
    for x in (_norm(x) for x in strip if x):
        texts = [t.replace(x, "") for t in texts]
    n = len(texts)
    pages, cur = {}, 0
    for kind, title in entries:
        if kind == "big":
            continue
        tag, _, name = title.partition("　")
        key = _norm(tag + name)
        start = cur if kind != "unit" or tag.startswith(tuple(nobreak)) else max(cur, 1)
        for k in range(start, n):
            body = texts[k]
            if k == 0:
                for s in skip_first:
                    body = body.split(s)[-1]
            if key in body:
                pages[title] = k + 1; cur = k
                break
    return pages


def export_best(docx_path):
    """优先用 Word 导出 PDF（和用户在 Word 里看到的一致）；没有 Word 或失败时退回 LibreOffice 预览版。返回 (pdf, 用的引擎)。"""
    sys.path.insert(0, str(HERE))
    import export_pdf
    try:
        return export_pdf.export(docx_path, "word"), "word"
    except Exception as ex:
        print(f"  Word 导出没成功（{type(ex).__name__}），改用 LibreOffice 出预览版；正式 PDF 请在 Word / WPS 里另存")
        return export_pdf.export(docx_path, "lo"), "lo"


def attach_cover(cover, pdf):
    code = """
import sys, os
from pypdf import PdfReader, PdfWriter
c, b = sys.argv[1], sys.argv[2]
w = PdfWriter(); w.append(PdfReader(c)); w.append(PdfReader(b))
w.set_page_label(0, 0, prefix="封面"); w.set_page_label(1, len(w.pages) - 1, style="/D", start=1)
w.write(b + ".tmp"); os.replace(b + ".tmp", b)
"""
    py = sys.executable
    try:
        import pypdf  # noqa: F401
    except ImportError:
        for rel in (".venv/Scripts/python.exe", ".venv/bin/python"):
            if (SKILL / rel).exists():
                py = str(SKILL / rel)
    subprocess.run([py, "-c", code, str(cover), str(pdf)], check=True)
