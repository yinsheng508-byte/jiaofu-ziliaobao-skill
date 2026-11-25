"""素材入库：把用户给的资料逐份编号（S01、S02…），按页抽成文字，生成《素材清单.md》。原文件不移动。

  python3 ingest.py <项目目录> <资料文件或目录>...

- PDF：pdftotext 逐页抽字，每页前加「===== 第 N 页 =====」，方便内容稿里标〔S01 p3〕
  每页平均不到 40 字的视为扫描件：渲染页图到 素材/S01_页图/，由 Agent 用 Read 看图
- Word：.docx 直接读段落和表格；.doc/.wps 先用 LibreOffice 转成 .docx
- 图片：列进清单，标「需看图」
- txt / md：原样复制
"""
import logging
import re
import shutil
import subprocess
import sys
from pathlib import Path

logging.getLogger("pypdf").setLevel(logging.ERROR)

DOC_EXT = {".pdf", ".docx", ".doc", ".wps", ".txt", ".md"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".webp", ".heic"}


def collect(paths):
    files = []
    for p in map(Path, paths):
        if p.is_dir():
            files += sorted(f for f in p.rglob("*") if f.is_file() and f.suffix.lower() in DOC_EXT | IMG_EXT
                            and not f.name.startswith(("~$", ".")) and f.name != "放什么.txt")
        elif p.is_file():
            files.append(p)
    return files


# 输入资料/<资料包名称>/ 下用户提供的三类东西（文件夹名带不带编号都认）
ROLES = (("课本", ("1_教材课本", "教材课本", "电子课本", "课本", "教材")),
         ("对标资料", ("2_对标资料", "对标资料", "内容源")),
         ("参考版式", ("3_参考版式", "参考版式", "版式参考", "参考")))


def classify(paths):
    """返回 (要入库的 [(文件, 角色)], 参考版式文件)。课本排在最前面（编号 S01 起），参考版式只当版式参考、不入库。
    输入文件夹里没有这三个子文件夹时，整个文件夹都当对标资料（老用法）。"""
    items, refs = [], []
    for p in map(Path, paths):
        sub = {role: next((p / n for n in names if (p / n).is_dir()), None) for role, names in ROLES} if p.is_dir() else {}
        if any(sub.values()):
            for role in ("课本", "对标资料"):
                if sub[role]:
                    items += [(f, role) for f in collect([sub[role]])]
            if sub["参考版式"]:
                refs += [f for f in sorted(sub["参考版式"].rglob("*")) if f.is_file() and not f.name.startswith((".", "~$")) and f.name != "放什么.txt"]
        else:
            items += [(f, "") for f in collect([p])]
    items.sort(key=lambda x: x[1] != "课本")      # 课本先编号
    return items, refs


def pdf_text(pdf):
    if shutil.which("pdftotext") and shutil.which("pdfinfo"):
        n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(
            ["pdfinfo", str(pdf)], capture_output=True, text=True, encoding="utf-8", errors="replace"
        ).stdout)[1])
        pages = []
        for i in range(1, n + 1):
            t = subprocess.run(["pdftotext", "-layout", "-f", str(i), "-l", str(i), str(pdf), "-"],
                               capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
            pages.append(t.rstrip())
        return pages
    from pypdf import PdfReader
    return [(p.extract_text() or "").rstrip() for p in PdfReader(str(pdf)).pages]


def docx_text(path):
    """按正文原顺序抽字：表格就地展开，文本框里的字跟在所在段落后面（标〔文本框〕）。
    旧版把表格全挪到文末、文本框整段丢失，a/an 规则这类知识点抽不出来。"""
    from docx import Document
    from docx.oxml.ns import qn
    W, MC = qn("w:t"), "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"

    def run_text(el, skip_txbx=True):
        out = []
        for node in el.iter():
            if node.tag == W and not any(a.tag in (qn("w:txbxContent"), MC) for a in node.iterancestors()
                                         if skip_txbx and a is not el):
                out.append(node.text or "")
        return "".join(out)

    def para(p, lines):
        lines.append(run_text(p))
        for box in p.iter(qn("w:txbxContent")):
            if any(a.tag == MC for a in box.iterancestors()):
                continue                       # 同一文本框的 VML 备份，跳过免得重复
            for bp in box.iter(qn("w:p")):
                t = "".join(n.text or "" for n in bp.iter(W))
                if t.strip():
                    lines.append("〔文本框〕" + t)

    def table(tb, lines):
        lines.append("[表格]")
        for tr in tb.findall(qn("w:tr")):
            cells = []
            for tc in tr.findall(qn("w:tc")):
                cells.append(" ".join(run_text(p) for p in tc.iter(qn("w:p"))).strip())
            lines.append(" | ".join(cells))
        lines.append("[/表格]")

    doc = Document(str(path))
    lines = []
    for el in doc.element.body:
        if el.tag == qn("w:p"):
            para(el, lines)
        elif el.tag == qn("w:tbl"):
            table(el, lines)
    return "\n".join(lines)


def to_docx(path, tmp):
    path = Path(path).resolve()
    tmp = Path(tmp).resolve()
    soffice = shutil.which("soffice")
    if not soffice and sys.platform == "win32":
        for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if Path(p).exists():
                soffice = p
                break
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "docx", "--outdir", str(tmp), str(path)],
                       check=True, capture_output=True, timeout=300)
        return tmp / (path.stem + ".docx")
    if sys.platform == "win32":
        out = tmp / (path.stem + ".docx")
        src_s = str(path).replace("'", "''")
        out_s = str(out).replace("'", "''")
        ps = (
            "$word = New-Object -ComObject Word.Application; "
            "$word.Visible = $false; $word.DisplayAlerts = 0; "
            "try { "
            f"$doc = $word.Documents.Open('{src_s}', $false, $true); "
            f"$doc.SaveAs([ref]'{out_s}', [ref]16); "
            "$doc.Close($false) "
            "} finally { $word.Quit() }"
        )
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True, timeout=240)
        return out
    raise FileNotFoundError("未找到 LibreOffice (soffice) 或 Microsoft Word，无法转换 .doc/.wps")


def ingest(project, inputs):
    import paths
    project = Path(project)
    P = paths.P(project, create=True)
    src_dir = P.素材
    src_dir.mkdir(parents=True, exist_ok=True)
    existing = sorted(src_dir.glob("S[0-9][0-9]_*"))
    start = len({p.name[:3] for p in existing}) + 1
    rows = []
    items, refs = classify(inputs)
    for i, (f, role) in enumerate(items, start):
        sid = f"S{i:02d}"
        ext = f.suffix.lower()
        out = src_dir / f"{sid}_{f.stem}.txt"
        kind, pages, chars, status = ext[1:], "", 0, "已抽字"
        if ext == ".pdf":
            ps = pdf_text(f)
            pages = len(ps)
            chars = sum(len(re.sub(r"\s", "", p)) for p in ps)
            out.write_text("\n\n".join(f"===== 第 {k} 页 =====\n{p}" for k, p in enumerate(ps, 1)), encoding="utf-8")
            if pages and chars / pages < 40:
                img = src_dir / f"{sid}_页图"
                img.mkdir(exist_ok=True)
                from export_pdf import pdf_to_png
                pdf_to_png(f, img / "p", 100)
                status = f"扫描件，需看图（{img.name}/）"
        elif ext in (".docx", ".doc", ".wps"):
            path = f if ext == ".docx" else to_docx(f, src_dir)
            text = docx_text(path)
            chars = len(re.sub(r"\s", "", text))
            out.write_text(text, encoding="utf-8")
            if path != f:
                path.unlink(missing_ok=True)
        elif ext in (".txt", ".md"):
            text = f.read_text(encoding="utf-8", errors="ignore")
            chars = len(re.sub(r"\s", "", text))
            out.write_text(text, encoding="utf-8")
        else:
            out = src_dir / f"{sid}_{f.name}"
            shutil.copyfile(f, out)
            kind, status = "图片", "需看图"
        rows.append((sid, f.name, f"{kind}（{role}）" if role else kind, pages, chars, status, str(f)))

    lst = P.素材清单
    head = ["# 素材清单", "", "| 编号 | 文件 | 类型 | 页数 | 字数 | 状态 | 原路径 |", "|---|---|---|---|---|---|---|"]
    old = lst.read_text(encoding="utf-8").splitlines()[4:] if lst.exists() else []
    body = old + [f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} | {r[4]} | {r[5]} | {r[6]} |" for r in rows]
    lst.write_text("\n".join(head + body) + "\n", encoding="utf-8")
    if refs:                                          # 参考版式：拷进 1_规划/参考/，第 ① 步拆参考时看
        P.参考.mkdir(parents=True, exist_ok=True)
        for f in refs:
            shutil.copy2(f, P.参考 / f.name)
        print(f"参考版式 {len(refs)} 个 -> {P.参考}（只看版式，不当内容来源）")
    import author   # 扫原作者信息候选 → 原作者信息.txt（成品不能留原作者信息）
    f, n = author.suggest(project)
    print(f"原作者信息候选 {n} 条 -> {f}（确认后把名字写进「屏蔽词」区）")
    return rows


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    for r in ingest(sys.argv[1], sys.argv[2:]):
        print(f"{r[0]}  {r[1]}  {str(r[3]) + '页  ' if r[3] else ''}{r[4]}字  {r[5]}")
