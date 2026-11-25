"""教辅资料包统一入口。

  python3 run.py setup                                    首次安装：装 Python 依赖和 /手写笔记 快捷指令，再自检（详见 references/安装配置.md）
  python3 run.py doctor                                   环境自检
  python3 run.py prefs [<项目>]                            看偏好（没有就给出模板）
  python3 run.py template list|use|save                    模板库：列出 / 套用到项目 / 把项目沉淀成模板
  python3 run.py theme list|save                           主题库
  python3 run.py new   <资料文件或目录>... --title 标题      建项目 + 素材入库 → 打印项目目录
  python3 run.py add   <项目> <资料文件或目录>...            追加素材（编号接着往后排）
  python3 run.py check <项目>                              检查内容稿 → 检查报告.md
  python3 run.py build <项目>                              项目有 排版/build.py 就跑它，再出预览、闸门、交付检查；没有就用经典引擎：
  python3 run.py build <项目> [--variants 完整版,考点模板] [--engine lo|word] [--allow-gaps] [--show-sources]
                                                          检查 → 出 Word（目录页码自动回填）→ 导出 PDF → 渲染预览 → 交付说明
  python3 run.py pdf   <项目> [--engine word]              只把 成品/ 里现有的 .docx 重新导出 PDF（Word 里手动精排过之后用，不重新生成 Word）
  python3 run.py gates <项目>                              成品闸门：重开校验、版本页数一致、填充率、出处核对 → 闸门报告.md
  python3 run.py audit <项目>                              交付检查：成品（Word/PDF/手写笔记/文件名/文件属性）里有没有原作者信息 → 交付检查报告.md
  python3 run.py scrub <项目>                              只清 成品/ 里 Word、PDF 的作者、最后修改人等属性（audit 前会自动清）
  python3 run.py render <项目> [--file 文件名.pdf] [--pages 3-5] [--dpi 100]   重新渲染预览图
  python3 run.py assemble <项目> <章键>... [--answers chapter|end]   选题配置/ 里的讲解 + 选题 → 内容稿.md（templates/英语语法手册/取题流程.md）
  python3 run.py verify   <项目> <章键>...                          内容稿每条答案和素材【答案】逐条比对
  python3 run.py qa <stems|scan|sources|grep|pages> <项目> ...       取题、自查小工具（见 scripts/qa.py）
  python3 run.py notebook <手写笔记目录> --name 书名 [--all]   手写笔记多章合订：页码连续、答案放章末 → 合订本.pdf
  python3 run.py mainimg <spec.json> | pages <pdf> <目录> | templates   商品主图 / 详情图（见 templates/商品主图/）
  python3 run.py note <手写稿.md> [--out 目录] [--font 09] [--pages 1-3]   手写笔记排版 → A4 图片 + PDF（见templates/手写笔记/手写笔记规范.md）

不带 --variants 时出全部 5 个版本：完整版、考点模板、练习题、参考答案、背诵默写。
"""
import argparse
import datetime as dt
import json
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
ROOT = SKILL.parent
OUTPUT = ROOT / "输出项目"
sys.path.insert(0, str(HERE))
import subprocess  # noqa: E402

PIP_PKGS = ["python-docx", "pillow", "pypdf", "pypdfium2"]


def _pip(py, pkgs):
    """装 Python 依赖：先 --user；系统 Python 不让装（Mac Homebrew 的 externally-managed）时加 --break-system-packages 再试。"""
    base = [str(py), "-m", "pip", "install", "-q", "--disable-pip-version-check"]
    r = None
    for extra in (["--user"], ["--user", "--break-system-packages"], []):
        r = subprocess.run(base + ["--no-warn-script-location"] + extra + pkgs, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0:
            return True
    print("  pip 报错：", (r.stderr or r.stdout).strip()[-600:])
    return False


def install_shortcut():
    """快捷指令（shortcuts/ 下每个文件夹一个：/手写笔记 /商品主图 /主图）：复制到 ~/.claude/skills/<名>/SKILL.md，
    把 <SKILL_DIR> 换成本机 skill 路径（Mac、Windows 同一种做法）。返回装好的名字。"""
    done = []
    for src in sorted((SKILL / "shortcuts").glob("*/SKILL.md")):
        dst = Path.home() / ".claude" / "skills" / src.parent.name
        if dst.is_symlink():   # 以前用软链接装的（或指向别处的同名 skill）：先断开，免得写回别的文件夹
            dst.unlink()
        dst.mkdir(parents=True, exist_ok=True)
        text = src.read_text(encoding="utf-8")
        (dst / "SKILL.md").write_text(text.replace("<SKILL_DIR>", str(SKILL)), encoding="utf-8")   # 不带 BOM，Claude Code 才读得出 skill 头
        done.append("/" + src.parent.name)
    return "、".join(done)


def setup():
    """首次安装：装 Python 依赖、装 /手写笔记 快捷指令、建 输入资料 / 输出项目，最后环境自检。系统软件（Word、LibreOffice、poppler、浏览器）见 references/安装配置.md。"""
    print("▶ Python 依赖：" + "、".join(PIP_PKGS), flush=True)
    if not _pip(sys.executable, PIP_PKGS):
        venv = SKILL / ".venv"
        print(f"  系统 Python 装不进去，改装到 {venv}（只放 PDF 相关的包）", flush=True)
        subprocess.run([sys.executable, "-m", "venv", str(venv)], check=False)
        vpy = venv / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
        _pip(vpy, ["pypdf", "pypdfium2", "pillow"])
        print("  ⚠️ python-docx 要装在运行 run.py 的这个 Python 里：装不上就换用 python.org 下载的 Python，再重跑 setup")
    print("▶ 快捷指令 ->", install_shortcut(), flush=True)
    OUTPUT.mkdir(exist_ok=True)
    inbox = ROOT / "输入资料"
    inbox.mkdir(exist_ok=True)
    tpl = SKILL / "assets" / "输入资料模板"
    for f in tpl.rglob("*"):                 # 输入资料的说明和「新资料包」模板文件夹（已有的不覆盖）
        t = inbox / f.relative_to(tpl)
        if f.is_dir():
            t.mkdir(parents=True, exist_ok=True)
        elif not t.exists():
            shutil.copy2(f, t)
    print("▶ 环境自检", flush=True)
    return subprocess.run([sys.executable, __file__, "doctor"]).returncode == 0   # 新进程才认得刚装的包



if __name__ == "__main__" and sys.argv[1:2] == ["setup"]:   # setup 在导入 python-docx 等依赖之前执行：这时依赖可能还没装
    sys.exit(0 if setup() else 1)

import build_docx  # noqa: E402
import check as checker  # noqa: E402
import export_pdf  # noqa: E402
import ingest as ingester  # noqa: E402
import render as renderer  # noqa: E402
import author  # noqa: E402
import project  # noqa: E402
import paths  # noqa: E402
import gates  # noqa: E402

FILE_NAMES = {
    "完整版": "{s}（完整版）",
    "考点模板": "【拆分1】{s} - 单考点答题模板",
    "练习题": "【拆分2】{s} - 单练习题",
    "参考答案": "【拆分3】{s} - 练习题参考答案",
    "背诵默写": "{s}背诵与默写",
}

CONTENT_TEMPLATE = """<!-- 内容稿：方案确认后再写，先只写第一章。
套了模板（templates/ 里有 build.py 的，如 知识手册_背诵默写）：按 排版/项目规范.md 的标记写。
用经典引擎（考点答题模板、英语语法手册）：按 templates/考点答题模板/标记写法.md 写，文件头照那里的示例。
每块标出处〔S01 p3〕；素材里没有的写【缺】，不要自己编。 -->
"""


def ok(flag, msg):
    print(("✅ " if flag else "❌ ") + msg)
    return flag


def _pypdf_python():
    try:
        import pypdf  # noqa: F401
        return sys.executable
    except ImportError:
        pass
    for rel in (".venv/Scripts/python.exe", ".venv/bin/python"):
        p = SKILL / rel
        if p.exists() and subprocess.run([str(p), "-c", "import pypdf"], capture_output=True).returncode == 0:
            return str(p)
    return None


def doctor():
    good = True
    try:
        import docx  # noqa: F401
        good &= ok(True, "python-docx")
    except ImportError:
        good &= ok(False, "python-docx 未安装：pip install python-docx")
    try:
        import PIL  # noqa: F401
        ok(True, "Pillow（拼版总览）")
    except ImportError:
        ok(False, "Pillow 未安装：只出单页预览，没有总览图（pip install pillow）")
    pypdf_py = _pypdf_python()
    has_pdftotext = shutil.which("pdftotext") is not None
    if has_pdftotext:
        ok(True, "pdftotext")
    elif pypdf_py:
        ok(True, "PDF 抽字（未装 pdftotext，已自动回落 pypdf；LibreOffice 出的中文 PDF 可能抽不准，建议装 poppler）")
    else:
        good &= ok(False, "pdftotext / pypdf 均未安装（pip install pypdf 或安装 poppler）")
    if shutil.which("pdfinfo") is not None or pypdf_py:
        ok(True, "pdfinfo" if shutil.which("pdfinfo") else "PDF 页数（自动回落 pypdf）")
    else:
        good &= ok(False, "pdfinfo / pypdf 均未安装")
    if shutil.which("pdftoppm"):
        ok(True, "pdftoppm（预览图渲染）")
    else:
        try:
            import pypdfium2  # noqa: F401
            ok(True, "预览图渲染（未装 pdftoppm，已自动回落 pypdfium2）")
        except ImportError:
            good &= ok(False, "pdftoppm / pypdfium2 均未安装（pip install pypdfium2 或安装 poppler）")
    has_lo = export_pdf.find_soffice() is not None
    has_word = export_pdf.find_word() is not None
    has_wps = export_pdf.find_wps() is not None
    has_chrome = export_pdf.find_chrome() is not None
    if has_lo:
        ok(True, "LibreOffice（预览导出、目录页码定位）")
    elif sys.platform == "win32" and has_word:
        ok(True, "LibreOffice 未安装（Windows 已自动回落 Microsoft Word COM 导出）")
    else:
        good &= ok(False, "LibreOffice（预览导出、目录页码定位）")
    ok(has_word, "Microsoft Word（成品导出首选）")
    if has_wps:
        ok(True, "WPS（用户手动精排、导出备选）")
    else:
        print("⚠️ WPS 未安装（可选，已有 Microsoft Word 时不影响）")
    ok(has_chrome, "Google Chrome / Edge（封面 封面.html、手写笔记渲染）")
    ok(pypdf_py is not None, "pypdf（封面拼进 PDF、页码标签、PDF 抽字回落）")
    try:
        mcp_cfg = json.loads((Path.home() / ".claude.json").read_text(encoding="utf-8")).get("mcpServers", {})
    except Exception:
        mcp_cfg = {}
    has_ocr = "paddleocr" in mcp_cfg and shutil.which("uvx") is not None
    print(("✅ " if has_ocr else "⚠️ ") + "PaddleOCR MCP（扫描件 OCR；可选；没接时退回看图手录，接法见 references/安装配置.md 第四节）")
    hw = json.loads((SKILL / "assets" / "手写笔记" / "字体表.json").read_text(encoding="utf-8"))
    miss = [v["name"] for k, v in hw.items() if not k.startswith("_") and not (SKILL / "assets" / "手写笔记" / v["file"]).exists()]
    good &= ok(not miss, "手写笔记字体（随包自带，不用装进系统）" + (f"：缺 {'、'.join(miss)}，重新解压 skill 包" if miss else ""))
    for src in sorted((SKILL / "shortcuts").glob("*/SKILL.md")):
        dst = Path.home() / ".claude" / "skills" / src.parent.name / "SKILL.md"
        mine = dst.exists() and str(SKILL) in dst.read_text(encoding="utf-8", errors="ignore")
        print(("✅ " if mine else "⚠️ ") + f"/{src.parent.name} 快捷指令" + ("" if mine else "（未装或指向别的 skill：运行 run.py setup）"))
    fonts = export_pdf.font_files()
    names = ", ".join(sorted({f.name for f in fonts})) or "无"
    style_fonts = build_docx.F
    good &= ok(bool(fonts), f"排版字体文件：{names}")
    if "Source Han" not in style_fonts["serif"] and "Noto Serif" not in style_fonts["serif"]:
        print("   提示：排版参数.json讲解字体是 %s；装了思源宋体后可改 Source Han Serif CN / Noto Serif SC" % style_fonts["serif"])
    OUTPUT.mkdir(exist_ok=True)
    (ROOT / "输入资料").mkdir(exist_ok=True)
    print("全部就绪" if good else "有 ❌ 先处理")
    return good





def new(inputs, title):
    proj = OUTPUT / f"{dt.date.today():%Y-%m-%d}_{title}"
    proj.mkdir(parents=True, exist_ok=True)
    rows = ingester.ingest(proj, inputs)
    for r in rows:
        print(f"  {r[0]}  {r[1]}  {str(r[3]) + '页  ' if r[3] else ''}{r[4]}字  {r[5]}")
    kinds = " ".join(r[2] for r in rows)
    if "（课本）" not in kinds:
        print("⚠️ 没有教材课本（输入资料/<资料包>/1_教材课本/）：内容只能以对标资料为准，按 references/取材规范.md 第四节处理")
    if not any(paths.P(proj).参考.glob("*")):
        print("提示：没有参考版式（3_参考版式/）：方案阶段从主题库挑风格给用户选")
    stub = paths.P(proj, create=True).内容稿
    if not stub.exists():
        stub.write_text(CONTENT_TEMPLATE, encoding="utf-8")
    project.project_docs(proj)
    print(proj)
    return proj


def run_check(proj):
    P = paths.of(proj)
    verdict, stats, g, s, m, e, w = checker.report(P.内容稿, P.检查报告)
    print(f"{verdict}｜缺 {g}｜无出处 {s}｜题答不符 {m}｜机器味和原作者信息 error {e} warn {w}")
    print("->", P.检查报告)
    return verdict.startswith("✅")


def build(proj, variants, engine, allow_gaps, show_sources, force, sample=False):
    P = paths.of(proj)
    proj = P.root
    content = P.内容稿
    passed = run_check(proj)
    if not passed and not (allow_gaps or force):
        raise SystemExit("检查没通过：先按《检查报告.md》改内容稿；只想看排版加 --allow-gaps（预览版）或 --force")
    meta, _ = build_docx.parse(content)
    short = meta.get("简称", proj.name)
    out_dir = P.样章 if sample else P.成品
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for v in variants:
        base = out_dir / FILE_NAMES[v].format(s=short)
        docx = base.with_suffix(".docx")
        _, sel = build_docx.build(content, docx, v, {}, show_sources, allow_gaps)
        pages = {}
        if v != "背诵默写":
            # Windows 没装 poppler 时，LibreOffice 出的 PDF 用 pypdf 抽中文常常抽不准，目录页码改用 Word 导出的 PDF 定位
            win_no_poppler = sys.platform == "win32" and not shutil.which("pdftotext") and export_pdf.find_word()
            pdf = export_pdf.export_word(docx) if win_no_poppler else export_pdf.export_lo(docx)
            entries = build_docx.toc_entries(sel, meta.get("目录篇目", "是") != "否")
            pages = export_pdf.find_pages(pdf, entries)
            build_docx.build(content, docx, v, pages, show_sources, allow_gaps)
            missing = [t for _, t in entries if t not in pages]
            if missing:
                print(f"  ⚠️ {v} 目录有 {len(missing)} 项没找到页码：{missing[:3]}")
        used = engine
        try:
            pdf = export_pdf.export(docx, engine)
            if engine == "word" and v != "背诵默写":   # Word 和 LibreOffice 分页不同：按 Word 的 PDF 再回填一次目录页码
                for _ in range(2):
                    pages2 = export_pdf.find_pages(pdf, entries)
                    if pages2 == pages:
                        break
                    pages = pages2
                    build_docx.build(content, docx, v, pages, show_sources, allow_gaps)
                    pdf = export_pdf.export(docx, engine)
        except Exception as ex:  # Word 导出失败时退回 LibreOffice，并在交付说明里写明
            print(f"  ⚠️ Word 导出失败（{type(ex).__name__}），本次用 LibreOffice 出预览 PDF")
            pdf, used = export_pdf.export_lo(docx), "lo"
        cover = P.封面 / "封面.html"
        if v == "完整版" and cover.exists():   # 封面只进完整版 PDF（Word 母稿里没有）
            pypdf_py = _pypdf_python() or sys.executable
            subprocess.run([pypdf_py, str(Path(__file__).with_name("cover.py")), str(cover), str(pdf)], check=True)
        n = export_pdf.page_count(pdf)
        results.append((v, docx, pdf, n, used))
        print(f"  {v}：{n} 页  {pdf.name}（{'Word' if used == 'word' else 'LibreOffice 预览'}）")
    preview = P.预览
    first = results[0][2]
    files, sheets = renderer.render(first, preview / results[0][0])
    print(f"预览：{preview / results[0][0]}（{len(files)} 页）")
    for s in sheets:
        print("  ", s)
    if not sample:
        handoff(proj, results, allow_gaps)
    run_gates(proj, out_dir)
    run_audit(proj, out_dir if sample else None)
    return results


def run_gates(proj, out_dir=None):
    ok_, b, w, rep = gates.run(proj, out_dir)
    print(f"{'✅' if ok_ else '❌'} 闸门 拦截 {b} / 警告 {w}  -> {rep}")
    return ok_


def project_build(proj, only=None, no_pdf=False, force=False, sample=False):
    """项目自带 排版/build.py 时：检查 → 跑它 → 出预览图 → 闸门 → 清属性 + 交付检查 → 交付说明。
    sample=True（出样章）：成品放 2_样张/样章/，不写交付说明。"""
    P = paths.of(proj)
    proj = P.root
    out_dir = P.样章 if sample else P.成品
    if not run_check(proj) and not force:
        raise SystemExit("检查没通过：先按《检查报告.md》改内容稿；只想先看排版效果加 --force（出的是预览，不能交付）")
    cmd = [sys.executable, str(P.排版 / "build.py"), str(proj), "--out", str(out_dir)] + (["--only", only] if only else []) + ([] if no_pdf else ["--pdf"])
    env = {**os.environ, "JIAOFU_SKILL": str(SKILL), "PYTHONUTF8": "1"}
    r = subprocess.run(cmd, env=env)
    if r.returncode:
        raise SystemExit("排版/build.py 出错了：按报错改 build.py 或内容稿，不要出半套文件")
    pdfs = sorted(out_dir.glob("*.pdf"))
    for pdf in pdfs:
        files, sheets = renderer.render(pdf, P.预览 / pdf.stem)
        print(f"预览：{P.预览 / pdf.stem}（{len(files)} 页）")
    if sample:
        run_gates(proj, out_dir)
        run_audit(proj, out_dir)
        print(f"样章在 {out_dir}；按 references/需求与方案.md 第七节写 {P.样章汇报.name}，交给用户确认")
        return
    L = ["# 交付说明", "", "| 文件 | 页数 |", "|---|---|"] + [f"| {p.name} | {export_pdf.page_count(p)} |" for p in pdfs]
    L += ["", "## 留给你做的", "", "- 用 Word 或 WPS 打开 .docx 逐页看一遍；在 Word 里手动改过的，改完跑 `run.py pdf <项目>` 重新导出，不要再 build（会覆盖）。"]
    P.交付说明.write_text("\n".join(L) + "\n", encoding="utf-8")
    run_gates(proj)
    run_audit(proj)


def run_audit(proj, sample_dir=None):
    cleaned = author.scrub(proj)   # Word 里精排保存后「最后修改人」会写进本机用户名，导出的 PDF 也会带：先清属性再查
    if cleaned:
        print(f"  已清空 {len(cleaned)} 个文件的作者 / 最后修改人属性")
    ok_, e, w, out = author.audit(proj, [sample_dir], Path(sample_dir).parent / "样章检查报告.md") if sample_dir else author.audit(proj)
    print(f"{'✅' if ok_ else '❌'} 交付检查（原作者信息）error {e} / warn {w}  -> {out}")
    if not ok_:
        print("   有残留：先改内容稿 / 图片 / 文件名，重新导出后再跑 run.py audit，通过前不要交付")
    return ok_


def handoff(proj, results, allow_gaps):
    L = ["# 交付说明", "", "| 版本 | 页数 | Word | PDF | PDF 导出方式 |", "|---|---|---|---|---|"]
    for v, docx, pdf, n, used in results:
        L.append(f"| {v} | {n} | {docx.name} | {pdf.name} | {'Word' if used == 'word' else 'LibreOffice（预览）'} |")
    L += ["", "## 留给你做的", ""]
    if any(r[4] != "word" for r in results):
        L.append("- PDF 是 LibreOffice 出的预览版。用 Word 或 WPS 打开 .docx，逐页看一遍，再「文件 → 另存为 → PDF」出正式版；"
                 "在 Word 里手动改过的，不要再跑 build（会覆盖），改完跑 `run.py pdf <项目> --engine word` 或直接在 Word 里导出。")
    if allow_gaps:
        L.append("- 内容稿里还有【缺】，成品里显示为红字，补齐后重新 build。")
    L.append("- 抽查《检查报告.md》里的 warn 项，以及每个单元第一页和答案页。")
    out = paths.of(proj).交付说明
    out.write_text("\n".join(L) + "\n", encoding="utf-8")
    print("->", out)


def pdf_only(proj, engine):
    P = paths.of(proj)
    for docx in sorted(P.成品.glob("*.docx")):
        if docx.name.startswith("~$"):
            continue
        try:
            pdf = export_pdf.export(docx, engine)
        except Exception as ex:
            raise SystemExit(f"{docx.name} 导出失败：{ex}。请在 Word/WPS 里手动另存为 PDF")
        cover = P.封面 / "封面.html"
        if "完整版" in docx.name and cover.exists():   # 和 build 一样把封面拼到完整版最前面
            pypdf_py = _pypdf_python() or sys.executable
            subprocess.run([pypdf_py, str(Path(__file__).with_name("cover.py")), str(cover), str(pdf)], check=True)
        print(f"  {pdf.name}  {export_pdf.page_count(pdf)} 页")
    run_audit(proj)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("setup")
    sub.add_parser("doctor")
    p = sub.add_parser("new")
    p.add_argument("inputs", nargs="+")
    p.add_argument("--title", required=True)
    p = sub.add_parser("add")
    p.add_argument("project")
    p.add_argument("inputs", nargs="+")
    p = sub.add_parser("check")
    p.add_argument("project")
    p = sub.add_parser("build")
    p.add_argument("project")
    p.add_argument("--variants", default=",".join(build_docx.VARIANTS))
    p.add_argument("--engine", default=("word" if sys.platform == "win32" and export_pdf.find_word() else "lo"), choices=["lo", "word"])   # Windows 有 Word 就直接用 Word 出正式 PDF
    p.add_argument("--allow-gaps", action="store_true")
    p.add_argument("--show-sources", action="store_true")
    p.add_argument("--force", action="store_true")
    p.add_argument("--only", help="项目自带 build.py 时：只出某个版本（如 背诵）")
    p.add_argument("--no-pdf", action="store_true", help="项目自带 build.py 时：只出 Word，不导 PDF")
    p.add_argument("--sample", action="store_true", help="出样章：成品放 2_样张/样章/，不写交付说明（确认点二用）")
    p = sub.add_parser("pdf")
    p.add_argument("project")
    p.add_argument("--engine", default="word", choices=["lo", "word"])
    p = sub.add_parser("audit")
    p.add_argument("project")
    p = sub.add_parser("prefs")
    p.add_argument("project", nargs="?")
    p = sub.add_parser("template")
    p.add_argument("action", choices=["list", "use", "save"])
    p.add_argument("args", nargs="*")
    p = sub.add_parser("theme")
    p.add_argument("action", choices=["list", "save"])
    p.add_argument("args", nargs="*")
    p = sub.add_parser("scrub")
    p.add_argument("project")
    p = sub.add_parser("gates")
    p.add_argument("project")
    p = sub.add_parser("render")
    p.add_argument("project")
    p.add_argument("--file")
    p.add_argument("--pages")
    p.add_argument("--dpi", type=int, default=70)
    for name in ("assemble", "verify", "qa", "note", "notebook", "mainimg"):
        p = sub.add_parser(name, add_help=False)
        p.add_argument("rest", nargs=argparse.REMAINDER)
    a = ap.parse_args()

    if a.cmd in ("assemble", "verify", "qa", "note", "notebook", "mainimg"):   # 直接转给同名脚本
        script = {"note": "handnote", "notebook": "handnote_book"}.get(a.cmd, a.cmd)
        env = {**os.environ, "PYTHONUTF8": "1"}
        sys.exit(subprocess.run([sys.executable, str(Path(__file__).with_name(f"{script}.py")), *a.rest], env=env).returncode)
    if a.cmd == "doctor":
        sys.exit(0 if doctor() else 1)
    if a.cmd == "new":
        new(a.inputs, a.title)
    elif a.cmd == "add":
        for r in ingester.ingest(a.project, a.inputs):
            print(f"  {r[0]}  {r[1]}  {str(r[3]) + '页  ' if r[3] else ''}{r[4]}字  {r[5]}")
    elif a.cmd == "check":
        sys.exit(0 if run_check(a.project) else 1)
    elif a.cmd == "build" and (paths.of(a.project).排版 / "build.py").exists():
        project_build(a.project, a.only, a.no_pdf, a.force or a.allow_gaps, a.sample)
    elif a.cmd == "build":
        variants = [v.strip() for v in a.variants.split(",") if v.strip()]
        bad = [v for v in variants if v not in build_docx.VARIANTS]
        if bad:
            raise SystemExit(f"未知版本：{bad}，可选 {build_docx.VARIANTS}")
        build(a.project, variants, a.engine, a.allow_gaps, a.show_sources, a.force, a.sample)
    elif a.cmd == "pdf":
        pdf_only(a.project, a.engine)
    elif a.cmd == "audit":
        sys.exit(0 if run_audit(a.project) else 1)
    elif a.cmd == "prefs":
        project.prefs(a.project)
    elif a.cmd == "template":
        if a.action == "list":
            project.template_list()
        elif a.action == "use" and len(a.args) == 2:
            project.template_use(a.args[0], a.args[1])
        elif a.action == "save" and len(a.args) == 2:
            project.template_save(a.args[0], a.args[1])
        else:
            raise SystemExit("用法：template list | template use <模板名> <项目> | template save <项目> <模板名>")
    elif a.cmd == "theme":
        if a.action == "list":
            project.theme_list()
        elif len(a.args) == 2:
            project.theme_save(a.args[0], a.args[1])
        else:
            raise SystemExit("用法：theme list | theme save <项目> <主题名>")
    elif a.cmd == "scrub":
        for f in author.scrub(a.project):
            print("  已清属性", f.name)
    elif a.cmd == "gates":
        sys.exit(0 if run_gates(a.project) else 1)
    elif a.cmd == "render":
        proj = Path(a.project)
        P = paths.of(proj)
        pdf = P.成品 / a.file if a.file else next(iter(sorted(P.成品.glob("*（完整版）.pdf"))), None)
        if not pdf or not pdf.exists():
            raise SystemExit("成品/ 里没有 PDF，先跑 build")
        files, sheets = renderer.render(pdf, P.预览 / pdf.stem, a.dpi, a.pages)
        print(f"{len(files)} 页 -> {P.预览 / pdf.stem}")
        for s in sheets:
            print("  ", s)


if __name__ == "__main__":
    main()
