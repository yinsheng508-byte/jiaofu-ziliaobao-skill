"""封面：<项目>/封面/封面模板.html → 封面_背诵版.pdf / 封面_默写版.pdf（Chrome / Edge 无界面导出）。
build.py --pdf 时自动调用；也可以单独跑：python3 排版/封面.py <项目> <关键词数>

模板里的占位符：
  {{书名}} {{学科}} {{年级}} {{教材版本}} {{卖点1}}… 取 样式.json「封面.通用」（换学科改这里，不用改 HTML）
  {{版}} {{印}} {{口号}} {{卖点2}} {{卖点2说明}}   按版本取 样式.json「封面.背诵 / 默写」（不写用默认）
  {{空数}}  关键词个数（build 统计）      {{图数}}  排版/插图/ 里的图片数
  {{装饰1}} {{装饰2}} 等没填的占位符        有同名 .svg（<项目>/封面/装饰1.svg）就嵌入，没有就留空
封面只写真实信息：不编出版社、书号、价格、推荐语，不写原作者名和原系列名。
"""
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
S = json.loads((HERE / "样式.json").read_text(encoding="utf-8"))

DEFAULT = {
    "背诵": {"版": "背诵版", "印": "<span>背</span><span>诵</span>", "口号": "一背就会",
             "卖点2": "关键词红字下划线", "卖点2说明": "考点一眼看清，和默写版逐页对应"},
    "默写": {"版": "默写版", "印": "<span>默</span><span>写</span>", "口号": "一默就清",
             "卖点2": "关键词原位挖空", "卖点2说明": "页码与背诵版一一对应，默完即可对照订正"},
}


def _chrome():
    import os
    scripts = Path(os.environ.get("JIAOFU_SKILL", HERE.parent.parent)) / "scripts"
    sys.path.insert(0, str(scripts))
    import export_pdf
    c = export_pdf.find_chrome()
    if not c:
        raise SystemExit("没有 Chrome / Edge，封面导不出")
    return c


def make(proj, mode, blanks):
    import os
    sys.path.insert(0, str(Path(os.environ.get("JIAOFU_SKILL", HERE.parent.parent)) / "scripts"))
    import paths
    cover_dir = paths.of(proj).封面
    tpl_file = cover_dir / "封面模板.html"
    if not tpl_file.exists():
        tpl_file = HERE / "封面" / "封面模板.html"        # 套模板时复制到了 排版/封面/
    if not tpl_file.exists() or S.get("封面") is False:
        return None
    cover_dir.mkdir(parents=True, exist_ok=True)
    tpl = tpl_file.read_text(encoding="utf-8")
    cfg = S.get("封面", {})
    vals = {**DEFAULT[mode], **cfg.get("通用", {}), **cfg.get(mode, {})}   # 优先级：版本专用 > 通用 > 默认
    vals["空数"] = str(blanks)
    vals["图数"] = str(len(list((HERE / "插图").glob("*.png")))) if (HERE / "插图").exists() else "0"
    for _ in range(2):                       # 两遍：卖点里还可以再引用 {{图数}}
        for k, v in vals.items():
            tpl = tpl.replace("{{" + k + "}}", v)
    for k in set(re.findall(r"\{\{([^}]+)\}\}", tpl)):
        svg = cover_dir / f"{k}.svg"
        tpl = tpl.replace("{{" + k + "}}", svg.read_text(encoding="utf-8") if svg.exists() else "")
    html = cover_dir / f"封面_{mode}版.html"
    html.write_text(tpl, encoding="utf-8")
    pdf = html.with_suffix(".pdf")
    subprocess.run([_chrome(), "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--allow-file-access-from-files",
                    f"--print-to-pdf={pdf}", html.resolve().as_uri()], check=True, capture_output=True)
    return pdf


if __name__ == "__main__":
    for m in ("背诵", "默写"):
        print(make(sys.argv[1], m, sys.argv[2] if len(sys.argv) > 2 else "0"))
