"""项目文档、偏好、模板库、主题库。由 run.py 调用：

  run.py prefs [<项目>]                    看偏好（工作文件夹的 偏好.md + 项目里的 偏好.md）；没有就给出模板
  run.py template list                     列出模板库
  run.py template use <模板名> <项目>       把模板复制进 <项目>/排版/
  run.py template save <项目> <模板名>      把 <项目>/排版/ 收进模板库（自动生成《模板说明.md》初稿）
  run.py theme list                        列出主题库
  run.py theme save <项目> <主题名>         把 <项目>/排版/样式.json 里的主题存进主题库

需求单、方案的模板直接取自 references/需求与方案.md 里的代码块，规范和脚本只有一份来源。
"""
import json
import re
import shutil
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
ROOT = SKILL.parent
TEMPLATES = SKILL / "templates"
THEMES = SKILL / "assets" / "主题"
PREFS = ROOT / "偏好.md"
SKIP = {"__pycache__", ".DS_Store", "插图", "课本插图_全部", "页图", "页图300", "网格", "图框", "_html"}

PREFS_TEMPLATE = """# 偏好（开工先读；这次用户说的和这里冲突时以这次为准，并问要不要更新）

常做的用户群：
默认交付形式：Word + PDF / 手写风 / 都要
默认改写档位：整合升级
默认主题：
喜欢的样式：
不喜欢的样式：
页眉 / 系列名习惯：
上传：不上传 / 飞书知识库（知识库、节点：） / ima 知识库（知识库：）
其他习惯：
"""

PROGRESS_TEMPLATE = """# 进度记录（每做完一步更新；换会话后先读这里）

## 需求单摘要
<!-- 给谁用 / 做什么用 / 交付形式 / 版本 / 改写档位 -->

## 三个确认点
- [ ] 确认点一：方案
- [ ] 确认点二：第一章样章
- [ ] 确认点三：交付

## 用户已定的事
<!-- 含还没回答的问题，到该定的节点再问 -->

## 样章定稿标准
<!-- 确认点二后写：每章结构、栏目、讲解深度、例题几道、答案放哪、版式参数改了什么 -->

## 各章进度
| 章 | 内容稿 | check | build | 填充率最低 | 人工清单 | 备注 |
|---|---|---|---|---|---|---|

## 下一步
"""

COVER_TABLE = ("# 素材覆盖表\n\n| 编号 | 覆盖哪些章节 / 考点 | 例题 | 练习 | 答案 | 独有内容 | 质量 / 问题 |\n"
               "|---|---|---|---|---|---|---|\n")


def _ref_blocks():
    text = (SKILL / "references" / "需求与方案.md").read_text(encoding="utf-8")
    return re.findall(r"```\n(.*?)```", text, re.S)


def project_docs(proj):
    """新项目里放好：进度记录（根目录）、需求单 / 方案 / 素材覆盖表（1_规划）（已有的不覆盖）。"""
    import paths
    P = paths.P(proj, create=True)
    blocks = _ref_blocks()
    need = next((b for b in blocks if b.startswith("# 需求单")), "# 需求单\n")
    plan = next((b for b in blocks if b.startswith("# 方案")), "# 方案\n")
    if PREFS.exists():
        need += "\n## 偏好里的默认值（已用的标「默认」）\n" + PREFS.read_text(encoding="utf-8").split("\n", 1)[1]
    for p, text in ((P.需求单, need + "\n## 参考拆解\n"), (P.方案, plan),
                    (P.进度记录, PROGRESS_TEMPLATE), (P.素材覆盖表, COVER_TABLE)):
        if not p.exists():
            p.write_text(text, encoding="utf-8")


# ---------- 偏好 ----------

def prefs(proj=None):
    if not PREFS.exists():
        print(f"还没有偏好文件：{PREFS}\n第一次用：把下面的问题用 AskUserQuestion 一次问完（最多 4 个），答案写进这个文件。\n")
        print(PREFS_TEMPLATE)
        return None
    print(f"== 偏好（{PREFS}）\n" + PREFS.read_text(encoding="utf-8"))
    if proj and (Path(proj) / "偏好.md").exists():
        print(f"== 项目偏好（优先）\n" + (Path(proj) / "偏好.md").read_text(encoding="utf-8"))
    return PREFS


# ---------- 模板库 ----------

def _summary(md):
    if not md.exists():
        return "（没有模板说明）"
    for line in md.read_text(encoding="utf-8").splitlines():
        if line.startswith(("适用：", "适用于：", "> 适用")):
            return line.lstrip("> ").strip()
    return md.read_text(encoding="utf-8").splitlines()[0].lstrip("# ")


def template_list():
    rows = sorted(p for p in TEMPLATES.iterdir() if p.is_dir()) if TEMPLATES.exists() else []
    if not rows:
        print("模板库是空的")
    for p in rows:
        has = "项目 build.py" if (p / "build.py").exists() else "skill 内置引擎"
        print(f"- {p.name}（{has}）：{_summary(p / '模板说明.md')}")
    return rows


def _copytree(src, dst):
    for f in sorted(src.rglob("*")):
        if any(part in SKIP for part in f.relative_to(src).parts) or not f.is_file():
            continue
        t = dst / f.relative_to(src)
        t.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(f, t)


def template_use(name, proj):
    src = TEMPLATES / name
    if not src.is_dir():
        raise SystemExit(f"没有模板「{name}」。先 run.py template list 看看有哪些")
    import paths
    dst = paths.P(proj, create=True).排版
    if dst.exists() and any(dst.iterdir()):
        raise SystemExit(f"{dst} 里已经有文件了；要重新套模板先把它改名备份")
    _copytree(src, dst)
    print(f"已复制模板「{name}」-> {dst}\n下一步：读 {dst / '模板说明.md'}，按方案改 样式.json 和 build.py")
    return dst


def template_save(proj, name):
    import author
    import paths
    P = paths.of(proj)
    proj = P.root
    src = P.排版
    if not src.is_dir():
        raise SystemExit(f"{src} 不存在：项目没有自己的排版工具，不用沉淀")
    dst = TEMPLATES / name
    if dst.exists():
        raise SystemExit(f"模板「{name}」已存在；换个名字，或先把旧的改名")
    _copytree(src, dst)
    shots = [f for f in sorted(P.预览.rglob("p-*.png")) if re.fullmatch(r"p-0*[1-4]", f.stem)][:4]
    if shots:
        (dst / "样章图").mkdir(exist_ok=True)
        for i, s in enumerate(shots, 1):
            shutil.copy2(s, dst / "样章图" / f"样章_{i}.png")
    md = dst / "模板说明.md"
    if md.exists() and not md.read_text(encoding="utf-8").startswith(f"# 模板：{name}"):   # 复制来的是来源模板的说明
        md.rename(dst / "来源模板说明（参考后删除）.md")
    if not md.exists():
        md.write_text(f"""# 模板：{name}

适用：（什么资料、什么用户群，一句话）

## 版本
（这个模板出哪几个版本）

## 文件
- build.py：内容稿 → 成品 Word（`python3 排版/build.py <项目> [--only 版本] [--pdf]`）
- 样式.json：主题和参数
- 项目规范.md：内容稿标记写法

## 样章定稿标准
（从来源项目的进度记录里抄过来）

## 踩过的坑
（来源项目里改过的地方、用户纠正过的地方）

来源项目：{proj.name}
""", encoding="utf-8")
    # 模板里不能留来源项目的原作者信息和本机路径
    terms = author.project_terms(proj)
    leaks = []
    for f in dst.rglob("*"):
        if f.suffix in (".md", ".py", ".json", ".html", ".css", ".txt"):
            t = f.read_text(encoding="utf-8", errors="ignore")
            leaks += [f"{f.name}：{w}" for lv, w, _ in author.scan(t, terms) if lv == "error"]
            if re.search(r"/Users/|C:\\\\Users\\\\|/home/", t):
                leaks.append(f"{f.name}：写死了本机路径")
    print(f"已存成模板「{name}」-> {dst}\n下一步：补全《模板说明.md》（适用、版本、样章定稿标准、踩过的坑）")
    if leaks:
        print("⚠️ 模板里有要清掉的东西：\n  " + "\n  ".join(leaks))
    return dst


# ---------- 主题库 ----------

def theme_list():
    files = sorted(THEMES.glob("*.json")) if THEMES.exists() else []
    if not files:
        print("主题库是空的")
    for f in files:
        t = json.loads(f.read_text(encoding="utf-8"))
        print(f"- {f.stem}：{t.get('风格', '')}｜适合 {t.get('适合', '')}｜主色 #{t.get('colors', {}).get('main', '')}")
    return files


def theme_save(proj, name):
    import paths
    style = paths.of(proj).排版 / "样式.json"
    if not style.exists():
        raise SystemExit(f"{style} 不存在")
    data = json.loads(style.read_text(encoding="utf-8"))
    theme = data.get("theme", data)
    cols = theme.get("colors", {})
    if len({v.upper() for v in cols.values() if isinstance(v, str)}) > 12:
        print("⚠️ 颜色值太多，检查是不是 ≤5 种颜色（同色的深浅不算）")
    THEMES.mkdir(parents=True, exist_ok=True)
    out = THEMES / f"{name}.json"
    if out.exists():
        raise SystemExit(f"主题「{name}」已存在")
    out.write_text(json.dumps(theme, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已存主题 -> {out}；补上「风格」「适合」两个字段")
    return out
