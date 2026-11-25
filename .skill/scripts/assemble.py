"""把每章的 讲解.md + pick.py 取出的题和答案拼成 内容稿.md。
用法：python3 assemble.py <项目> <章键>... [--answers chapter|end]
  章键对应 选题配置/<键>.json（选题）和 选题配置/<键>_讲解.md（讲解，第一个章的文件头就是全书文件头）
  --answers chapter（默认）：每章顺序是 讲解 → 练习 → 本章答案（[答案] 标记，成品里另起一页）
  --answers end：答案集中到书末 # 参考答案
标点只改格式不改字（见 tidy）。
"""
import argparse, json, re, subprocess, sys
from pathlib import Path

PICK = Path(__file__).with_name("pick.py")


def tidy(t):
    """素材的标点格式（只改标点，不改字）：中文引号两侧空格去掉；中文后的英文问号、冒号、逗号换全角；
    英文逗号、分号后补空格；单词后句号紧跟大写单词时补空格（a.m. 之类不动）；选项字母统一「A．」。"""
    t = re.sub(r"(?<=[一-鿿，。：；])[ \t]+(?=[“”])", "", t)
    t = re.sub(r"(?<=[“”])[ \t]+(?=[一-鿿])", "", t)
    t = re.sub(r"(?<=[一-鿿])\?", "？", t)
    t = re.sub(r"(?<=[一-鿿]):", "：", t)
    t = re.sub(r"(?<=[一-鿿]),(?=[一-鿿])", "，", t)
    t = re.sub(r"(?<=[a-z]),(?=[A-Za-z])", ", ", t)
    t = re.sub(r"(?<=[a-z]);(?=[A-Za-z])", "; ", t)
    t = re.sub(r"(?<=[a-z]{2})\.(?=[A-Z][a-z])", ". ", t)
    t = re.sub(r"\((?: )+(?=[a-z])", "(", t)                     # 「( talk)」→「(talk)」
    t = re.sub(r"(?<=[a-z]\))(?=[a-z])", " ", t)                  # 「(eat)the」→「(eat) the」
    t = re.sub(r"(?<=[a-z])\((?=[一-鿿])", " (", t)              # 「countryside(社会主义新农村)」
    t = re.sub(r"(?<=[A-Za-z])，\s*(?=[A-Za-z＿])", ", ", t)      # 英文句里的中文逗号
    t = "\n".join(re.sub(r"(?<=[”。])\s{2,}(?=[“])", "", l) for l in t.split("\n"))   # 解析里 PDF 断行留下的大段空格
    t = "\n".join(re.sub(r"(^(?:\d+\.)?|｜)([A-D])\.\s*", r"\1\2．", l) if re.match(r"^(\d+\.)?[A-D]\.", l) else l
                  for l in t.split("\n"))
    return t


def tidy_answers(t):
    """解析里不用箭头串顺序（references/文风规范.md 机器腔），换成顿号。"""
    def fix(l):
        l = re.sub(r"\s*→\s*", "、", l)
        l = re.sub(r"(?<=[一-鿿，。；：、])[ \t]+(?=[一-鿿，。；：、“（])", "", l)   # PDF 转来的断行空格「动 词原形」
        l = re.sub(r"(?<=[一-鿿”）)])\s*;\s*", "；", l)   # 中文解析里的英文分号、逗号、叹号换全角（英文例句里的不动）
        l = re.sub(r"(?<=[一-鿿”）)])\s*,\s*(?=[一-鿿“A-Za-z])", "，", l)
        l = re.sub(r"(?<=[一-鿿])!", "！", l)
        return re.sub(r"“\s+", "“", re.sub(r"\s+”", "”", l))
    return "\n".join(fix(l) if l.startswith("解析：") else l for l in t.split("\n"))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("project")
    ap.add_argument("keys", nargs="+")
    ap.add_argument("--answers", default="chapter", choices=["chapter", "end"])
    a = ap.parse_args()
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import paths
    P = paths.of(a.project)
    proj = P.root; cfg = P.选题配置
    parts, answers = [], []
    for i, key in enumerate(a.keys):
        if not (cfg / f"{key}.json").exists():   # 只有讲解、没有练习的章（如书末附录）
            parts.append(re.sub(r"^---.*?---\s*", "", (cfg / f"{key}_讲解.md").read_text(encoding="utf-8").rstrip(), flags=re.S) if i else
                         (cfg / f"{key}_讲解.md").read_text(encoding="utf-8").rstrip())
            continue
        out = json.loads(subprocess.run([sys.executable, "-X", "utf8", "-I", str(PICK), str(P.素材), str(cfg / f"{key}.json")],
                                        capture_output=True, text=True, encoding="utf-8", check=True).stdout)
        explain = (cfg / f"{key}_讲解.md").read_text(encoding="utf-8").rstrip()
        if i:  # 只有第一个章保留文件头
            explain = re.sub(r"^---.*?---\s*", "", explain, flags=re.S)
        unit = re.search(r"^# (.+)$", explain, re.M).group(1)
        ex = re.sub(r"^(### .+基础练习)", r"[分页]\n\1", tidy(out["ex"]), count=1, flags=re.M)
        ans = tidy_answers(tidy(out["ans"]))
        if a.answers == "chapter":
            short = re.search(r"^### (.+?) \S*练习", ex, re.M).group(1)   # Starter / Unit 1 …
            parts.append(explain + "\n\n" + ex + f"\n\n[答案] {short} 参考答案\n" + ans)
        else:
            parts.append(explain + "\n\n" + ex)
            answers.append(f"## {unit}\n" + ans)
    text = "\n\n".join(parts) + "\n"
    if answers:
        text += "\n# 参考答案\n" + "\n".join(answers) + "\n"
    P.内容稿.write_text(text, encoding="utf-8")
    print(f"内容稿.md 已生成（{len(a.keys)} 章，答案{'在每章后' if a.answers == 'chapter' else '在书末'}）")


if __name__ == "__main__":
    main()
