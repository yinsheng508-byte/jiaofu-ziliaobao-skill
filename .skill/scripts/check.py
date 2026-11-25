"""内容稿检查：出处覆盖、【缺】、机器味、题目和答案对不对得上。生成《检查报告.md》。

  python3 check.py <内容稿.md> [--out 检查报告.md]

退出码：有 error 级问题时为 1。
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_docx import SRC_RE, parse  # noqa: E402
import author  # noqa: E402

SKILL = Path(__file__).resolve().parent.parent
LINT = json.loads((SKILL / "config" / "检查词表.json").read_text(encoding="utf-8"))["rules"]


def groups(blocks):
    """切成需要标出处的内容块：每个考点讲解、每个例题、每篇练习、每篇答案。"""
    out, cur = [], None
    for b in blocks:
        k = b["kind"]
        start = None
        if k == "kd":
            start = ("考点讲解", b["text"])
        elif k == "li":
            start = ("例题", b["text"][:24])
        elif k == "ex_title":
            start = ("练习", b["text"])
        elif k == "ans_title":
            start = ("答案", b["text"])
        elif k in ("unit", "answer_start", "ans_unit"):
            cur = None
            continue
        if start:
            cur = {"type": start[0], "title": SRC_RE.sub("", start[1]).strip(), "blocks": [b]}
            out.append(cur)
        elif cur:
            cur["blocks"].append(b)
    return out


def norm_title(t):
    t = SRC_RE.sub("", t)
    t = re.sub(r"^[一二三四五六七八九十\d]+[、.．]\s*", "", t)
    return re.sub(r"\s+", "", t)


def count_questions(blocks):
    nums = {re.match(r"^(\d+)[.．]", b["text"]).group(1) for b in blocks
            if b["kind"] == "q" and re.match(r"^\d+[.．]", b["text"])}
    # 短文里的编号空（课文语法填空「__ 24 __」）也是一道题；语法选择的空另有选项行，按题号去重
    for b in blocks:
        if b["kind"] == "ex_para":
            nums |= set(re.findall(r"__[\s ]*(\d+)[\s ]*__", b["text"]))
    return len(nums)


def count_answers(blocks):
    n = 0
    for b in blocks:
        if b["kind"] == "ans_q" and re.match(r"^\d+[.．]", b["text"]):
            n += 1
        elif b["kind"] in ("ans_text", "ans_ans") and re.match(r"^\d+[.．]", b["text"]):
            n += 1
    return n


MARK = re.compile(r"^\s*(?:@[^\s：:]+[：:]?|\[[^\]]*\])\s*")   # 行首的模板标记（@考点 [框 …]）不算正文


def lint(blocks, subject="", generic=False):
    hits = []
    for b in blocks:
        text = b["text"]
        if not text or b["kind"] in ("passage", "ex_para", "source"):
            continue  # 课文和练习短文是素材原文，不按机器味改
        for r in LINT:
            zones = r["zones"]
            if "all" not in zones and b["zone"] not in zones:
                continue
            if subject and subject in r.get("skip_subjects", []):
                continue
            if b["kind"] == "q" and r["name"] == "套话收尾":
                continue
            if generic and r.get("classic_only"):
                continue
            if re.search(r["pattern"], MARK.sub("", text)):
                hits.append((r["level"], r["name"], text[:60], r["hint"]))
    return hits


ANY_SRC = re.compile(r"〔[^〕]*〕|（(?:S\d\d|课本|教材)[^）]*）")


def generic_groups(content):
    """项目自带排版工具（模板格式）时：按 ### 知识点（没有就按 ## 节）切块，每块要有出处。"""
    text = re.sub(r"<!--.*?-->", "", Path(content).read_text(encoding="utf-8"), flags=re.S)
    lines = text.splitlines()
    level = "### " if any(l.startswith("### ") for l in lines) else "## "
    out, cur = [], None
    for l in lines:
        if l.startswith(level):
            cur = {"type": "知识点", "title": ANY_SRC.sub("", l[len(level):]).strip(), "text": l}
            out.append(cur)
        elif l.startswith(("# ", "## ")) and level == "### ":
            cur = None
        elif cur:
            cur["text"] += "\n" + l
    return out


def check(content):
    meta, blocks = parse(content)
    if (Path(content).parent / "排版" / "build.py").exists():     # 模板格式：通用检查，不按经典标记核题答
        gs = generic_groups(content)
        no_src = [g for g in gs if not ANY_SRC.search(g["text"])]
        gaps = [l for l in Path(content).read_text(encoding="utf-8").splitlines() if "【缺】" in l]
        hits = lint(blocks, meta.get("学科", ""), generic=True)
        stats = {"知识块": len(gs)}
        return stats, no_src, gaps, hits, []
    gs = groups(blocks)
    no_src = [g for g in gs if not any(SRC_RE.search(b["text"]) for b in g["blocks"])]
    gaps = [b["text"] for b in blocks if "【缺】" in b["text"]]
    hits = lint(blocks, meta.get("学科", ""))
    ex = {norm_title(g["title"]): count_questions(g["blocks"]) for g in gs if g["type"] == "练习"}
    an = {norm_title(g["title"]): count_answers(g["blocks"]) for g in gs if g["type"] == "答案"}
    mismatch = []
    for t, n in ex.items():
        if t not in an:
            mismatch.append(f"「{t}」有 {n} 道题，答案区找不到这一篇")
        elif an[t] != n:
            mismatch.append(f"「{t}」有 {n} 道题，答案 {an[t]} 条")
    for t in an:
        if t not in ex:
            mismatch.append(f"答案区「{t}」在练习里找不到")
    stats = {
        "单元": sum(b["kind"] == "unit" for b in blocks),
        "考点": sum(b["kind"] == "kd" for b in blocks),
        "例题": sum(b["kind"] == "li" for b in blocks),
        "练习篇数": len(ex),
        "练习题数": sum(ex.values()),
        "答案条数": sum(an.values()),
    }
    return stats, no_src, gaps, hits, mismatch


def author_hits(content):
    """原作者信息：整份内容稿逐行查（课文、练习短文、文件头也查），出处标记〔Sxx pN〕不算。"""
    terms = author.project_terms(Path(content))
    out = []
    for i, line in enumerate(Path(content).read_text(encoding="utf-8").splitlines(), 1):
        for lv, word, ctx in author.scan(SRC_RE.sub("", line), terms):
            out.append((lv, i, word, ctx))
    return terms, out


def report(content, out):
    stats, no_src, gaps, hits, mismatch = check(content)
    terms, ah = author_hits(content)
    a_err = [h for h in ah if h[0] == "error"]
    errors = [h for h in hits if h[0] == "error"]
    L = [f"# 检查报告：{Path(content).name}", ""]
    L.append("统计：" + "　".join(f"{k} {v}" for k, v in stats.items()))
    L.append("")
    verdict = "❌ 还不能出成品" if (gaps or errors or no_src or mismatch or a_err) else "✅ 可以出成品"
    L += [f"**结论：{verdict}**", ""]
    L.append(f"## 一、【缺】（{len(gaps)} 处，交给用户补）")
    L += [f"- {g}" for g in gaps] or ["- 无"]
    L.append("")
    L.append(f"## 二、没标出处的内容块（{len(no_src)} 个）")
    L += [f"- {g['type']}：{g['title']}" for g in no_src] or ["- 无"]
    L.append("")
    L.append(f"## 三、题目和答案对不上（{len(mismatch)} 处）")
    L += [f"- {m}" for m in mismatch] or ["- 无"]
    L.append("")
    L.append(f"## 四、机器味（error {len(errors)} / warn {len(hits) - len(errors)}）")
    L += [f"- [{lv}] {name}：{txt} —— {hint}" for lv, name, txt, hint in hits] or ["- 无"]
    L.append("")
    L.append(f"## 五、原作者信息（error {len(a_err)} / warn {len(ah) - len(a_err)}；屏蔽词 {len(terms)} 个，来自 {author.TERMS_FILE}）")
    if not terms:
        L.append(f"- ⚠️ {author.TERMS_FILE} 的「屏蔽词」区是空的，只查了通用引流词；先把原作者名、机构名、品牌名填进去")
    L += [f"- [{lv}] 第 {i} 行：**{word}** —— {ctx}" for lv, i, word, ctx in ah] or ["- 无"]
    Path(out).write_text("\n".join(L) + "\n", encoding="utf-8")
    return verdict, stats, len(gaps), len(no_src), len(mismatch), len(errors) + len(a_err), len(hits) - len(errors) + len(ah) - len(a_err)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("content")
    ap.add_argument("--out")
    a = ap.parse_args()
    out = a.out or str(Path(a.content).with_name("检查报告.md"))
    verdict, stats, g, s, m, e, w = report(a.content, out)
    print(f"{verdict}｜缺 {g}｜无出处 {s}｜题答不符 {m}｜机器味和原作者信息 error {e} warn {w}")
    print("->", out)
    sys.exit(1 if verdict.startswith("❌") else 0)
