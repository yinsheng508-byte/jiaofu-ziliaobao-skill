"""答案复核：按 选题配置/*.json 回到素材原文，用独立的扫描方式找每道题的【答案】，和内容稿答案区逐条比对。
用法：python3 verify.py <项目目录> <单元键>...   输出不符的条目；全部一致时打印「全部一致」。
"""
import json, re, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import paths  # noqa: E402
P = paths.of(sys.argv[1])
proj = P.root
content = P.内容稿.read_text(encoding="utf-8")
mine = {}
for a, tag in re.findall(r"参考答案：(.+?)〔(S\d{2,3} [^〕]+)〕", content):
    mine[tag] = re.sub(r"[；;,，\s]", "", a).rstrip("。.")


def src_lines(sid):
    return next(P.素材.glob(f"{sid}_*.txt")).read_text(encoding="utf-8").split("\n")


bad = total = 0
for key in sys.argv[2:]:
    cfg = json.loads((P.选题配置 / f"{key}.json").read_text(encoding="utf-8"))
    for g in cfg["groups"]:
        for sec in g["sections"]:
            L = src_lines(sec["sid"])[sec["from"] - 1:sec["to"]]
            if sec["kind"] == "passage":
                text = " ".join(L)
                m = re.search(r"【答案】(.+?)(?=【解析】|【导语】|$)", text)
                amap = dict(re.findall(r"(\d+)\s*[.．]\s*([A-D])\b", m.group(1)))
                if not re.search(r"(?m)^\s*\d+\s*[.．].*A[．.]", "\n".join(L)):   # 短文填空：没有选项，答案是词
                    parts = re.split(r"(?:^|\s)(\d{1,2})\s*[.．](?!\d)", " " + m.group(1))
                    amap = {parts[i]: parts[i + 1].strip() for i in range(1, len(parts) - 1, 2)}
                amap.update({v: amap.pop(k) for k, v in sec.get("ansfix", {}).items() if k in amap})
                picks = sorted(int(k) for k in amap)
            elif sec["kind"] == "endkey":   # 卷末答案：按「k.」到「k+1.」截取（和 pick.py 的切法不同，互相核对）
                t = " " + " ".join(l.strip() for l in src_lines(sec["sid"])[sec["ans_from"] - 1:sec["ans_to"]]
                                   if l.strip() not in ("[表格]", "[/表格]")).replace("|", " ")
                amap, picks = {}, sec["pick"]
                for a, b, letters in re.findall(r"(\d+)\s*-\s*(\d+)\s*[.．]?\s*([A-D]{2,})(?![a-z])", t):
                    for i, c in enumerate(letters):
                        amap[str(int(a) + i)] = c
                for n in picks:
                    if str(n) not in amap:
                        m = re.search(rf"(?<!\d){n}\s*[.．](?!\d)\s*(.+?)(?=(?<!\d){n + 1}\s*[.．](?!\d)|$)", t)
                        amap[str(n)] = m.group(1).strip().rstrip("。.").strip() if m else "（素材里没找到）"
                        if sec.get("ans_cut"):
                            amap[str(n)] = amap[str(n)].split(sec["ans_cut"])[0].strip()
                        if sec.get("ans_letter") and re.match(r"^[A-D](?![a-z’'])|^[A-D](?=[a-z])", amap[str(n)]):
                            amap[str(n)] = amap[str(n)][0]
                amap.update(sec.get("ansfix", {}))
                if sec.get("ans_letter"):
                    amap = {k: (a[0] if re.match(r"^[A-D](?![a-z’'])|^[A-D](?=[a-z])", a) else a) for k, a in amap.items()}
            elif sec["kind"] == "cloze":   # 答案一行一个，按顺序就是空号
                keys = [l.strip() for l in src_lines(sec["sid"])[sec["ans_from"] - 1:sec["ans_to"]] if l.strip()]
                amap = {str(i): a for i, a in enumerate(keys, 1)}
                amap.update(sec.get("ansfix", {}))
                picks = list(range(1, len(keys) + 1))
            elif sec["kind"] == "batch":
                line = " ".join(l for l in L if "【答案】" in l).replace("【答案】", " ")
                parts = re.split(r"(?:^|\s)(\d{1,2})\s*[.．]", line)
                amap = {parts[i]: parts[i + 1] for i in range(1, len(parts) - 1, 2)}
                picks = sec["pick"]
            else:
                amap, picks = {}, sec["pick"]
                for n in picks:
                    for i, l in enumerate(L):
                        if re.match(rf"^\s*{n}\s*[．.](?!\d)", l):
                            for j in range(i, min(i + 25, len(L))):
                                if L[j].strip().startswith("【答案】"):
                                    amap[str(n)] = L[j].strip()[4:]
                                    break
                            break
            for n in picks:
                total += 1
                tag = f"{sec['sid']} {sec['label']}{n}"
                src = re.sub(r"[；;,，\s]", "", amap.get(str(n), "（素材里没找到）")).rstrip("。.")
                if mine.get(tag) != src:
                    bad += 1
                    print(f"不符 {tag}：内容稿={mine.get(tag)}  素材={src}")
print(f"核对 {total} 条" + ("，全部一致" if not bad else f"，不符 {bad} 条"))
