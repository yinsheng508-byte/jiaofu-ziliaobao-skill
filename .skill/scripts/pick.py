"""从素材 txt 里按题号原样取题、答案、解析，转成内容稿标记。只做格式转换，不改字。
用法：python3 pick.py <素材目录> <配置json>  → 打印 {"ex": 练习区文本, "ans": 答案区文本}
配置写法、kind、fix / ansgroup / join 等选项和取舍规则见 <skill>/templates/英语语法手册/取题流程.md；一般不直接调，用 run.py assemble。
"""
import json, re, sys
from pathlib import Path

SRC = Path(sys.argv[1])
CFG = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))

QN = re.compile(r"^\s*(\d{1,3})\s*[．.](?!\d)\s*")


def blank(s):
    s = re.sub(r"_{2,}", lambda m: "＿" * max(4, min(len(m.group()) // 2, 8)), s)
    return re.sub(r"＿{9,}", "＿" * 8, s)


def split_opts(s):
    # A．IB．HeC．YouD．She → A．I｜B．He｜C．You｜D．She
    s = re.sub(r"(?<=\S)\s*(?=[B-D][．.]\s?)", "｜", s)
    return re.sub(r"\s*｜\s*", "｜", s)


PAGE_NO = re.compile(r"^\s*\d+\s*/\s*\d+\s*$")   # PDF 转来的页码行「9 / 24」


def load(sid):
    f = next(SRC.glob(f"{sid}_*.txt"))
    return ["" if PAGE_NO.match(l) else l.rstrip() for l in f.read_text(encoding="utf-8").split("\n")]  # 清空不删行，行号不变


def choice_items(lines, start, end, lead=False):
    """单选/填空：返回 {题号: (题干行们, 答案, 解析行们)}"""
    items, cur = {}, None
    for l in lines[start - 1:end]:
        s = l.strip()
        if not s:
            continue
        m = QN.match(s)
        if m and (cur is None or cur["stage"] != "q" or cur["ans"] is not None):
            if cur is None or cur["ans"] is not None:
                cur = {"q": [l.lstrip()], "ans": None, "why": [], "stage": "q"}
                items[int(m.group(1))] = cur
                continue
        if cur is None:
            continue
        if s.startswith("【答案】"):
            if cur["ans"] is not None:   # 题干丢了、只剩答案的孤立题：丢弃，直到下一个题号
                cur = None
                continue
            cur["ans"], cur["stage"] = s[4:].strip(), "a"
        elif s.startswith(("【详解】", "【解析】")):
            cur["why"].append(s[4:].strip()); cur["stage"] = "w"
        elif cur["stage"] == "q":
            s2 = l.strip(" \t")
            if lead and re.match(r"^[ \t]{4,}\S", l) and not re.match(r"^\s*[A-D][．.]", l):
                ws = len(l) - len(l.lstrip(" \t"))   # 行首的空抽字后只剩空格（完成句子「      for you?」），空长随原文
                s2 = "＿" * max(4, min(12, ws // 3)) + " " + s2
            cur["q"].append(s2)
        elif cur["stage"] in ("a", "w"):
            cur["why"].append(s)
    return items


def fmt_q(n, qlines, tag, hint_blank=False):
    # 分成两行写的选项（A．…｜B．… / C．… D．…）先并成一行
    merged = []
    for s in qlines:
        if merged and re.match(r"^\s*[B-D][．.]", s) and re.match(r"^\s*A[．.]", merged[-1]):
            merged[-1] = merged[-1].rstrip() + "  " + s.strip()
        else:
            merged.append(s)
    qlines = merged
    out = []
    def dash(s):   # 对话的「-」「--」换成破折号；后面跟着大段空白的是空，空白留着
        s = s.strip() if re.match(r"^\s*-", s) else s
        s = re.sub(r"^(\s*\d{1,3}\s*[.．]\s*)-{1,3}", r"\1—", s)
        s = re.sub(r"(^|(?<=\s))-{1,3}(?= {2,})", "—", s)
        s = re.sub(r"(^|(?<=\s))-{1,3} ?(?=[A-Z(（_＿])", "—", s)
        s = re.sub(r"(?<=[.?!])\s*-{1,3}\s*(?=[A-Z])", "　—", s)   # 「….- I am not sure.」一行里的第二句
        return re.sub(r"—-+ ?", "—", s)
    qlines = [dash(s) for s in qlines]
    for i, s in enumerate(qlines):
        if i and re.fullmatch(r"[\s\u3000]{30,}[.?!？。]?\s*", s):   # 连词成句：一整行空白是作答线
            out.append("[横线]")
            continue
        if re.match(r"^\s*A[．.]", s):           # 选项行：只分栏，里面的空格不当成空
            out.append(split_opts(re.sub(r"\s{2,}", "  ", s.strip())))
            continue
        s = re.sub(r"(?<=_)\s{2,}|\s{2,}(?=_)", " ", s)   # 「______     (eat)」：下划线旁的空格不是另一个空
        if i == 0:
            s = re.sub(r"^\s*\d{1,3}\s*[．.]", "", s)
        s = re.sub(r"\s{4,}(?=—)", "　　", s)   # 对话两句之间的大空格是间隔，不是空
        s = re.sub(r"　[ \t]+　", " ＿＿＿＿ ", s)   # 「全角+空格+全角」是一个空，两个挨着就是两个空
        s = re.sub(r"(^|(?<=\S))\s{4,}",            # 其余大段空白：一个空，长度随原文空白
                   lambda m: (" " if m.start() else "") + "＿" * max(4, min(10, len(m.group()) // 3)) + " ", s)
        if hint_blank and i == 0 and "＿" not in s and "__" not in s:   # 原文的空是下划线格式，抽字后丢了：补在提示词前
            s = re.sub(r"\s*(?=[(（][^)）]*[)）])", " ＿＿＿＿ ", s, count=1)
        s = re.sub(r"(?<=\S) {2,}", " ", re.sub(r"(?<=＿) (?=[.?,!])", "", blank(s))).strip()
        s = re.sub(r"(?<=[A-Za-z’'])(?=＿)|(?<=＿)(?=[A-Za-z(（])", " ", s)   # 空贴着单词时两边留空格
        if i == 0:
            s = f"{n}.{s}{tag}"
        elif re.match(r"^A[．.]", s):
            s = split_opts(s)
        out.append(s)
    return out


def endkey_map(L, sec):
    """卷末答案 → {题号: 答案}。「1-5.DCACB」按字母拆；其余按「数字.」切开，答案里的「。……」解析部分去掉。"""
    t = " ".join(l.strip() for l in L[sec["ans_from"] - 1:sec["ans_to"]] if l.strip() not in ("[表格]", "[/表格]"))
    t = re.sub(r"^[一二三四五六七八九十]+、[^\d\s]*", "", t.strip())
    t = t.replace("|", " ")
    amap = {}
    for a, b, letters in re.findall(r"(\d+)\s*-\s*(\d+)\s*[.．]?\s*([A-D]{2,})(?![a-z])", t):
        for i, c in enumerate(letters):
            amap[int(a) + i] = c
    t = re.sub(r"\d+\s*-\s*\d+\s*[.．]?\s*[A-D]{2,}(?![a-z])", " ", t)
    parts = re.split(r"(?<![\d.])(\d{1,2})\s*[.．](?!\d)", " " + t)   # 「……用 “visit” 。2. is」：题号前不一定有空格
    for i in range(1, len(parts) - 1, 2):
        amap.setdefault(int(parts[i]), re.sub(r"\s*[。.]?\s*$", "", parts[i + 1].strip()))
    amap.update({int(k): v for k, v in sec.get("ansfix", {}).items()})
    return amap


def fmt_why(why):
    why = [re.sub(r"【(分析|详解|点睛)】", "", w) for w in why if w]
    return "解析：" + "".join(why) if why else None


ex, ans = [], []
for grp in CFG["groups"]:
    if grp.get("break"):   # 这一组另起一页（回收孤页时用）
        ex.append("[分页]")
    ex.append(f"### {grp['title']}")
    ans.append(f"### {grp['title']}")
    n = 0
    for sec in grp["sections"]:
        if sec["head"]:   # 同一大题接着取另一份素材时 head 写空
            ex.append(sec["head"])
        ans.append(sec["head"])
        L = load(sec["sid"])
        if sec["kind"] == "items":
            items = choice_items(L, sec["from"], sec["to"], sec.get("lead_blank", False))
            for k in sec["pick"]:
                it = items[k]
                for a_, b_ in sec.get("fix", {}).get(str(k), {}).items():   # 素材笔误，覆盖表里有记录
                    it["q"] = [q.replace(a_, b_) for q in it["q"]]
                merged = []   # PDF 抽字把一句题干断成两行：上一行以字母或逗号结尾、下一行不是选项或对话，就接回去
                for q in it:
                    if merged and re.search(r"[A-Za-z,，-]$", merged[-1]) and not re.match(r"^\s*(—|-|[A-D][．.]|[一-鿿])", q):
                        merged[-1] = merged[-1].rstrip("-") + ("" if merged[-1].endswith("-") else " ") + q.strip()
                    else:
                        merged.append(q)
                it = merged
                if k in sec.get("join", []):   # 很短的一问一答排成一行（同节已有题这样排），省版面
                    qs = [q for q in it["q"] if not re.match(r"^\s*[A-D][．.]", q)]
                    rest = [q for q in it["q"] if re.match(r"^\s*[A-D][．.]", q)]
                    it["q"] = ["      ".join(q.strip() for q in qs)] + rest
                n += 1
                ex += fmt_q(n, it["q"], f"〔{sec['sid']} {sec['label']}{k}〕", sec.get("hint_blank", False))
                if sec.get("lines") and ex[-1] != "[横线]":   # 原题已有作答线就不再加
                    ex.append(f"[横线{sec['lines']}]" if sec["lines"] > 1 else "[横线]")
                a_txt = sec.get("ansgroup", {}).get(str(k)) or re.sub(r'\s{3,}', '；', it['ans'].strip())   # 一条长空填几个词时，答案按空分组
                ans.append(f"{n}.参考答案：{a_txt}〔{sec['sid']} {sec['label']}{k}〕")
                w = fmt_why(it["why"])
                if w:
                    ans.append(w)
        elif sec["kind"] == "batch":  # 题目在前、一行【答案】1.x 2.y 集中给出（没有解析）
            qs, ansline, whys, box, inbox, inwhy = {}, "", {}, [], False, False
            for l in L[sec["from"] - 1:sec["to"]]:
                if l.strip() == "[表格]" and not qs:      # 选词填空的词框
                    inbox = True; continue
                if inbox:
                    if l.strip() == "[/表格]":
                        inbox = False
                    else:
                        box += [w for w in re.split(r"\s{2,}|\|", l.strip()) if w]
                    continue
                if l.strip().startswith(("【解析】", "【详解】")):   # 集中解析：按题号分给各题
                    inwhy, l = True, l.strip()[4:]
                if inwhy:
                    m = QN.match(l.strip())
                    if m:
                        whys[int(m.group(1))] = QN.sub("", l.strip())
                    elif whys and l.strip():
                        whys[max(whys)] += l.strip()
                    continue
                if l.strip().startswith("【答案】"):
                    ansline += " " + l.strip()[4:]
                elif ansline and l.strip():
                    ansline += " " + l.strip()      # 答案折到下一行
                elif QN.match(l.strip()):
                    qs[int(QN.match(l.strip()).group(1))] = [l.lstrip()]
            parts = re.split(r"(?:^|\s)(\d{1,2})\s*[.．]", ansline)
            sep = r"\s+" if sec.get("ans_sep") == "single" else r"\s{2,}"   # 两空之间用几个空格隔开，看素材写法
            amap = {int(parts[i]): re.sub(r"\s+" if int(parts[i]) in sec.get("split_items", []) else sep, "；",
                                          parts[i + 1].strip()) for i in range(1, len(parts) - 1, 2)}
            if box:
                ex += ["[表 无表头]", " | ".join(box), "[/表]"]
            for k in sec["pick"]:
                n += 1
                ex += fmt_q(n, qs[k], f"〔{sec['sid']} {sec['label']}{k}〕")
                ans.append(f"{n}.参考答案：{amap[k]}〔{sec['sid']} {sec['label']}{k}〕")
                if whys.get(k):
                    ans.append("解析：" + whys[k])
        elif sec["kind"] == "passage":
            block = [l.strip() for l in L[sec["from"] - 1:sec["to"]] if l.strip()]
            body, opts, ansline, whys = [], [], None, {}
            stage = "body"
            for s in block:
                if s.startswith("【答案】"):
                    ansline, stage = s[4:], "ans"; continue
                if s.startswith(("【解析】", "【导语】")):
                    stage = "why"; whys["_"] = s[4:] if s.startswith("【解析】") else s; continue
                if stage == "ans":   # 答案另起一行（中考专项训练「【答案】」单独一行）
                    ansline += " " + s; continue
                m = QN.match(s)
                if stage == "body" and m and re.search(r"A[．.]", s):
                    opts.append((int(m.group(1)), s)); continue
                if stage == "body":
                    body.append(s)
                elif stage == "why" and m:
                    whys[int(m.group(1))] = QN.sub("", s)
                elif stage == "why":
                    last = max(k for k in whys if k != "_") if len(whys) > 1 else "_"
                    whys[last] += s
            fix = sec.get("fix", {})
            for s in body:
                for a, b in fix.items():
                    s = s.replace(a, b)
                s = blank(s)
                nb = "\u00a0" * 3   # 不换行空格：编号空不会在行尾断开
                s = re.sub(r"\s{2,}(\d{1,2})\s{2,}", lambda m: f" __{nb}{m.group(1)}{nb}__ ", s)
                s = re.sub(r"^(\d{1,2})\s{3,}", lambda m: f"__{nb}{m.group(1)}{nb}__ ", s)   # 段首的空（行首空白抽字时已去掉）
                ex.append(("| " + s) if not s.startswith(("（", "阅读")) else s)
            for k, s in opts:
                ex.append(f"{k}.{split_opts(QN.sub('', s))}〔{sec['sid']} {sec['label']}{k}〕")
            if not opts:   # 短文填空（没有选项）：空号从短文里找，答案按「n．」切开；出处标在短文末尾
                ex[-1] += f"〔{sec['sid']} {sec['label'].strip()}〕"
                gaps = [int(g) for g in re.findall(r"__\u00a0+(\d{1,2})\u00a0+__", "\n".join(ex))]
                parts = re.split(r"(?:^|\s)(\d{1,2})\s*[.．](?!\d)", " " + ansline)
                g_map = {int(parts[i]): parts[i + 1].strip() for i in range(1, len(parts) - 1, 2)}
                g_map.update({int(k_): v_ for k_, v_ in sec.get("ansfix", {}).items()})
                if whys.get("_"):
                    ans.append(whys["_"])
                for k in gaps:
                    ans.append(f"{k}.参考答案：{g_map[k]}〔{sec['sid']} {sec['label']}{k}〕")
                    if k in whys:
                        ans.append("解析：" + whys[k])
                n += len(gaps)
                continue
            a_map = dict(re.findall(r"(\d+)\s*[.．]\s*([A-Za-z’'/ ]+?)(?=\s+\d+\s*[.．]|$)", ansline))
            for a, b in sec.get("ansfix", {}).items():   # 素材答案行题号错（如 20 应为 10）
                a_map[b] = a_map.pop(a)
            if whys.get("_"):
                ans.append(whys["_"])
            for k, _ in opts:
                ans.append(f"{k}.参考答案：{a_map[str(k)].strip()}〔{sec['sid']} {sec['label']}{k}〕")
                if k in whys:
                    ans.append("解析：" + whys[k])
            n += len(opts)
        elif sec["kind"] == "endkey":   # 题目只有题号，答案集中在卷末「一、1. xx 2. xx」或「1-5.DCACB」（八上巩固练习）
            qs, cur = {}, None
            for l in L[sec["from"] - 1:sec["to"]]:
                l = re.sub(r"^\s*[(（][\s\u3000]*[)）]\s*(?=\d)", "", l)   # 「( )1. …」单选题号前的括号（九年级专练 PDF）
                m = QN.match(l.strip())
                if m and int(m.group(1)) in qs or re.match(r"^\d{1,2} [A-Z]", l.strip()):   # 题号重复（「1.Your bike」应为 21）、题号缺点（「8 The old」）：这题连同选项丢掉
                    cur = []
                elif m:
                    cur = qs[int(m.group(1))] = [l.strip()]
                elif cur is not None and l.strip() and l.strip() not in ("-", "－", "—") \
                        and not re.fullmatch(r"\d{1,3}|=+ 第 \d+ 页 =+", l.strip()):   # 翻译题下面的「-」占位行、PDF 页码行
                    lead = sec.get("lead_blank") and re.match(r"^[ \t]{4,}\S", l) and not re.match(r"^\s*[A-D][．.]", l)
                    cur.append(("＿＿＿＿ " if lead else "") + l.strip())   # 作答行行首的空抽字后只剩空格（docx 专题）
            amap = endkey_map(L, sec)
            for k in sec["pick"]:
                it = qs[k]
                for a_, b_ in sec.get("fix", {}).get(str(k), {}).items():
                    it = [q.replace(a_, b_) for q in it]
                merged = []   # PDF 抽字把一句题干断成两行：上一行以字母或逗号结尾、下一行不是选项或对话，就接回去
                for q in it:
                    if merged and re.search(r"[A-Za-z,，-]$", merged[-1]) and not re.match(r"^\s*(—|-|[A-D][．.]|[一-鿿])", q):
                        merged[-1] = merged[-1].rstrip("-") + ("" if merged[-1].endswith("-") else " ") + q.strip()
                    else:
                        merged.append(q)
                it = merged
                if k in sec.get("join", []):   # 两行的一问一答并成一行（回收孤页时用）
                    stem = [q for q in it if not re.match(r"^\s*[A-D][．.]", q)]
                    it = ["　　".join(q.strip() for q in stem)] + [q for q in it if q not in stem]
                n += 1
                ex += fmt_q(n, it, f"〔{sec['sid']} {sec['label']}{k}〕", sec.get("hint_blank", False))
                if sec.get("lines"):
                    ex.append(f"[横线{sec['lines']}]" if sec["lines"] > 1 else "[横线]")
                a_, why_ = amap[k], ""
                if sec.get("ans_letter") and re.match(r"^[A-D](?![a-z’'])|^[A-D](?=[a-z])", a_):   # 「2.D 句意:……」：字母是答案，后面是解析
                    a_, why_ = a_[0], a_[1:].strip(" :：")
                elif sec.get("ans_cut") and sec["ans_cut"] in a_:   # 「1.cut 。祈使句用动词原形。」：前面是答案，后面是解析
                    a_, why_ = (x.strip() for x in a_.split(sec["ans_cut"], 1))
                if sec.get("lines") and re.search(r"[A-Za-z]$", a_) and re.match(r"[A-Z]", a_) and len(a_.split()) >= 4:   # 翻译、完成句子的整句答案补回句号
                    a_ += "."
                ans.append(f"{n}.参考答案：{sec.get('ansgroup', {}).get(str(k), a_)}〔{sec['sid']} {sec['label']}{k}〕")
                if why_:
                    ans.append("解析：" + why_ + ("" if re.search(r"[。！？!?”」）)]$", why_) else "。"))
        elif sec["kind"] == "cloze":   # 课文语法填空：短文里「1.____(feel)」编号空，答案在文末一行一个（八上 S26/S27）
            nb = " " * 3
            k0 = n
            for l in L[sec["from"] - 1:sec["to"]]:
                s = l.strip()
                if not s:
                    continue
                for a, b in sec.get("fix", {}).items():
                    s = s.replace(a, b)
                s = re.sub(r"(\d{1,2})\s*[.．]\s*_{3,}\s*",   # 空号接着本组题号往下排，括号提示词前留空格
                           lambda m: f" __{nb}{k0 + int(m.group(1))}{nb}__ ", s)
                s = re.sub(r"\s{2,}", " ", s).strip()
                ex.append("| " + s)
            keys = [l.strip() for l in L[sec["ans_from"] - 1:sec["ans_to"]] if l.strip()]
            for i, a_ in enumerate(keys, 1):
                a_ = sec.get("ansfix", {}).get(str(i), a_)
                ans.append(f"{k0 + i}.参考答案：{a_}〔{sec['sid']} {sec['label']}{i}〕")
            n += len(keys)
print(json.dumps({"ex": "\n".join(ex), "ans": "\n".join(ans)}, ensure_ascii=False))
