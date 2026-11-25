"""列出内容稿用到的插图及来源：python3 -I 排版/查图.py"""
import re
from pathlib import Path
H = Path(__file__).resolve().parent
t = Path("内容稿.md").read_text(encoding="utf-8")
names = []
for m in re.finditer(r"^\s*\[图[右组]\s+(.*)\]\s*$", t, re.M):
    inner = re.sub(r"(宽|高)=[\d.]+!?", "", m.group(1))
    names += [x.strip() for x in inner.split(";") if x.strip()]
names += [m.group(1) for m in re.finditer(r"\[\[图\s+(.+?)\s+高=", t)]
have = {}
for d in ["插图", "课本插图_全部"]:
    for f in sorted((H / d).glob("*.png")):
        have.setdefault(f.stem.split("_", 2)[-1], f"{d}/{f.name}")
seen = []
for n in names:
    if n in seen: continue
    seen.append(n)
    print(("✓ " + have[n]) if n in have else ("✗ " + n))
