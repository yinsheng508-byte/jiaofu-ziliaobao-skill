#!/usr/bin/env python3
"""商品主图 / 详情图渲染：spec.json -> 每页 HTML -> Chrome / Edge 无头截图 PNG（1242x1656，3:4）。

用法:
  python3 run.py mainimg <spec.json> [输出目录]          渲染；输出目录默认是 spec.json 所在目录（01.png、02.png …，HTML 留在 html/）
  python3 run.py mainimg pages <资料.pdf> <目录> [页数]   把资料前几页转成 200dpi 图片（放主图展示位用）
  python3 run.py mainimg templates                       列出主图模板（assets/主图模板/）

spec.json 里写 "模板": "<名>" 套用主图模板（配色、底部条默认文字）；"theme" / "footer" 再覆盖。
店铺名（shop）必须由用户每次提供，脚本和模板里都不存。

行内标记（所有文案字段通用）:
  [[文字]]  蓝底白字标签      {{文字}}  红色强调      **文字**  加粗
"""
import html
import json
import logging
import math
import os
import re
import subprocess
import sys
from pathlib import Path

W, H = 1242, 1656
HERE = Path(__file__).resolve().parent
TEMPLATES = HERE.parent / "assets" / "主图模板"
sys.path.insert(0, str(HERE))
import export_pdf  # noqa: E402
CHROME = export_pdf.find_chrome() or ("chrome" if sys.platform == "win32" else "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")

DEFAULT_THEME = {
    "blue": "#1C4ED8",        # 外框主蓝（纯色，不做渐变）
    "blue_dark": "#12358F",   # 深蓝（文字）
    "blue_light": "#3D6FE0",  # 标签浅蓝
    "red": "#F01C1C",         # 大标题红（纯色，不描边不投影）
    "kicker": "#1D1D1F",      # 标题上方小字（深色）
    "panel": "#EAF0F5",       # 内板底色（纯平）
    "paper": "#F8F4EA",       # 资料封面纸色
    "ink": "#3A2A1E",         # 资料封面字色
}

SANS = '"PingFang SC","Hiragino Sans GB","Heiti SC","Microsoft YaHei","微软雅黑",sans-serif'
HEAVY = "'LTHeavy','PingFang SC','Microsoft YaHei','SimHei',sans-serif"   # 兰亭黑特黑：标题、店铺名、底部大字；Windows 用微软雅黑粗体
SERIF = '"FZDaBiaoSong-B06","方正大标宋_GBK","STSong","Songti SC","SimSun","宋体",serif'


# ---------- 工具 ----------
def rich(text):
    """转义后处理行内标记。"""
    if text is None:
        return ""
    t = html.escape(str(text)).replace("\n", "<br>")
    t = re.sub(r"\[\[(.+?)\]\]", r'<span class="chip">\1</span>', t)
    t = re.sub(r"\{\{(.+?)\}\}", r'<span class="red">\1</span>', t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    return t


def plain(text):
    return re.sub(r"\[\[|\]\]|\{\{|\}\}|\*\*", "", str(text or ""))


def text_em(text):
    """估算一行文字宽度（单位 em）：中文 1，ASCII 0.56。"""
    n = 0.0
    for ch in plain(text):
        n += 0.56 if ord(ch) < 128 else 1.0
    return max(n, 1.0)


def fit(lines, box_w, max_px, min_px=40):
    longest = max(text_em(l) for l in lines) if lines else 1
    return int(max(min_px, min(max_px, box_w / longest)))


def img_src(path, base, box=None):
    """返回图片 file:// 地址；box=(宽,高) 时检查原图是否够清晰（不足会被放大发糊）。"""
    p = Path(path)
    if not p.is_absolute():
        p = (base / p).resolve()
    if not p.exists():
        sys.exit(f"图片不存在: {p}")
    if box:
        try:
            from PIL import Image
            with Image.open(p) as im:
                w, h = im.size
            scale = min(box[0] / w, box[1] / h)  # contain 模式下的实际缩放
            if scale > 1.15:
                print(f"⚠ 图片分辨率偏低，会被放大 {scale:.1f} 倍可能发糊: {p.name} ({w}x{h})，建议换更高清的图或 PDF 用更高 -r 重转")
        except Exception:
            pass
    return p.as_uri()


A4_RATIO = 1 / 1.4142  # 宽/高


def img_ratio(path, base):
    p = Path(path)
    if not p.is_absolute():
        p = (base / p).resolve()
    from PIL import Image
    with Image.open(p) as im:
        w, h = im.size
    return w / h


# ---------- 公共框架 ----------
def css(t):
    return f"""
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{width:{W}px;height:{H}px;overflow:hidden}}
body{{font-family:{SANS};background:{t['blue']};position:relative;color:#222}}
.shop{{position:absolute;left:52px;top:0;height:150px;display:flex;align-items:center;color:#fff;font-family:{HEAVY};font-size:72px;letter-spacing:4px;z-index:2}}
.bgshape{{position:absolute;left:0;top:0;width:{W}px;height:{H}px;z-index:0}}
.panel{{position:absolute;left:34px;right:34px;top:140px;height:1270px;z-index:1}}
.h{{font-family:{HEAVY};font-weight:400}}
.chip{{display:inline-block;background:{t['blue']};color:#fff;padding:0 .28em;border-radius:.16em;margin:0 .08em;line-height:1.3}}
.red{{color:{t['red']};font-weight:600}}
b{{font-weight:600}}
.big{{color:{t['red']};font-family:{HEAVY};line-height:1.16;letter-spacing:2px}}
.kicker{{color:{t['kicker']};font-family:{HEAVY};letter-spacing:3px}}
/* 底部条：左「发货方式」，右「蓝色页签 + 白色胶囊」，与内板弧形咬合 */
.foot{{position:absolute;left:0;right:0;top:0;bottom:0;z-index:2;pointer-events:none}}
.foot .s{{position:absolute;left:52px;top:1412px;color:#fff;font-size:38px;font-weight:600;letter-spacing:2px}}
.foot .l{{position:absolute;left:48px;top:1462px;color:#fff;font-family:{HEAVY};font-size:92px;letter-spacing:3px;line-height:1.2}}
.foot .tab{{position:absolute;left:{FOOT_X + 36}px;top:1404px;height:92px;display:flex;align-items:center;color:#fff;font-family:{HEAVY};font-size:44px;letter-spacing:3px;white-space:nowrap}}
.foot .pill{{position:absolute;left:{FOOT_X}px;right:-60px;top:1500px;height:116px;border-radius:58px;background:#fff;display:flex;align-items:center;justify-content:center;padding-right:60px;
  color:{t['blue_dark']};font-family:{HEAVY};font-size:60px;letter-spacing:4px}}
/* 资料封面 mock */
.mock{{position:absolute;inset:0;background:{t['paper']};display:flex;flex-direction:column;align-items:center;font-family:{SERIF};color:{t['ink']}}}
.mock::before{{content:"";position:absolute;inset:18px;border:2px solid rgba(139,107,74,.55)}}
.mock::after{{content:"";position:absolute;inset:28px;border:1px solid rgba(139,107,74,.35)}}
.mock .top{{margin-top:12%;font-size:28px;letter-spacing:4px;opacity:.85}}
.mock .ttl{{margin-top:30px;text-align:center;line-height:1.25;letter-spacing:4px}}
.mock .sub{{margin-top:22px;font-size:28px;letter-spacing:3px;opacity:.8}}
.mock .ill{{position:absolute;left:22%;right:22%;bottom:44px}}
.mock .ill svg{{width:100%;height:auto;display:block}}
"""


FOOT_X = 476  # 底部右侧（页签/胶囊）起点


def bg_shape(t, shop, foot):
    """内板外形：左上让出店铺名页签、右侧弧形抬高；底部左侧让出发货信息、右侧弧形下沉。"""
    L, R, r = 34, W - 34, 30
    top_lo, top_hi = 150, 46
    tab_x = 52 + int(text_em(shop) * 76) + 40          # 店铺名页签宽度随店名长短
    tab_x = max(300, min(tab_x, 760))
    if foot:
        # 底边：页签（ribbon）右侧才弧形下沉，与顶部店铺名页签上下呼应
        rib = foot.get("ribbon", "")
        bx = FOOT_X + 36 + (int(text_em(rib) * 46 + 36) if rib else 0)
        bx = min(bx, R - r - 130)  # 页签文字很长时也给弧线留出位置
        bot_lo, bot_hi = 1484, 1392
        bottom = (f"L{R},{bot_lo - r} Q{R},{bot_lo} {R - r},{bot_lo} L{bx + 100},{bot_lo} "
                  f"C{bx + 45},{bot_lo} {bx + 55},{bot_hi} {bx},{bot_hi} L{L + r},{bot_hi} Q{L},{bot_hi} {L},{bot_hi - r}")
    else:
        b = H - 34
        bottom = f"L{R},{b - r} Q{R},{b} {R - r},{b} L{L + r},{b} Q{L},{b} {L},{b - r}"
    d = (f"M{L},{top_lo + r} Q{L},{top_lo} {L + r},{top_lo} L{tab_x},{top_lo} "
         f"C{tab_x + 55},{top_lo} {tab_x + 45},{top_hi} {tab_x + 100},{top_hi} L{R - r},{top_hi} Q{R},{top_hi} {R},{top_hi + r} "
         f"{bottom} Z")
    return f'<svg class="bgshape" viewBox="0 0 {W} {H}"><path d="{d}" fill="{t["panel"]}"/></svg>'


def frame(t, shop, body, foot, panel_h=1270):
    foot = foot or {}
    parts = []
    if foot.get("small"):
        parts.append(f'<div class="s">{rich(foot["small"])}</div>')
    if foot.get("big"):
        parts.append(f'<div class="l">{rich(foot["big"])}</div>')
    if foot.get("ribbon"):
        parts.append(f'<div class="tab">{rich(foot["ribbon"])}</div>')
    if foot.get("pill"):
        parts.append(f'<div class="pill">{rich(foot["pill"])}</div>')
    fonts = ('@font-face{font-family:"LTHeavy";src:local("Lantinghei SC Heavy"),local("FZLTTHK--GBK1-0"),'
             'local("Microsoft YaHei Bold"),local("MicrosoftYaHei-Bold"),local("SimHei");font-weight:400}')
    return f"""<!doctype html><html><head><meta charset="utf-8"><style>{fonts}{css(t)}</style></head>
<body>{bg_shape(t, shop, foot if parts else None)}<div class="shop">{rich(shop)}</div><div class="panel" style="height:{panel_h}px">{body}</div>
<div class="foot">{''.join(parts)}</div></body></html>"""


ILLUSTRATION = """<svg viewBox="0 0 440 260" xmlns="http://www.w3.org/2000/svg">
 <ellipse cx="220" cy="238" rx="190" ry="14" fill="#d9cdb6" opacity=".6"/>
 <path d="M40 200 Q130 170 220 196 Q310 170 400 200 L400 226 Q310 198 220 224 Q130 198 40 226Z" fill="#c9a77a"/>
 <path d="M48 196 Q130 150 216 186 L216 214 Q130 180 48 214Z" fill="#fffdf6" stroke="#b08a5a" stroke-width="2"/>
 <path d="M392 196 Q310 150 224 186 L224 214 Q310 180 392 214Z" fill="#fffdf6" stroke="#b08a5a" stroke-width="2"/>
 <g stroke="#c8b493" stroke-width="2"><path d="M80 190 Q140 168 196 190"/><path d="M84 200 Q140 180 196 200"/><path d="M244 190 Q300 168 360 190"/><path d="M244 200 Q300 180 356 200"/></g>
 <g transform="translate(300 60) rotate(28)"><rect width="16" height="120" fill="#e9b949"/><polygon points="0,120 16,120 8,142" fill="#f3d9a6"/><polygon points="5,134 11,134 8,142" fill="#333"/><rect width="16" height="14" y="-14" fill="#e88a8a"/></g>
 <g fill="#e2a23a" opacity=".9"><path d="M70 90 q20 -30 40 0 q-20 20 -40 0z"/><path d="M110 60 q18 -26 36 0 q-18 18 -36 0z"/><path d="M60 130 q16 -24 32 0 q-16 16 -32 0z"/></g>
 <path d="M100 160 Q96 110 128 70" stroke="#8a6a3a" stroke-width="3" fill="none"/>
 <circle cx="372" cy="52" r="16" fill="#8fb6e8" opacity=".7"/>
</svg>"""


def product_box(prod, base, t, box_w):
    # box_w 为 (宽, 高) 时用于清晰度检查；仿封面只用宽度
    bw = box_w[0] if isinstance(box_w, tuple) else box_w
    """产品展示：用户图片优先，否则资料封面 mock。"""
    prod = prod or {}
    if prod.get("image"):
        # 默认 contain：等比缩放、完整显示，绝不拉伸、不裁切
        # 展示框已按图片原比例算好，contain 只是兜底，保证不拉伸不裁切
        return f'<img src="{img_src(prod["image"], base, box_w)}" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain;background:#fff">'
    c = prod.get("cover", {})
    title = c.get("title", ["资料合集"])
    if isinstance(title, str):
        title = [title]
    fs = fit(title, bw - 110, 84, 40)
    ttl = "<br>".join(rich(x) for x in title)
    top = f'<div class="top">{rich(c["top"])}</div>' if c.get("top") else '<div class="top">— ★ ★ ★ —</div>'
    sub = f'<div class="sub">{rich(c["sub"])}</div>' if c.get("sub") else ""
    toc = ""
    if c.get("items"):
        rows = "".join(f'<div style="display:flex;gap:10px;align-items:baseline"><span>{rich(x)}</span><span style="flex:1;border-bottom:2px dotted rgba(139,107,74,.5)"></span></div>' for x in c["items"][:4])
        toc = f'<div style="margin-top:36px;width:70%;font-size:25px;line-height:1.8;letter-spacing:2px;opacity:.8;z-index:1">{rows}</div>'
    return f'<div class="mock">{top}<div class="ttl" style="font-size:{fs}px">{ttl}</div>{sub}{toc}<div class="ill">{ILLUSTRATION}</div></div>'


# ---------- 页面类型 ----------
def page_cover(p, base, t):
    # 右栏：一页资料展示位，卡片按资料页原比例贴合（默认 A4 竖版），不拉伸
    prod = p.get("product") or {}
    ratio = img_ratio(prod["image"], base) if prod.get("image") else A4_RATIO
    col_x, col_w, col_top, col_bot = 556, 604, 40, 1250  # 内板坐标
    note = p.get("card_note")  # 资料卡下方的一行补充（如页数）
    if note:
        col_bot -= 96
    over = 50                                              # 边框+内边距
    iw = min(col_w - over, (col_bot - col_top - over) * ratio)
    ih = iw / ratio
    cw, chh = iw + over, ih + over
    cx = col_x + (col_w - cw) / 2
    ctop = col_top + (col_bot - col_top - chh) / 2 + 10
    # 左栏
    left_w = 510
    kicker = p.get("kicker")
    title = p.get("title", [])
    if isinstance(title, str):
        title = [title]
    tfs = fit(title, left_w - 30, 190, 90)
    html_left = []
    if kicker:
        kfs = fit([kicker], left_w - 30, 80, 44)
        html_left.append(f'<div class="kicker" style="font-size:{kfs}px;margin-bottom:22px">{rich(kicker)}</div>')
    html_left.append(f'<div class="big" style="font-size:{tfs}px">{"<br>".join(rich(x) for x in title)}</div>')
    if p.get("slogan"):
        html_left.append(f'<div style="margin-top:34px;font-size:44px;font-weight:600;color:{t["blue_dark"]};letter-spacing:2px">{rich(p["slogan"])}</div>')
    # 封面左下：3 条产品卖点（不写注意事项），白卡 + 蓝色对勾，醒目好读
    sell = p.get("selling", [])[:4]
    sb = ""
    if sell:
        sfs = fit(sell, left_w - 200, 46, 30)  # 特黑字较宽，多留余量
        rows = "".join(f'''<div style="display:flex;align-items:center;gap:18px;background:#fff;border-radius:16px;padding:18px 24px">
            <div style="flex:none;width:{sfs+14}px;height:{sfs+14}px;border-radius:12px;background:{t['blue']};display:flex;align-items:center;justify-content:center">
              <svg width="{sfs*0.7}" height="{sfs*0.7}" viewBox="0 0 10 10"><path d="M1.5 5.2 L4 7.6 L8.6 2.4" stroke="#fff" stroke-width="1.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg></div>
            <div style="font-family:{HEAVY};font-size:{sfs}px;color:#1d1d1f;line-height:1.3;letter-spacing:1px;white-space:nowrap">{rich(x)}</div></div>''' for x in sell)
        sb = f'<div style="display:flex;flex-direction:column;gap:22px;width:{left_w - 44}px">{rows}</div>'
    # 左栏整组垂直居中，标题与卖点紧凑成一体
    # 左栏与右侧资料卡上下对齐：标题组顶边对齐卡片顶边，卖点底边对齐卡片底边
    # 可选：科目方块（合集类商品，如「历史 地理 道法 生物」），填在标题与卖点之间
    subj = p.get("subjects") or []
    sj = ""
    if subj:
        ncol = 2 if len(subj) in (2, 4) else min(len(subj), 3)
        tw = (left_w - 44 - 20 * (ncol - 1)) / ncol
        jfs = fit(subj, tw - 40, 72, 40)
        tiles = "".join(f'<div style="background:{t["blue"]};color:#fff;font-family:{HEAVY};font-size:{jfs}px;letter-spacing:6px;text-align:center;padding:16px 0;border-radius:18px">{rich(x)}</div>' for x in subj)
        sj = f'<div style="display:grid;grid-template-columns:repeat({ncol},1fr);gap:20px;width:{left_w - 44}px">{tiles}</div>'
    # 资料卡偏矮（横向拼图等）时左栏不再跟卡片等高，改为至少 1120px、在内板里居中，避免标题/方块/卖点挤在一起
    ltop, lh = ctop - 14, chh + 14
    if lh < 1120:
        lh = 1120
        ltop = (col_top + 1250 - lh) / 2 + 10
    left = f'<div style="position:absolute;left:52px;top:{ltop}px;height:{lh}px;width:{left_w}px;display:flex;flex-direction:column;justify-content:space-between"><div>{"".join(html_left)}</div>{sj}{sb}</div>'

    tag = p.get("tag", "笔记同款")
    tag_html = ""
    if tag:
        tag_html = f'''<div style="position:absolute;left:{cx + cw/2}px;top:{ctop-44}px;transform:translateX(-50%);z-index:3;
          background:{t['blue']};color:#fff;font-family:{HEAVY};font-size:44px;letter-spacing:4px;padding:8px 44px;border-radius:40px;
          white-space:nowrap">{rich(tag)}</div>'''
    card = f'''<div style="position:absolute;left:{cx}px;width:{cw}px;top:{ctop}px;height:{chh}px;border:8px solid {t['blue']};border-radius:30px;background:#fff;padding:16px;">
      <div style="position:relative;width:{iw}px;height:{ih}px;overflow:hidden;outline:2px solid {t['blue_dark']}22">{product_box(prod, base, t, (iw, ih))}</div></div>'''
    note_html = ""
    if note:
        note_html = f'''<div style="position:absolute;left:{cx + cw/2}px;top:{ctop + chh + 26}px;transform:translateX(-50%);white-space:nowrap;
          background:#fff;border:5px solid {t['blue']};border-radius:40px;padding:6px 40px;color:{t['blue_dark']};font-family:{HEAVY};font-size:40px;letter-spacing:2px">{rich(note)}</div>'''
    return left + card + tag_html + note_html


def heading_block(p, t, top=70):
    title = p.get("heading", "")
    lines = title if isinstance(title, list) else [title]
    fs = fit(lines, 1060, 120, 70)
    out = f'<div class="big" style="position:absolute;left:0;right:0;top:{top}px;text-align:center;font-size:{fs}px">{"<br>".join(rich(x) for x in lines)}</div>'
    h = int(fs * 1.12 * len(lines))
    if p.get("sub"):
        out += f'''<div style="position:absolute;left:0;right:0;top:{top + h + 24}px;text-align:center">
          <span style="display:inline-block;background:{t['blue']};color:#fff;font-family:{HEAVY};font-size:42px;letter-spacing:3px;padding:10px 40px;border-radius:40px">{rich(p["sub"])}</span></div>'''
        h += 24 + 72
    return out, top + h


def images_row(imgs, base, t, top, height, captions=None):
    """内页图一行排开，每个框按图片原比例贴合，统一高度，不拉伸不裁切。"""
    if not imgs:
        return "", top
    gap, pad = 26, 32
    ratios = [img_ratio(i, base) for i in imgs]
    # 角标：图片左上角的版本/页面名（如 背诵版 / 默写版），不占额外高度
    caps = []
    for k in range(len(imgs)):
        c = (captions or [])[k] if k < len(captions or []) else ""
        caps.append(f'<div style="position:absolute;left:50%;top:-30px;transform:translateX(-50%);white-space:nowrap;background:{t["red"]};color:#fff;font-size:32px;font-weight:600;padding:2px 24px;border-radius:30px;font-family:{HEAVY};letter-spacing:2px">{rich(c)}</div>' if c else "")
    max_w = 1094
    ih = height - pad
    total = sum(ih * r + pad for r in ratios) + gap * (len(imgs) - 1)
    if total > max_w:  # 太宽就整体缩小高度
        ih = (max_w - gap * (len(imgs) - 1) - pad * len(imgs)) / sum(ratios)
    cells = "".join(
        f'''<div style="flex:none;position:relative;border:6px solid {t['blue']};border-radius:16px;background:#fff;padding:10px;">
          <img src="{img_src(i, base, (ih * r, ih))}" style="display:block;width:{ih * r}px;height:{ih}px;object-fit:contain">{cap}</div>'''
        for i, r, cap in zip(imgs, ratios, caps))
    return f'<div style="position:absolute;left:40px;right:40px;top:{top}px;display:flex;justify-content:center;gap:{gap}px">{cells}</div>', top + ih + pad + 34


def page_points(p, base, t):
    head, y = heading_block(p, t)
    imgs = p.get("images", [])[:3]
    pts = p.get("points", [])
    if imgs:
        # 先保证每张卖点卡至少 150px 高，图片只用剩余空间，避免文字挤出卡片
        need = len(pts) * 150 + 24 * max(len(pts) - 1, 0) + 50
        img_h = min(p.get("image_h", 420), 1270 - need - (y + 40) - 34)
        row, y = images_row(imgs, base, t, y + 40, img_h, p.get("captions"))
    else:
        row, y = "", y + 50
    avail = 1270 - y - 50
    gap = 24
    ch = (avail - gap * (len(pts) - 1)) / max(len(pts), 1)
    title_fs = 50 if ch > 200 else 44
    desc_fs = 34 if ch > 200 else 31
    cards = []
    for i, pt in enumerate(pts):
        num = pt.get("num", f"{i+1:02d}")
        cards.append(f'''<div style="height:{ch}px;background:#fff;border-radius:20px;display:flex;align-items:center;padding:0 36px;gap:30px;">
          <div style="flex:none;width:104px;height:104px;border-radius:50%;background:{t['blue']};color:#fff;display:flex;align-items:center;justify-content:center;font-family:{HEAVY};font-size:46px;">{html.escape(str(num))}</div>
          <div style="flex:1;min-width:0">
            <div style="font-family:{HEAVY};font-size:{title_fs}px;color:{t['blue_dark']};letter-spacing:2px;line-height:1.3">{rich(pt.get("title",""))}</div>
            <div style="margin-top:10px;font-size:{desc_fs}px;color:#444;line-height:1.45">{rich(pt.get("desc",""))}</div>
          </div></div>''')
    body = f'<div style="position:absolute;left:40px;right:40px;top:{y}px;display:flex;flex-direction:column;gap:{gap}px">{"".join(cards)}</div>'
    return head + row + body


def page_grid(p, base, t):
    head, y = heading_block(p, t)
    items = p.get("items", [])
    cols = p.get("cols", 2 if len(items) <= 8 else 3)
    rows = -(-len(items) // cols)
    gap = 22
    avail = 1270 - y - 90
    note_h = 0
    if p.get("note"):
        note_h = 90
        avail -= note_h
    ch = (avail - gap * (rows - 1)) / max(rows, 1)
    name_fs = 48 if cols == 2 else 40
    desc_fs = 30 if cols == 2 else 27
    cells = []
    for it in items:
        if isinstance(it, str):
            it = {"name": it}
        desc = f'<div style="margin-top:8px;font-size:{desc_fs}px;color:#555;line-height:1.35">{rich(it["desc"])}</div>' if it.get("desc") else ""
        cells.append(f'''<div style="height:{ch}px;background:#fff;border-radius:20px;border-left:14px solid {t['blue']};padding:0 28px;display:flex;flex-direction:column;justify-content:center;">
          <div style="font-family:{HEAVY};font-size:{name_fs}px;color:{t['blue_dark']};line-height:1.25;letter-spacing:2px">{rich(it.get("name",""))}</div>{desc}</div>''')
    grid = f'<div style="position:absolute;left:40px;right:40px;top:{y + 44}px;display:grid;grid-template-columns:repeat({cols},1fr);gap:{gap}px">{"".join(cells)}</div>'
    note = ""
    if p.get("note"):
        note = f'<div style="position:absolute;left:40px;right:40px;bottom:40px;text-align:center;font-size:38px;font-weight:600;color:{t["blue_dark"]}">{rich(p["note"])}</div>'
    return head + grid + note


DEFAULT_NOTICE = [
    {"title": "关于商品", "body": "商品为{{电子版文件}}，无实物，无快递。"},
    {"title": "关于发货", "body": "下单后订单自动发货，默认发**百度网盘链接**，如需其他方式请联系在线客服。"},
    {"title": "如何获取", "body": [
        "点开订单物流详情发货内容，点击发货链接内容后面的“复制全部”，打开网盘即可查看保存。",
        "下单后自动发送链接到客服聊天框，长按链接可复制。复制后打开“百度网盘app”，会自动识别链接，转存后即可下载。"]},
    {"title": "关于退款", "body": "购前请看清再拍，一经下单表示同意以上内容。虚拟产品具有可复制性，发货后{{不支持}}以“看错了、拍过了、以为是纸质版”等任何理由退款。"},
]


def page_notice(p, base, t):
    """购前须知说明页：大标题 + 若干「▶ 小标题 + 白色说明卡」。"""
    panel_h = p["_panel_h"]
    secs = p.get("sections") or DEFAULT_NOTICE
    text_w = 1174 - 80 - 76  # 卡片内文字宽度

    def paras(sec):
        b = sec.get("body", "")
        if isinstance(b, list):
            return [f"{i+1}.{x}" for i, x in enumerate(b)]
        return str(b).split("\n")

    head_h = 190
    # 从大到小试字号，直到全部内容放得下
    for fs in range(42, 25, -1):
        lh = fs * 1.55
        total = head_h
        for sec in secs:
            lines = sum(math.ceil(text_em(x) * fs / text_w) for x in paras(sec))
            total += 50 * 1.3 + 18 + 52 + lines * lh + 30
        if total <= panel_h - 70:
            break
    title_fs = min(54, fs + 12)
    blocks = []
    for sec in secs:
        body = "".join(f'<div>{rich(x)}</div>' for x in paras(sec))
        blocks.append(f'''<div>
          <div style="display:flex;align-items:center;justify-content:center;gap:22px;font-family:{HEAVY};font-size:{title_fs}px;color:{t['blue_dark']};letter-spacing:4px">
            <svg width="{title_fs*0.8}" height="{title_fs*0.8}" viewBox="0 0 10 10"><path d="M1 1 L9 5 L1 9 L3 5Z" fill="{t['red']}"/></svg>{rich(sec.get("title",""))}</div>
          <div style="margin-top:18px;background:#fff;border-radius:22px;padding:26px 38px;font-size:{fs}px;line-height:1.55;color:#222;">{body}</div>
        </div>''')
    heading = p.get("heading", "购前须知")
    bang = '''<svg width="62" height="86" viewBox="0 0 62 86" style="margin-left:14px;transform:rotate(12deg)">
      <rect x="6" y="4" width="18" height="56" rx="8" fill="#FFC21A" stroke="#222" stroke-width="3"/><circle cx="15" cy="74" r="9" fill="#FFC21A" stroke="#222" stroke-width="3"/>
      <rect x="36" y="10" width="16" height="48" rx="7" fill="#FFC21A" stroke="#222" stroke-width="3"/><circle cx="44" cy="70" r="8" fill="#FFC21A" stroke="#222" stroke-width="3"/></svg>'''
    return f'''<div style="position:absolute;left:40px;right:40px;top:64px;bottom:36px;display:flex;flex-direction:column;justify-content:space-between">
      <div style="display:flex;align-items:center;height:{head_h - 40}px"><span class="big" style="font-size:112px">{rich(heading)}</span>{bang}</div>
      {"".join(blocks)}</div>'''


PAGES = {"cover": page_cover, "points": page_points, "grid": page_grid, "notice": page_notice}


def load_template(name):
    if not name:
        return {}
    f = TEMPLATES / f"{name}.json"
    if not f.exists():
        sys.exit(f"没有主图模板「{name}」。可用：" + "、".join(p.stem for p in sorted(TEMPLATES.glob("*.json"))))
    return json.loads(f.read_text(encoding="utf-8"))


def list_templates():
    for f in sorted(TEMPLATES.glob("*.json")):
        d = json.loads(f.read_text(encoding="utf-8"))
        print(f"- {f.stem}：{d.get('风格', '')}｜适合 {d.get('适合', '')}")


def pdf_pages(pdf, out, n=6):
    out = Path(out); out.mkdir(parents=True, exist_ok=True)
    export_pdf.pdf_to_png(pdf, str(out / "page"), dpi=200, first=1, last=int(n))
    for f in sorted(out.glob("page-*.png")):
        print(f)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    if sys.argv[1] == "templates":
        return list_templates()
    if sys.argv[1] == "pages":
        return pdf_pages(*sys.argv[2:5])
    spec_path = Path(sys.argv[1]).resolve()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    base = spec_path.parent
    out = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else base
    (out / "html").mkdir(parents=True, exist_ok=True)

    shop = spec.get("shop")
    if not shop:
        sys.exit("spec.json 缺少 shop（店铺名称需用户手动提供）")
    tpl = load_template(spec.get("模板"))
    t = {**DEFAULT_THEME, **tpl.get("theme", {}), **spec.get("theme", {})}
    foot_default = {**tpl.get("footer", {}), **spec.get("footer", {})}

    for i, p in enumerate(spec["pages"], 1):
        kind = p.get("type", "points")
        if kind not in PAGES:
            sys.exit(f"未知页面类型: {kind}（可用: {', '.join(PAGES)}）")
        # 购前须知页默认不放底部条，内板拉长装下全部说明；任何页 "footer": false 都可去掉底部条
        hide_foot = p.get("footer") is False or (kind == "notice" and "footer" not in p)
        panel_h = H - 140 - 34 if hide_foot else 1270
        p["_panel_h"] = panel_h
        body = PAGES[kind](p, base, t)
        foot = {} if hide_foot else {**foot_default, **(p.get("footer") or {})}
        doc = frame(t, p.get("shop", shop), body, foot, panel_h)
        hp = out / "html" / f"{i:02d}.html"
        hp.write_text(doc, encoding="utf-8")
        png = out / f"{i:02d}.png"
        r = subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
                            "--allow-file-access-from-files", "--force-device-scale-factor=1",
                            f"--window-size={W},{H}", "--virtual-time-budget=3000",
                            f"--screenshot={png}", hp.as_uri()],
                           capture_output=True, text=True, timeout=120)
        if not png.exists():
            sys.exit(f"第 {i} 页截图失败:\n{r.stderr[-800:]}")
        print(f"✓ {png}")


if __name__ == "__main__":
    main()
