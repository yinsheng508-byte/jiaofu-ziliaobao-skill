"""项目文件夹的统一路径：所有脚本和模板都从这里取路径，不各自拼。

  输出项目/<日期>_<资料包名称>/
    进度记录.md                 做到哪、用户定过的事、下一步（换对话时先读）
    1_规划/                     阶段一 → 确认点一
      需求单.md  素材/  素材清单.md  原作者信息.txt  素材覆盖表.md  参考/  方案.md  提问记录.md
    2_样张/                     阶段二 → 确认点二
      内容稿.md  检查报告.md  排版/  封面/  预览/  样章/  样章汇报.md  选题配置/  手写笔记/
    3_成品/                     阶段三 → 确认点三（交付物都在这一层）
      <名称>（版本）.docx / .pdf  交付说明.md  手写笔记/  主图/（商品主图、详情图）
      检查/  闸门报告.md  交付检查报告.md

老项目（文件都平铺在项目根目录）自动识别，路径照旧。
"""
from pathlib import Path

STAGES = ("1_规划", "2_样张", "3_成品")


def find_root(p):
    """从项目里任意一个文件或文件夹往上找项目根目录（有 进度记录.md 或 1_规划/ 的那一层）。"""
    p = Path(p).resolve()
    for q in [p, *p.parents]:
        if (q / "1_规划").is_dir() or (q / "进度记录.md").exists():
            return q
        if q.name in STAGES:
            return q.parent
    return p if p.is_dir() else p.parent


class P:
    def __init__(self, proj, create=False):
        self.root = Path(proj)
        self.legacy = not (self.root / "1_规划").exists() and any(
            (self.root / x).exists() for x in ("素材", "内容稿.md", "成品", "素材清单.md"))
        if create and not self.legacy:
            for s in STAGES:
                (self.root / s).mkdir(parents=True, exist_ok=True)
        r = self.root
        plan = r if self.legacy else r / "1_规划"
        sample = r if self.legacy else r / "2_样张"
        final = r / "成品" if self.legacy else r / "3_成品"
        report = r if self.legacy else final / "检查"
        self.规划, self.样张, self.成品 = plan, sample, final
        self.进度记录 = r / "进度记录.md"
        # 阶段一
        self.需求单 = plan / "需求单.md"
        self.素材 = plan / "素材"
        self.素材清单 = plan / "素材清单.md"
        self.原作者信息 = plan / "原作者信息.txt"
        self.素材覆盖表 = plan / "素材覆盖表.md"
        self.参考 = plan / "参考"
        self.方案 = plan / "方案.md"
        self.提问记录 = plan / "提问记录.md"
        # 阶段二
        self.内容稿 = sample / "内容稿.md"
        self.检查报告 = sample / "检查报告.md"
        self.排版 = sample / "排版"
        self.封面 = sample / "封面"
        self.预览 = sample / "预览"
        self.样章 = sample / "样章" if not self.legacy else r / "样章"
        self.样章汇报 = sample / "样章汇报.md"
        self.选题配置 = sample / "选题配置"
        self.手写笔记 = sample / "手写笔记"
        # 阶段三
        self.手写成品 = (r / "手写笔记" / "成品") if self.legacy else final / "手写笔记"
        self.检查 = report
        self.闸门报告 = report / "闸门报告.md"
        self.交付检查报告 = report / "交付检查报告.md"
        self.交付说明 = (r if self.legacy else final) / "交付说明.md"
        self.主图 = final / "主图"            # 商品主图 / 详情图（spec.json、卖点说明.md、01.png…）


def of(path):
    """给项目里任意路径，返回这个项目的 P。"""
    return P(find_root(path))
