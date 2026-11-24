"""项目/科目名称规范化与匹配（设计文档第 5 章）。"""
from __future__ import annotations

import re
from dataclasses import dataclass

# 全角→半角标点统一，同时去除全角/半角空格
_PUNCT_MAP = {
    "（": "(", "）": ")", "：": ":", "，": ",", "－": "-", "—": "-",
    "“": '"', "”": '"', "‘": "'", "’": "'", "、": "、",
}
_SPACES = ("\u3000", " ", "\t", "\u00a0")

# 栏目前缀（标点统一后的形态）
_PREFIXES = (
    "一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
    "加:", "减:", "其中:", "*", "＋", "+",
)


def normalize(name: str | None) -> str:
    """规范化报表项目名称：去空格、统一标点、去括号说明、去栏目前缀。"""
    if name is None:
        return ""
    t = str(name)
    for sp in _SPACES:
        t = t.replace(sp, "")
    for a, b in _PUNCT_MAP.items():
        t = t.replace(a, b)
    t = t.strip("*")
    while True:  # 去除括号及其内容（可嵌套）
        new = re.sub(r"\([^()]*\)", "", t)
        if new == t:
            break
        t = new
    changed = True
    while changed:  # 去除重复前缀
        changed = False
        for p in _PREFIXES:
            if t.startswith(p):
                t = t[len(p):]
                changed = True
    return t.strip("*-—").strip()


@dataclass
class NameHit:
    row: int          # excel 行号
    name_col: int     # 项目名称所在列（用于判断左栏/右栏）
    text: str         # 原始文本


def build_index(cells: list[NameHit]) -> dict[str, list[NameHit]]:
    """规范化名称 → 命中列表。"""
    index: dict[str, list[NameHit]] = {}
    for cell in cells:
        key = normalize(cell.text)
        if key:
            index.setdefault(key, []).append(cell)
    return index


def match_unique(
    index: dict[str, list[NameHit]],
    match_name: str,
    aliases: list[str] | None = None,
    mode: str = "normalized_exact",
) -> tuple[NameHit | None, list[NameHit], str]:
    """按配置匹配，返回 (首选命中, 全部候选, 说明)。

    normalized_exact: 规范化后精确匹配；命中多个时选左栏优先并返回全部候选。
    contains: 规范化后包含匹配兜底。
    """
    aliases = aliases or []
    for key in [match_name, *aliases]:
        cands = index.get(normalize(key), [])
        if cands:
            if len(cands) > 1:
                best = min(cands, key=lambda c: c.name_col)
                return best, cands, f"精确命中 {len(cands)} 处，取最左列"
            return cands[0], cands, "精确命中"
    if mode == "contains":
        target = normalize(match_name)
        cands = [c for k, lst in index.items() if target and target in k for c in lst]
        if len(cands) == 1:
            return cands[0], cands, "包含匹配"
        if len(cands) > 1:
            best = min(cands, key=lambda c: (len(normalize(c.text)), c.name_col))
            return best, cands, f"包含匹配命中 {len(cands)} 处，取最短名称"
    return None, [], "未命中"
