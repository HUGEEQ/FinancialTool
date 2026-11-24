"""核心数据模型：映射项、计算结果、校验问题。"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum


class Level(str, Enum):
    E = "E"
    W = "W"
    I = "I"


class RowStatus(str, Enum):
    PASS = "PASS"        # ✓
    WARN = "WARN"        # ⚠ 需已知悉
    FAIL = "FAIL"        # ✗ 阻断导出
    PENDING = "PENDING"  # — 待录入（人工项）
    CONST = "CONST"      # ○ 常量项

    @property
    def icon(self) -> str:
        return {"PASS": "✓", "WARN": "⚠", "FAIL": "✗", "PENDING": "—", "CONST": "○"}[self.value]


@dataclass
class Source:
    """一条取数来源。type: IS/BS/CF（报表取数）或 TB（科目余额表）。"""
    type: str
    match_name: str = ""
    aliases: list[str] = field(default_factory=list)
    field: str = ""          # TB 字段：累计借/累计贷/期末借/期末贷/期初借/期初贷/年初借/年初贷/本期借/本期贷
    sign: int = 1
    missing_ok: bool = False  # 源行缺失时按 0 填列（I 级提示）
    match_mode: str = "normalized_exact"  # 或 contains


@dataclass
class ItemSpec:
    """输出表一个数据项的映射规格（行次 1..70 → 输出 Sheet 第 row+1 行 C 列）。"""
    row: int
    name: str
    kind: str                 # VALUE / SUM / CONST / MANUAL / FORMULA
    sources: list[Source] = field(default_factory=list)
    formula: str = ""         # FORMULA 表达式，#n 引用行次 n 的取整值
    const_value: Decimal = Decimal("0")
    divide: bool = True       # False = 不做 ÷10000 换算（人数类）
    max_floor: Decimal | None = None  # 万单位下限（固定资产投资额）
    note: str = ""

    @property
    def deps(self) -> list[int]:
        """FORMULA 引用的行次列表。"""
        if self.kind != "FORMULA":
            return []
        import re
        return sorted({int(m) for m in re.findall(r"#(\d+)", self.formula)})


@dataclass
class SourceHit:
    """一个来源的实际取数明细（用于日志与预览钻取）。"""
    type: str
    ref: str          # 报表：Sheet!项目名；TB：科目编码 科目名称
    field: str
    raw_amount: Decimal
    converted: Decimal  # ÷10000 后（sign 已乘）
    sign: int
    empty: bool = False  # 源单元格为空（区别于显式 0）
    cell: str = ""       # 读取的单元格地址，如 利润表!E24 / 科目余额表!I117


@dataclass
class CalcResult:
    row: int
    name: str
    status: RowStatus
    raw: Decimal | None = None        # 万，未舍入（None=失败）
    rounded: int | None = None        # 最终写入值
    expression: str = ""
    sources: list[SourceHit] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)


@dataclass
class Issue:
    """校验问题。rows 为挂接的输出行次；表级规则 rows 为空。"""
    rule: str
    level: Level
    title: str
    detail: str = ""
    rows: list[int] = field(default_factory=list)
