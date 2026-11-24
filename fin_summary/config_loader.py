"""映射配置加载与基础 Schema 校验（VC-08）。"""
from __future__ import annotations

import hashlib
import json
from decimal import Decimal
from pathlib import Path

from .models import ItemSpec, Source

REQUIRED_TOP = {"config_version", "sheets", "items"}
VALID_KINDS = {"VALUE", "SUM", "CONST", "MANUAL", "FORMULA"}
VALID_SOURCE_TYPES = {"IS", "BS", "CF", "TB"}
TB_FIELDS = {"年初借", "年初贷", "期初借", "期初贷", "本期借", "本期贷",
             "累计借", "累计贷", "期末借", "期末贷"}


class MappingError(Exception):
    pass


def load_mapping(path: Path | None = None) -> tuple[dict, list[ItemSpec], str, str]:
    """加载 mapping.json，返回 (原始配置, 70 项规格列表, 版本号, sha256)。"""
    if path is None:
        cand = Path("config/mapping.json")
        path = cand if cand.exists() else Path(__file__).parent / "resources" / "mapping.json"
    raw_bytes = path.read_bytes()
    sha = hashlib.sha256(raw_bytes).hexdigest()[:16]
    try:
        cfg = json.loads(raw_bytes.decode("utf-8"))
    except json.JSONDecodeError as e:
        raise MappingError(f"mapping.json 不是合法 JSON: {e}") from e

    missing = REQUIRED_TOP - set(cfg)
    if missing:
        raise MappingError(f"mapping.json 缺少必需键: {missing}")

    items: list[ItemSpec] = []
    rows_seen: set[int] = set()
    for i, raw in enumerate(cfg["items"]):
        try:
            row = int(raw["row"])
            name = str(raw["name"])
            kind = str(raw["kind"])
        except KeyError as e:
            raise MappingError(f"items[{i}] 缺少字段 {e}") from e
        if not 1 <= row <= 70:
            raise MappingError(f"items[{i}] 行次 {row} 超出 1..70")
        if row in rows_seen:
            raise MappingError(f"行次 {row} 重复映射")
        rows_seen.add(row)
        if kind not in VALID_KINDS:
            raise MappingError(f"行次 {row} kind 非法: {kind}")

        sources = []
        for s in raw.get("sources", []):
            stype = str(s.get("type", ""))
            if stype not in VALID_SOURCE_TYPES:
                raise MappingError(f"行次 {row} source.type 非法: {stype}")
            field = str(s.get("field", ""))
            if stype == "TB" and field not in TB_FIELDS:
                raise MappingError(f"行次 {row} TB source.field 非法: {field}")
            sources.append(Source(
                type=stype,
                match_name=str(s.get("match_name", "")),
                aliases=[str(a) for a in s.get("aliases", [])],
                field=field,
                sign=int(s.get("sign", 1)),
                missing_ok=bool(s.get("missing_ok", False)),
                match_mode=str(s.get("match_mode", "normalized_exact")),
            ))
        if kind in {"VALUE", "SUM"} and not sources:
            raise MappingError(f"行次 {row} {kind} 缺少 sources")
        if kind == "FORMULA" and not str(raw.get("formula", "")).strip():
            raise MappingError(f"行次 {row} FORMULA 缺少 formula")

        items.append(ItemSpec(
            row=row, name=name, kind=kind, sources=sources,
            formula=str(raw.get("formula", "")),
            const_value=Decimal(str(raw.get("const_value", "0"))),
            divide=bool(raw.get("divide", True)),
            max_floor=Decimal(str(raw["max_floor"])) if "max_floor" in raw else None,
            note=str(raw.get("note", "")),
        ))

    if rows_seen != set(range(1, 71)):
        lack = sorted(set(range(1, 71)) - rows_seen)
        raise MappingError(f"行次未覆盖: {lack}")

    return cfg, items, str(cfg["config_version"]), sha
