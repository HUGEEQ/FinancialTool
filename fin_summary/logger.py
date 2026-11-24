"""日志留存：run_id、计算日志 JSONL、校验日志 JSONL、审计摘要。"""
from __future__ import annotations

import hashlib
import json
import secrets
import sys
from datetime import datetime
from pathlib import Path

from . import APP_NAME, __version__
from .models import CalcResult, Issue


def default_log_root() -> Path:
    """日志根目录：打包客户端写入用户主目录（避免双击启动时 CWD 不可写）；开发模式用 ./logs。"""
    if getattr(sys, "frozen", False):
        return Path.home() / "FinSummaryTool" / "logs"
    return Path("logs")


def gen_run_id() -> str:
    return f"{datetime.now():%Y%m%d_%H%M%S}_{secrets.token_hex(3)}"


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


class RunLogger:
    def __init__(self, run_id: str, month: str, base_dir: Path | None = None):
        self.run_id = run_id
        self.month = month
        self.base = (base_dir or default_log_root())
        today = datetime.now()
        self._dirs = {
            "app": self.base / "app",
            "calc": self.base / "calc" / f"{today:%Y/%m}",
            "validation": self.base / "validation" / f"{today:%Y/%m}",
            "error": self.base / "error",
            "audit": self.base / "audit" / f"{today:%Y/%m}",
        }
        for d in self._dirs.values():
            d.mkdir(parents=True, exist_ok=True)
        self.calc_path = self._dirs["calc"] / f"{run_id}.jsonl"
        self.validation_path = self._dirs["validation"] / f"{run_id}.jsonl"
        self.calc_path.touch()
        self.validation_path.touch()
        self.started_at = today.isoformat(timespec="seconds")

    def write_calc(self, results: list[CalcResult]) -> None:
        lines = []
        for r in results:
            rec = {
                "run_id": self.run_id,
                "stage": "CALC_ITEM",
                "row": r.row,
                "item": r.name,
                "status": r.status.value,
                "expression": r.expression,
                "sources": [
                    {"type": s.type, "ref": s.ref, "field": s.field,
                     "raw_amount": str(s.raw_amount), "converted": str(s.converted), "sign": s.sign}
                    for s in r.sources
                ],
                "raw_result": str(r.raw) if r.raw is not None else None,
                "rounded": str(r.rounded) if r.rounded is not None else None,
                "messages": r.messages,
            }
            lines.append(json.dumps(rec, ensure_ascii=False))
        self.calc_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    def write_validation(self, issues: list[Issue], row_statuses: dict[int, str]) -> None:
        lines = []
        for i in issues:
            lines.append(json.dumps({
                "run_id": self.run_id, "stage": "VALIDATION", "rule": i.rule,
                "level": i.level.value, "title": i.title, "detail": i.detail,
                "rows": i.rows,
            }, ensure_ascii=False))
        lines.append(json.dumps({
            "run_id": self.run_id, "stage": "ROW_STATUS", "matrix": row_statuses,
        }, ensure_ascii=False))
        self.validation_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    def write_summary(self, payload: dict) -> Path:
        payload.update({
            "run_id": self.run_id,
            "app": APP_NAME,
            "version": __version__,
            "month": self.month,
            "started_at": self.started_at,
            "finished_at": datetime.now().isoformat(timespec="seconds"),
        })
        path = self._dirs["audit"] / f"{self.run_id}_summary.json"
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return path
