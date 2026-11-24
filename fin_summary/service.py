"""应用服务层：编排 计算 → 预览数据 → 导出。不依赖 UI。"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from .config_loader import MappingError, load_mapping
from .engine import compute
from .loader import Inputs, ParseError, layout_snapshot, load_inputs
from .logger import RunLogger, file_sha256, gen_run_id
from .models import CalcResult, Issue, Level, RowStatus
from .validator import validate

MANUAL_ROWS = (57, 58, 59, 60)


@dataclass
class Outcome:
    """一次计算的完整产物（预览所需全部数据）。"""
    run_id: str
    month: str
    items: list          # list[ItemSpec]
    config: dict
    mapping_version: str
    mapping_sha: str
    inputs: Inputs
    results: list[CalcResult]
    issues: list[Issue]
    manual: dict[int, Decimal]
    acknowledged: set[int] = field(default_factory=set)   # 已"已知悉"的 ⚠ 行
    exported_path: str | None = None
    calc_log: str = ""
    validation_log: str = ""
    audit_summary: str = ""

    # ---- 汇总 ----
    def counts(self) -> dict[str, int]:
        c = {"PASS": 0, "WARN": 0, "FAIL": 0, "PENDING": 0, "CONST": 0}
        for r in self.results:
            c[r.status.value] += 1
        return c

    def e_count(self) -> int:
        return sum(1 for i in self.issues if i.level == Level.E)

    def table_issues(self) -> list[Issue]:
        return [i for i in self.issues if not i.rows]

    def can_export(self) -> tuple[bool, str]:
        """导出条件：无错误、无待填报、且全部可确认项已人工确认（V2.10 逐项确认）。"""
        fails, pendings = [], []
        for r in self.results:
            if r.status == RowStatus.FAIL:
                fails.append(f"行 {r.row}「{r.name}」")
            elif r.status == RowStatus.PENDING:
                pendings.append(f"行 {r.row}（{r.name}）")
        done, total = self.confirm_progress()
        unconfirmed = total - done
        if not (fails or pendings or unconfirmed):
            return True, "可以导出"
        parts = []
        if fails:
            parts.append(f"校验失败 {len(fails)} 项: " + "、".join(fails[:3]) + ("…" if len(fails) > 3 else ""))
        if pendings:
            parts.append(f"人工数据待填报 {len(pendings)} 项: " + "、".join(pendings))
        if unconfirmed:
            parts.append(f"尚有 {unconfirmed} 项未人工确认（已确认 {done}/{total}）")
        return False, "；".join(parts)

    def recompute(self, manual: dict[int, Decimal]) -> None:
        """人工数据变化后重算（保留仍为可确认状态的已确认行）。"""
        self.manual = dict(manual)
        self.results = compute(self.items, self.inputs, self.config, self.manual)
        self.issues = validate(self.inputs, self.results, self.config, self.items)
        self.acknowledged &= self.confirmable_rows()

    def row_matrix(self) -> dict[int, str]:
        return {r.row: r.status.value for r in self.results}

    def confirmable_rows(self) -> set[int]:
        """可人工确认的行：正常/警告/固定值（错误行须先修正，待填报行须先补数）。"""
        return {r.row for r in self.results
                if r.status in (RowStatus.PASS, RowStatus.WARN, RowStatus.CONST)}

    def confirm_progress(self) -> tuple[int, int]:
        """返回 (已确认数, 需确认总数)。"""
        confirmable = self.confirmable_rows()
        return len(self.acknowledged & confirmable), len(confirmable)


def compute_run(inputs_paths: dict, month: str, manual: dict[int, Decimal] | None = None,
                mapping_path: str | None = None, log_dir: str | None = None,
                progress=None) -> Outcome:
    """计算主流程：加载配置 → 解析输入 → 计算 → 校验 → 写日志。失败抛 ParseError/MappingError。"""
    manual = {k: Decimal(str(v)) for k, v in (manual or {}).items()}
    run_id = gen_run_id()
    logger = RunLogger(run_id, month, Path(log_dir) if log_dir else None)

    def report(step: str):
        if progress:
            progress(step)

    report("加载映射配置")
    cfg, items, version, sha = load_mapping(Path(mapping_path) if mapping_path else None)

    report("解析输入表")
    inputs = load_inputs(inputs_paths)

    report("逐项计算")
    results = compute(items, inputs, cfg, manual)

    report("勾稽校验")
    issues = validate(inputs, results, cfg, items)

    logger.write_calc(results)
    logger.write_validation(issues, {r.row: r.status.value for r in results})
    logger.write_summary({
        "status": "已计算未导出",
        "inputs": {"workbook": inputs.workbook_path, "sha256_16": inputs.sha256},
        "structure": layout_snapshot(inputs),
        "mapping": {"version": version, "sha256_16": sha},
        "validation": {
            "E": sum(1 for i in issues if i.level == Level.E),
            "W": sum(1 for i in issues if i.level == Level.W),
            "I": sum(1 for i in issues if i.level == Level.I),
            "row_matrix": {r.row: r.status.value for r in results},
        },
        "manual": {str(k): str(v) for k, v in manual.items()},
    })

    outcome = Outcome(
        run_id=run_id, month=month, items=items, config=cfg,
        mapping_version=version, mapping_sha=sha, inputs=inputs,
        results=results, issues=issues, manual=manual,
        calc_log=str(logger.calc_path), validation_log=str(logger.validation_path),
    )
    return outcome


def export_run(outcome: Outcome, outdir: str, confirmations: list[dict] | None = None) -> Path:
    """导出：全新生成工作簿 → 写临时文件 → 回读复核 → 落盘 → 审计摘要补全。"""
    ok, reason = outcome.can_export()
    if not ok:
        raise PermissionError(f"导出条件不满足: {reason}")

    out = Path(outdir)
    path = _write(outcome, out)
    outcome.exported_path = str(path)

    # 审计摘要补全
    from .logger import RunLogger  # 局部引入避免循环
    logger = RunLogger(outcome.run_id, outcome.month)
    payload = {
        "status": "已导出",
        "output": {"path": str(path), "sha256": file_sha256(path)},
        "inputs": {"workbook": outcome.inputs.workbook_path, "sha256_16": outcome.inputs.sha256},
        "structure": layout_snapshot(outcome.inputs),
        "mapping": {"version": outcome.mapping_version, "sha256_16": outcome.mapping_sha},
        "validation": {
            "E": sum(1 for i in outcome.issues if i.level == Level.E),
            "W": sum(1 for i in outcome.issues if i.level == Level.W),
            "I": sum(1 for i in outcome.issues if i.level == Level.I),
            "row_matrix": outcome.row_matrix(),
        },
        "manual": {str(k): str(v) for k, v in outcome.manual.items()},
        "confirmations": confirmations or [],
    }
    summary = logger.write_summary(payload)
    outcome.audit_summary = str(summary)
    return path


def _write(outcome: Outcome, outdir: Path) -> Path:
    from .exporter import write_output
    return write_output(outcome.results, outcome.month, outdir, outcome.run_id)
