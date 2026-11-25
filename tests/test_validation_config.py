"""校验规则可配置性与「接受差异」测试（V2.18）。

场景：上传的 4 张表本身默认正确，但按标准财务口径不符、按本公司口径可接受——
可用三种方式放行：配置降级/关闭/容差，或界面单次「接受差异」。
"""
import json
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from fin_summary import service
from fin_summary.config_loader import MappingError, load_mapping
from fin_summary.validator import issue_key

SAMPLE = Path(__file__).resolve().parents[1] / "doc" / "财务计算模板.xlsx"
MANUAL = {57: Decimal(1346), 58: Decimal(1327), 59: Decimal(1328), 60: Decimal(1)}
pytestmark = pytest.mark.skipif(not SAMPLE.exists(), reason="样例模板不存在")


@pytest.fixture(scope="module")
def outcome():
    return service.compute_run({"workbook": str(SAMPLE)}, "2026-10", MANUAL, log_dir="logs_test")


def _mutated_cf(tmp_path) -> str:
    """构造现金流量表勾稽不符的样例：现金期末余额 +5000 元。"""
    src = load_workbook(SAMPLE)
    cf = src["现金流量表"]
    for r in range(3, cf.max_row + 1):
        for col in (1, 5):
            name = cf.cell(row=r, column=col).value
            if name and "现金的期末余额" in str(name):
                cf.cell(row=r, column=col + 2).value = (cf.cell(row=r, column=col + 2).value or 0) + 5000
                src.save(tmp_path / "mut_cf.xlsx")
                src.close()
                return str(tmp_path / "mut_cf.xlsx")
    src.close()
    raise AssertionError("未找到现金期末余额行")


def _mapping_with(tmp_path, rules: dict, tolerance=None) -> str:
    cfg = json.loads(Path("fin_summary/resources/mapping.json").read_text(encoding="utf-8"))
    cfg["validation"]["rules"].update(rules)
    if tolerance is not None:
        cfg["validation"]["tolerance"] = tolerance
    p = tmp_path / "mapping_test.json"
    p.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    return str(p)


def test_baseline_has_no_blocking(outcome):
    assert not [i for i in outcome.issues if i.level.value == "E"]


def test_severity_downgrade_unblocks(tmp_path):
    """VC-17 降为 W：差异仍在（可见），但不再阻断导出。"""
    mut = _mutated_cf(tmp_path)
    mp = _mapping_with(tmp_path, {"VC-17": {"severity": "W"}})
    out = service.compute_run({"workbook": mut}, "2026-10", MANUAL,
                              mapping_path=mp, log_dir="logs_test")
    v17 = [i for i in out.issues if i.rule == "VC-17" and "期末余额 − 期初余额" in i.title]
    assert v17 and v17[0].level.value == "W"
    out.acknowledged |= out.confirmable_rows()
    ok, reason = out.can_export()
    assert ok, reason


def test_severity_info_unblocks(tmp_path):
    """VC-17 降为 I：仅提示。"""
    mut = _mutated_cf(tmp_path)
    mp = _mapping_with(tmp_path, {"VC-17": {"severity": "I"}})
    out = service.compute_run({"workbook": mut}, "2026-10", MANUAL,
                              mapping_path=mp, log_dir="logs_test")
    v17 = [i for i in out.issues if i.rule == "VC-17" and "期末余额 − 期初余额" in i.title]
    assert v17 and v17[0].level.value == "I"


def test_disabled_rule_not_reported(tmp_path):
    """VC-17 关闭：不再产生该规则的问题。"""
    mut = _mutated_cf(tmp_path)
    mp = _mapping_with(tmp_path, {"VC-17": {"enabled": False}})
    out = service.compute_run({"workbook": mut}, "2026-10", MANUAL,
                              mapping_path=mp, log_dir="logs_test")
    assert not [i for i in out.issues if i.rule == "VC-17"]


def test_tolerance_absorbs_small_diff(tmp_path):
    """容差 10,000 元：5,000 元的差异被吸收，不报错。"""
    mut = _mutated_cf(tmp_path)
    mp = _mapping_with(tmp_path, {"VC-17": {"tolerance": 10000}})
    out = service.compute_run({"workbook": mut}, "2026-10", MANUAL,
                              mapping_path=mp, log_dir="logs_test")
    assert not [i for i in out.issues if i.rule == "VC-17" and "期末余额 − 期初余额" in i.title]


def test_accept_issue_unblocks_and_records(tmp_path):
    """界面机制：E 级差异默认阻断；「接受差异」后放行且原因留痕；撤销后重新阻断。"""
    mut = _mutated_cf(tmp_path)
    out = service.compute_run({"workbook": mut}, "2026-10", MANUAL, log_dir="logs_test")
    blocking = out.blocking_issues()
    assert blocking and all(i.rule == "VC-17" for i in blocking)
    out.acknowledged |= out.confirmable_rows()
    ok, reason = out.can_export()
    assert not ok and "校验差异未接受" in reason

    issue = blocking[0]
    out.accept_issue(issue, "本公司口径：审计已确认可接受")
    assert not out.blocking_issues()
    ok, _ = out.can_export()
    assert ok
    assert out.accepted_issues[issue_key(issue)]["reason"] == "本公司口径：审计已确认可接受"

    out.unaccept_issue(issue)
    assert out.blocking_issues() and not out.can_export()[0]


def test_acceptance_dropped_when_issue_gone():
    """重算后差异消失 → 已接受记录自动清除（不残留）。"""
    out = service.compute_run({"workbook": str(SAMPLE)}, "2026-10", MANUAL, log_dir="logs_test")
    fake = type(out.issues[0])(rule="VC-17", level=type(out.issues[0].level).E,
                               title="x", detail="", rows=[]) if out.issues else None
    if fake is not None:
        out.accept_issue(fake, "test")
        out.recompute(MANUAL)
        assert not out.accepted_issues


def test_config_rejects_bad_severity(tmp_path):
    cfg = json.loads(Path("fin_summary/resources/mapping.json").read_text(encoding="utf-8"))
    cfg["validation"]["rules"]["VC-17"]["severity"] = "X"
    p = tmp_path / "bad.json"
    p.write_text(json.dumps(cfg), encoding="utf-8")
    with pytest.raises(MappingError):
        load_mapping(p)
