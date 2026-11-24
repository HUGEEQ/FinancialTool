"""样例数据端到端回归：doc/财务计算模板.xlsx 全 70 行计算 + 校验 + 导出。"""
from decimal import Decimal
from pathlib import Path

import pytest

from fin_summary import service
from fin_summary.config_loader import load_mapping
from fin_summary.engine import compute
from fin_summary.loader import load_inputs
from fin_summary.models import Level, RowStatus
from fin_summary.validator import validate

SAMPLE = Path(__file__).resolve().parents[1] / "doc" / "财务计算模板.xlsx"
MANUAL = {57: Decimal(1346), 58: Decimal(1327), 59: Decimal(1328), 60: Decimal(1)}

pytestmark = pytest.mark.skipif(not SAMPLE.exists(), reason="样例模板不存在")


@pytest.fixture(scope="module")
def outcome():
    return service.compute_run({"workbook": str(SAMPLE)}, "2026-10", MANUAL, log_dir="logs_test")


# ---- 设计文档 14 章样例基准（Q2 口径：先汇总后 ROUND_HALF_UP） ----

BASE = {
    1: 9886, 2: 19253, 3: 74, 4: 0, 5: 2323, 6: 0, 7: 1144, 8: 2118,
    9: 0, 10: 0, 11: -9, 12: 9578, 13: 0, 14: 0, 15: 0, 16: 0,
    17: -3338, 18: 157, 19: 0, 20: -3181, 21: -3181, 22: -3181,
    23: 83594, 24: 33439, 25: 255, 26: 0, 27: 144, 28: 183, 29: 53,
    30: 121, 31: 0, 32: 84115, 33: 21879, 34: 0, 35: 13320,
    36: 58720, 37: 16380, 38: 37000, 39: -521, 40: -521,
    41: 21, 42: 15102, 43: 0, 44: 437, 45: 311, 46: 0, 47: 0,
    48: 514, 49: 371, 50: 0, 51: 0, 52: 11859, 53: 9763, 54: 8609,
    55: 8557, 56: 6, 57: 1346, 58: 1327, 59: 1328, 60: 1,
    61: 749, 62: 4383, 63: 14409, 64: -2270, 65: 437, 66: 4383, 67: 11859,
    68: 0, 69: 0, 70: 0,
}


def test_all_70_rows_computed(outcome):
    assert len(outcome.results) == 70


def test_baseline_values(outcome):
    by_row = {r.row: r for r in outcome.results}
    diff = []
    for row, expect in BASE.items():
        got = by_row[row].rounded
        if got != expect:
            diff.append(f"行{row}: 期望 {expect} 实际 {got}")
    assert not diff, "基准不一致:\n" + "\n".join(diff)


def test_no_fail_rows(outcome):
    fails = [(r.row, r.messages) for r in outcome.results if r.status == RowStatus.FAIL]
    assert not fails, fails


def test_empty_source_rows_pass_with_note(outcome):
    """真空单元格按 0 填列：结果正确 → PASS + I 级提示，不标异常、不阻断导出（V2.4）。"""
    by_row = {r.row: r for r in outcome.results}
    # 样例中真真空单元格：汇兑净损失/资产减值损失/对联营/公允价值/政府补助
    for row in (9, 10, 14, 15, 19):
        r = by_row[row]
        assert r.status == RowStatus.PASS, f"行{row} 不应标记异常"
        assert r.sources and r.sources[0].empty is True
        assert any("单元格为空，按 0 填列" in m for m in r.messages)
    assert not [r for r in outcome.results if r.status == RowStatus.WARN]


def test_source_hit_cell_addresses(outcome):
    """地址贯通（V2.3）：每个来源可定位到具体单元格，钻取/诊断可据此复算。"""
    by_row = {r.row: r for r in outcome.results}
    r9 = by_row[9]
    assert r9.sources[0].cell == "利润表!E24" and r9.sources[0].empty is True
    r11 = by_row[11]
    assert r11.sources[0].cell == "利润表!E26"
    assert r11.sources[0].raw_amount == Decimal("87991.56")
    r8 = by_row[8]  # TB 来源（利息费用）
    assert r8.sources[0].cell.startswith("科目余额表!I")
    for r in outcome.results:
        for s in r.sources:
            if s.type in ("IS", "BS", "CF"):
                assert s.cell and s.cell.startswith(("利润表!", "资产负债表!", "现金流量表!"))


def test_no_e_level_issues(outcome):
    e = [i for i in outcome.issues if i.level == Level.E]
    assert not e, [f"{i.rule} {i.title}" for i in e]


def test_cross_whitelist_note(outcome):
    """V2.16：已知口径提示说人话——净额列示、差额即坏账准备。"""
    hits = [i for i in outcome.issues if i.rule == "VC-19" and i.level == Level.I]
    assert hits
    t = hits[0]
    assert "净额列示" in t.title and "坏账准备" in t.title
    assert "差额" in t.detail and "8.80" in t.detail


def test_manual_pending_blocks_export():
    """CLI 未提供人工数据 → 待填报 → 禁止导出（0 是有效值，但"未提供"仍须拦截）。"""
    out = service.compute_run({"workbook": str(SAMPLE)}, "2026-10", None, log_dir="logs_test")
    pend = {r.row for r in out.results if r.status == RowStatus.PENDING}
    assert pend == {57, 58, 59, 60}
    ok, reason = out.can_export()
    assert not ok and "待填报" in reason


def test_manual_zero_is_valid_value():
    """V2.11：人数填 0 是真实值（如本期无劳务派遣），不报待填报、不阻断导出。"""
    out = service.compute_run({"workbook": str(SAMPLE)}, "2026-10",
                              {57: 0, 58: 0, 59: 0, 60: 0}, log_dir="logs_test")
    by_row = {r.row: r for r in out.results}
    for row in (57, 58, 59, 60):
        assert by_row[row].status == RowStatus.PASS, f"行{row} 不应待填报"
        assert by_row[row].rounded == 0
    assert not [r for r in out.results if r.status == RowStatus.PENDING]
    out.acknowledged |= out.confirmable_rows()
    ok, _ = out.can_export()
    assert ok


def test_manual_partial_missing_still_pending():
    """CLI 只提供部分人数（空槽位）→ 未提供的行仍待填报。"""
    from fin_summary.cli import _parse_manual
    manual = _parse_manual("1346,,1328,")
    out = service.compute_run({"workbook": str(SAMPLE)}, "2026-10", manual, log_dir="logs_test")
    by_row = {r.row: r for r in out.results}
    assert by_row[57].rounded == 1346 and by_row[59].rounded == 1328
    pend = {r.row for r in out.results if r.status == RowStatus.PENDING}
    assert pend == {58, 60}


def test_partial_confirm_blocks_export(outcome):
    """V2.10：未全部人工确认 → 禁止导出。"""
    ok, reason = outcome.can_export()
    assert not ok and "未人工确认" in reason
    outcome.acknowledged |= outcome.confirmable_rows()
    outcome.acknowledged.discard(1)  # 取消一项
    ok, reason = outcome.can_export()
    assert not ok and "尚有 1 项未人工确认" in reason


def test_confirm_all_then_export(tmp_path, outcome):
    """全部人工确认后方可导出。"""
    outcome.acknowledged |= outcome.confirmable_rows()
    done, total = outcome.confirm_progress()
    assert done == total == 70
    ok, _ = outcome.can_export()
    assert ok
    path = service.export_run(outcome, str(tmp_path))
    assert path.exists()

    # V2.6：文件名 财务计算结果_{月份}_{时间戳}.xlsx；Sheet「财务数据计算结果」
    assert path.name.startswith("财务计算结果_2026-10_")
    from openpyxl import load_workbook
    ws = load_workbook(path)["财务数据计算结果"]
    assert ws["C1"].value == "2026-10本年累计 原始值(万元)"
    assert ws["D1"].value == "2026-10本年累计 取整后(万元)"
    by_row = {r.row: r for r in outcome.results}
    for row in range(1, 71):
        v = ws.cell(row=row + 1, column=4).value          # D 取整后
        assert v == by_row[row].rounded, f"D{row + 1}"
        raw = ws.cell(row=row + 1, column=3).value        # C 原始值（全精度）
        if by_row[row].raw is not None:
            assert abs(float(raw) - float(by_row[row].raw)) < 1e-6, f"C{row + 1}"


def test_formula_expression_readable(outcome):
    """V2.6：公式项表达式用项目名而非 #行次 引用。"""
    by_row = {r.row: r for r in outcome.results}
    assert by_row[63].expression == "营业盈余+生产税净额+资产折旧+劳动者报酬"
    assert "营业利润" in by_row[64].expression and "#17" not in by_row[64].expression
    assert "利息费用*0.5" in by_row[64].expression
    # 钻取用的取值明细同样可读
    assert "各项取值" in by_row[63].formula_detail and "营业盈余" in by_row[63].formula_detail


def test_formula_cycle_detected(tmp_path):
    cfg, items, _, _ = load_mapping()
    items[62].formula = "#64+#63"  # 行次63 与 64 构造环（63→64→…→63）
    from fin_summary.engine import compute as C
    from fin_summary.loader import load_inputs
    inputs = load_inputs({"workbook": str(SAMPLE)})
    from fin_summary.loader import ParseError
    with pytest.raises(ParseError) as ei:
        C(items, inputs, cfg, MANUAL)
    assert ei.value.rule == "VC-20"
