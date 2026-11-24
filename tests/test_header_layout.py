"""表头驱动列定位 + VC-24/VC-25 结构校验测试（设计文档 V2.3）。"""
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from fin_summary.config_loader import load_mapping
from fin_summary.engine import compute
from fin_summary.loader import ParseError, layout_snapshot, load_inputs, locate_report_columns, locate_tb_columns
from fin_summary.validator import validate

SAMPLE = Path(__file__).resolve().parents[1] / "doc" / "财务计算模板.xlsx"
pytestmark = pytest.mark.skipif(not SAMPLE.exists(), reason="样例模板不存在")


@pytest.fixture(scope="module")
def wb():
    book = load_workbook(SAMPLE, read_only=True, data_only=True)
    yield book
    book.close()


def test_report_columns_located_by_header(wb):
    assert locate_report_columns(wb["利润表"], "利润表") == {
        "name_cols": (1, 6), "values": (5, 10), "value_header": "本年实际数"}
    assert locate_report_columns(wb["资产负债表"], "资产负债表") == {
        "name_cols": (1, 5), "values": (3, 7), "value_header": "期末余额"}
    assert locate_report_columns(wb["现金流量表"], "现金流量表") == {
        "name_cols": (1, 5), "values": (3, 7), "value_header": "本年累计金额"}


def test_tb_columns_located_by_two_level_header(wb):
    assert locate_tb_columns(wb["科目余额表"]) == {
        "年初借": 3, "年初贷": 4, "期初借": 5, "期初贷": 6,
        "本期借": 7, "本期贷": 8, "累计借": 9, "累计贷": 10,
        "期末借": 11, "期末贷": 12,
    }


def test_layout_snapshot_recorded():
    inputs = load_inputs({"workbook": str(SAMPLE)})
    snap = layout_snapshot(inputs)
    assert snap["利润表"]["value_cols"] == [5, 10]
    assert snap["科目余额表"]["field_cols"]["累计借"] == 9


def _mutated(tmp_path, mutate) -> str:
    """复制样例工作簿并施加变异，返回变异文件路径。"""
    src = load_workbook(SAMPLE)
    try:
        mutate(src)
        p = tmp_path / "mutated.xlsx"
        src.save(p)
        return str(p)
    finally:
        src.close()


def test_value_header_renamed_raises_vc24(tmp_path):
    """右栏值表头文本被改 → VC-24 阻断，且消息含实际表头便于人工修正。"""
    p = _mutated(tmp_path, lambda wb: wb["利润表"].__setitem__("J1", "本年数"))
    with pytest.raises(ParseError) as ei:
        load_inputs({"workbook": p})
    assert ei.value.rule == "VC-24"
    msg = ei.value.message
    assert "利润表" in msg and "本年实际数" in msg and "J1=本年数" in msg


def test_seq_header_missing_raises_vc24(tmp_path):
    """删掉「行次」表头 → VC-24（名称列锚点失效）。"""
    p = _mutated(tmp_path, lambda wb: wb["资产负债表"].__setitem__("F1", None))
    with pytest.raises(ParseError) as ei:
        load_inputs({"workbook": p})
    assert ei.value.rule == "VC-24" and "行次" in ei.value.message


def test_tb_group_header_broken_raises_vc24(tmp_path):
    """科目余额表分组表头被改 → VC-24，消息含两级实际表头。"""
    p = _mutated(tmp_path, lambda wb: wb["科目余额表"].__setitem__("I1", "累计"))
    with pytest.raises(ParseError) as ei:
        load_inputs({"workbook": p})
    assert ei.value.rule == "VC-24"
    assert "科目余额表" in ei.value.message and "第 2 行实际表头" in ei.value.message


def test_column_shift_relocates_instead_of_reading_wrong(tmp_path):
    """删除利润表 D 列（本月实际数）→ 值列整体左移并被正确定位，而非静默读错列。"""
    def m(wb):
        wb["利润表"].delete_cols(4)  # E→D、J→I，行次列同步左移
    p = _mutated(tmp_path, m)
    inputs = load_inputs({"workbook": p})
    assert inputs.is_table.layout["name_cols"] == (1, 5)
    assert inputs.is_table.layout["values"] == (4, 9)


def _issues_for(path: str):
    inputs = load_inputs({"workbook": path})
    cfg, items, _, _ = load_mapping()
    return validate(inputs, compute(items, inputs, cfg, {}), cfg, items)


def test_seq_continuity_clean_on_sample():
    assert not [i for i in _issues_for(str(SAMPLE)) if i.rule == "VC-25"]


def test_seq_reversed_raises_vc25(tmp_path):
    """利润表行次 21 被改为 1（逆序）→ VC-25 W 级提示可能错行。"""
    p = _mutated(tmp_path, lambda wb: wb["利润表"].__setitem__("B24", 1))
    vc25 = [i for i in _issues_for(p) if i.rule == "VC-25"]
    assert vc25 and vc25[0].level.value == "W"
    assert "利润表" in vc25[0].title and "不连续" in vc25[0].title


def test_seq_duplicated_raises_vc25(tmp_path):
    """行次重复（B24 改为 20，与 B23 相同）→ VC-25。"""
    p = _mutated(tmp_path, lambda wb: wb["利润表"].__setitem__("B24", 20))
    vc25 = [i for i in _issues_for(p) if i.rule == "VC-25"]
    assert vc25 and "20 → 第 24 行 20" in vc25[0].title
