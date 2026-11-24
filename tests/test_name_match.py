"""名称规范化与匹配单元测试（设计文档 5.1 示例）。"""
import pytest

from fin_summary.name_match import NameHit, build_index, match_unique, normalize


@pytest.mark.parametrize("raw,expect", [
    ("　　加：公允价值变动收益（损失以“-”号填列）", "公允价值变动收益"),
    ("三、营业利润（亏损以“－”号填列）", "营业利润"),
    ("\u3000\u3000其中:原材料", "原材料"),
    ("所有者权益（或股东权益）合计", "所有者权益合计"),
    ("归属于母公司所有者权益合计", "归属于母公司所有者权益合计"),
    ("一、营业收入", "营业收入"),
    ("减：营业外支出", "营业外支出"),
    ("库存商品(产成品)", "库存商品"),
    ("*归属于少数股东的综合收益总额", "归属于少数股东的综合收益总额"),
    ("资　　产", "资产"),
])
def test_normalize(raw, expect):
    assert normalize(raw) == expect


def _idx(*texts):
    return build_index([NameHit(row=i + 3, name_col=1 + (i % 2) * 5, text=t)
                        for i, t in enumerate(texts)])


def test_match_exact():
    idx = _idx("一、营业收入", "　　其中：主营业务收入")
    hit, cands, note = match_unique(idx, "营业收入")
    assert hit is not None and hit.row == 3
    assert len(cands) == 1


def test_match_multi_candidates_prefers_left():
    idx = _idx("五、现金及现金等价物净增加额", "现金及现金等价物净增加额")
    hit, cands, note = match_unique(idx, "现金及现金等价物净增加额")
    assert hit.row == 3  # 左栏（name_col=1）优先
    assert len(cands) == 2


def test_match_miss_and_contains():
    idx = _idx("五、现金及现金等价物净增加额")
    hit, _, _ = match_unique(idx, "净增加额")
    assert hit is None  # 默认精确模式不兜底
    hit, _, note = match_unique(idx, "净增加额", mode="contains")
    assert hit is not None and "包含" in note
