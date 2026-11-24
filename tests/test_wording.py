"""wording 文案库测试：规则标题全覆盖 + 数字格式（V2.7 财务人员友好）。"""
import re
from decimal import Decimal
from pathlib import Path

import pytest

from fin_summary import service
from fin_summary.ui import wording as W

SAMPLE = Path(__file__).resolve().parents[1] / "doc" / "财务计算模板.xlsx"
pytestmark = pytest.mark.skipif(not SAMPLE.exists(), reason="样例模板不存在")


def test_rule_titles_cover_sample_issues():
    """样例运行产生的全部规则编号都必须有财务语言标题。"""
    out = service.compute_run({"workbook": str(SAMPLE)}, "2026-10",
                              {57: 1346, 58: 1327, 59: 1328, 60: 1}, log_dir="logs_test")
    rules = {i.rule for i in out.issues}
    missing = rules - set(W.RULE_TITLES)
    assert not missing, f"规则缺少财务语言标题: {missing}"


def test_rule_titles_cover_code_rules():
    """loader/validator/engine 源码中出现的规则编号都必须有标题（防新增规则漏配）。"""
    from fin_summary import engine, loader, validator
    src = "".join(Path(m.__file__).read_text(encoding="utf-8") for m in (loader, validator, engine))
    rules = set(re.findall(r'"(VC-\d+)"', src))
    missing = rules - set(W.RULE_TITLES)
    assert not missing, f"代码中的规则缺少财务语言标题: {missing}"


def test_rule_titles_have_no_codes():
    """标题是给人看的，不能再包含 VC 编号。"""
    for rule, title in W.RULE_TITLES.items():
        assert "VC-" not in title, f"{rule} 标题含编号: {title}"


def test_check_groups_rules_have_titles():
    for _name, rules, _pass in W.CHECK_GROUPS:
        for rule in rules:
            assert rule in W.RULE_TITLES, f"核对分组 {rule} 缺少标题"


def test_fmt_helpers():
    assert W.fmt_wan(Decimal("9885.88063")) == "9,885.88"
    assert W.fmt_wan(None) == "—"
    assert W.fmt_wan_signed(Decimal("15.07")) == "+15.07"
    assert W.fmt_wan_signed(Decimal("-0.01")) == "-0.01"
    assert W.fmt_int(9886) == "9,886"
    assert W.fmt_int(None) == "—"
    assert W.fmt_yuan(Decimal("98858806.30")) == "98,858,806.30"
    assert W.yuan_to_wan(Decimal("98858806.30")) == Decimal("9885.880630")
