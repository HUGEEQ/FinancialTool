"""2026年9期真实数据回归 + 现金等价物口径机制测试（V2.19）。

根因（2026年9期报错）：现金流量表「现金净增加额 ≠ 期末 − 期初」，差额 -98.78 元
恰为「其他货币资金」期末余额——该科目按本公司口径不计入现金及现金等价物，
货币资金 = 银行存款 + 其他货币资金，故跨表与表内勾稽均存在该口径差。
机制：mapping.json 的 cash_equiv_exclude_accounts 声明这类科目，
差额与科目余额匹配时降为 I 级「已知口径」提示，不阻断导出。
"""
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from fin_summary import service

REPO = Path(__file__).resolve().parents[1]
SEP = REPO / "doc" / "2026年9期.xlsx"
SAMPLE = REPO / "doc" / "财务计算模板.xlsx"
MANUAL = {57: Decimal(1346), 58: Decimal(1327), 59: Decimal(1328), 60: Decimal(1)}


@pytest.fixture(scope="module")
def sep_outcome():
    return service.compute_run({"workbook": str(SEP)}, "2026-09", MANUAL, log_dir="logs_test")


@pytest.mark.skipif(not SEP.exists(), reason="2026年9期样例不存在")
class TestSep2026:
    def test_no_blocking_issues(self, sep_outcome):
        assert not sep_outcome.blocking_issues()

    def test_no_fail_rows(self, sep_outcome):
        assert not [r for r in sep_outcome.results if r.status == type(sep_outcome.results[0].status).FAIL]

    def test_cash_equiv_diff_is_info_level(self, sep_outcome):
        """期末−期初与净增加额的差额（其他货币资金 98.78 元）→ I 级已知口径，不阻断。"""
        hits = [i for i in sep_outcome.issues
                if i.rule == "VC-17" and "其他货币资金" in i.title]
        assert hits and hits[0].level.value == "I"
        assert "98.78" in hits[0].detail

    def test_vc18_caliber_info(self, sep_outcome):
        hits = [i for i in sep_outcome.issues
                if i.rule == "VC-18" and "其他货币资金" in i.title]
        assert hits and hits[0].level.value == "I"

    def test_exportable_after_confirm(self, tmp_path, sep_outcome):
        sep_outcome.acknowledged |= sep_outcome.confirmable_rows()
        ok, reason = sep_outcome.can_export()
        assert ok, reason
        path = service.export_run(sep_outcome, str(tmp_path))
        assert path.exists()


@pytest.mark.skipif(not SAMPLE.exists(), reason="样例模板不存在")
class TestSyntheticCashCaliber:
    """合成场景（CI 可跑）：给样例加「其他货币资金 98.78」并让 CF 净增加额少算 98.78，
    验证工具能识别为已知口径差异（而非 E 级阻断）。"""

    def _build(self, tmp_path) -> str:
        """货币资金内部转账：银行存款 −98.78、其他货币资金 +98.78（总额不变，不影响其他核对）；
        CF 经营流入/经营净额/净增加额同时 +98.78（保持①④勾稽，仅⑤差 98.78）。"""
        wb = load_workbook(SAMPLE)
        tb = wb["科目余额表"]
        moved = False
        for r in range(3, tb.max_row + 1):
            code = str(tb.cell(row=r, column=1).value or "")
            if code == "1002":   # 银行存款（一级科目）
                tb.cell(row=r, column=11).value = (tb.cell(row=r, column=11).value or 0) - 98.78
            elif code == "1009":  # 其他货币资金
                tb.cell(row=r, column=11).value = 98.78
                moved = True
        assert moved, "样例中应存在 1009 其他货币资金科目"
        cf = wb["现金流量表"]
        targets = ("经营活动现金流入小计", "经营活动产生的现金净流量", "现金及现金等价物净增加额")
        for r in range(3, cf.max_row + 1):
            for col in (1, 5):
                s = str(cf.cell(row=r, column=col).value or "")
                if any(t in s for t in targets):
                    cf.cell(row=r, column=col + 2).value = (cf.cell(row=r, column=col + 2).value or 0) + 98.78
        p = tmp_path / "synthetic.xlsx"
        wb.save(p)
        wb.close()
        return str(p)

    def test_known_caliber_detected(self, tmp_path):
        out = service.compute_run({"workbook": self._build(tmp_path)}, "2026-10", MANUAL,
                                  log_dir="logs_test")
        v17 = [i for i in out.issues if i.rule == "VC-17"]
        assert v17 and all(i.level.value == "I" for i in v17), [i.title for i in v17]
        assert not out.blocking_issues()

    def test_disabled_caliber_config_restores_error(self, tmp_path):
        """关闭口径声明后，同一差异重新按 E 级报出（证明识别确实来自该机制）。"""
        import json
        cfg = json.loads((REPO / "fin_summary" / "resources" / "mapping.json").read_text(encoding="utf-8"))
        cfg["cash_equiv_exclude_accounts"] = []
        mp = tmp_path / "mapping_nocal.json"
        mp.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        out = service.compute_run({"workbook": self._build(tmp_path)}, "2026-10", MANUAL,
                                  mapping_path=str(mp), log_dir="logs_test")
        blocking = out.blocking_issues()
        assert blocking and all(i.rule == "VC-17" for i in blocking)
