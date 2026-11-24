"""输入模式测试：文件夹 / 压缩包自动识别 + 模板层级显示名（V2.8）。"""
import shutil
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from fin_summary.exporter import display_names
from fin_summary.loader import ParseError, load_inputs
from fin_summary.ui import wording as W

SAMPLE = Path(__file__).resolve().parents[1] / "doc" / "财务计算模板.xlsx"
SHEETS = ("资产负债表", "利润表", "现金流量表", "科目余额表")
pytestmark = pytest.mark.skipif(not SAMPLE.exists(), reason="样例模板不存在")


def _baseline():
    inputs = load_inputs({"workbook": str(SAMPLE)})
    return inputs.tb.n_rows, inputs.is_table.n_rows


def _make_folder(tmp_path) -> Path:
    """把样例工作簿按字节复制为 4 个以表名命名的文件（保留公式缓存值）。"""
    folder = tmp_path / "报表文件夹"
    folder.mkdir()
    for s in SHEETS:
        shutil.copy(SAMPLE, folder / f"{s}.xlsx")
    return folder


def test_folder_mode_recognizes_by_filename(tmp_path):
    folder = _make_folder(tmp_path)
    inputs = load_inputs({"folder": str(folder)})
    assert inputs.tb.n_rows == _baseline()[0]
    # 名称索引条目一致（n_rows 受样式行影响，不作为数据断言）
    base = load_inputs({"workbook": str(SAMPLE)})
    assert set(inputs.is_table.index) == set(base.is_table.index)


def test_folder_mode_matches_workbook_results(tmp_path):
    """文件夹模式与工作簿模式的 70 行计算结果完全一致。"""
    from fin_summary import service
    folder = _make_folder(tmp_path)
    manual = {57: 1346, 58: 1327, 59: 1328, 60: 1}
    a = service.compute_run({"folder": str(folder)}, "2026-10", manual, log_dir="logs_test")
    b = service.compute_run({"workbook": str(SAMPLE)}, "2026-10", manual, log_dir="logs_test")
    assert [r.rounded for r in a.results] == [r.rounded for r in b.results]


def test_folder_mode_prefers_matching_sheet(tmp_path):
    """文件名含表名时，取该文件内同名 Sheet（而非第一个 Sheet）。"""
    folder = tmp_path / "rev"
    folder.mkdir()
    for s in SHEETS:
        wb = Workbook()
        wb.remove(wb.active)
        wb.create_sheet("无关表")  # 第一个 Sheet 是空表头
        ws = wb.create_sheet(s)
        if s == "科目余额表":
            ws["A1"], ws["B1"] = "科目代码", "科目名称"
            col = 3
            for g in ("年初余额", "期初余额", "本期发生额", "本年累计", "期末余额"):
                ws.cell(row=1, column=col, value=g)
                ws.cell(row=2, column=col, value="借方")
                ws.cell(row=2, column=col + 1, value="贷方")
                col += 2
            ws["A3"], ws["B3"] = "1001", "库存现金"
            wb.save(folder / f"{s}.xlsx")
            wb.close()
            continue
        # 双栏表头（与真实报表列序一致：左栏名称/行次/期间列… 右栏名称/行次/期间列…）
        if s == "资产负债表":
            headers = ["资产", "行次", "期末余额", "年初余额",
                       "负债和所有者权益", "行次", "期末余额", "年初余额"]
        elif s == "利润表":
            headers = ["项目", "行次", "上年实际数", "本月实际数", "本年实际数",
                       "项目", "行次", "上年实际数", "本月实际数", "本年实际数"]
        else:  # 现金流量表
            headers = ["项目", "行次", "本年累计金额", "上年同期金额",
                       "补充资料", "行次", "本年累计金额", "上年同期金额"]
        for i, h in enumerate(headers, start=1):
            ws.cell(row=1, column=i, value=h)
        ws.cell(row=3, column=1, value="某项目")
        ws.cell(row=3, column=2, value=1)
        wb.save(folder / f"{s}.xlsx")
        wb.close()
    inputs = load_inputs({"folder": str(folder)})
    # 每个源文件都正确取到同名 Sheet（若误取「无关表」空表头会触发 VC-24）
    assert inputs.bs_table.layout["value_header"] == "期末余额"
    assert inputs.is_table.layout["value_header"] == "本年实际数"


def test_folder_mode_single_workbook_fallback(tmp_path):
    """文件夹里只有一个含全部 4 张表的文件 → 自动按工作簿模式识别。"""
    folder = tmp_path / "single"
    folder.mkdir()
    src = load_workbook(SAMPLE)
    src.save(folder / "某财务软件导出.xlsx")
    src.close()
    inputs = load_inputs({"folder": str(folder)})
    assert inputs.workbook_path is not None


def test_folder_mode_missing_raises(tmp_path):
    """文件夹里只有资产负债表（且该文件不含其他表）→ VC-02 报缺表。"""
    folder = tmp_path / "partial"
    folder.mkdir()
    src = load_workbook(SAMPLE)
    for s in ("利润表", "现金流量表", "科目余额表", "财务数据计算结果"):
        if s in src.sheetnames:
            del src[s]
    src.save(folder / "资产负债表.xlsx")
    src.close()
    with pytest.raises(ParseError) as ei:
        load_inputs({"folder": str(folder)})
    assert ei.value.rule == "VC-02" and "利润表" in ei.value.message


def test_zip_mode(tmp_path):
    folder = _make_folder(tmp_path)
    zpath = tmp_path / "报表.zip"
    with zipfile.ZipFile(zpath, "w") as z:
        for f in folder.glob("*.xlsx"):
            z.write(f, arcname=f.name)
    inputs = load_inputs({"zip": str(zpath)})
    assert inputs.tb.n_rows == _baseline()[0]


def test_zip_mode_nested_and_bad_file(tmp_path):
    zpath = tmp_path / "bad.zip"
    zpath.write_bytes(b"not a zip")
    with pytest.raises(ParseError) as ei:
        load_inputs({"zip": str(zpath)})
    assert ei.value.rule == "VC-01"


def test_display_names_have_hierarchy():
    names = display_names()
    assert len(names) == 70
    # 层级缩进：顶级项无缩进，其中项有缩进
    assert not names[1].startswith(" ")
    assert names[8].startswith(" ")          # 其中：利息费用
    assert names[55].startswith(" ")         # 其中：在岗职工工资总额
    # 编号与映射名不同（模板带序号）
    assert names[1].startswith("1.")


def test_indent_name_fullwidth():
    assert W.indent_name("营业总收入") == "营业总收入"
    # 层级折算：模板缩进约 4 空格一档，每层 1 个全角空格
    assert W.indent_name("     其中：利息费用") == "　" * 1 + "其中：利息费用"       # 5 空格 → 1 层
    assert W.indent_name("            其中：营业外收入") == "　" * 3 + "其中：营业外收入"  # 12 空格 → 3 层
    deep = W.indent_name("             其中：在岗职工工资总额")            # 13 空格 → 3 层
    assert deep == "　" * 3 + "其中：在岗职工工资总额" and not deep.startswith(" ")
