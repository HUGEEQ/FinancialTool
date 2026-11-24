"""输出生成：全新生成「财务数据计算结果」工作表（无颜色、无计算逻辑列）+ 回读复核（VC-23）。

输出布局（V2.14）：
    A=项目(含层级缩进)  B=行次  C=原始值(万元，全精度)  D=取整后(万元)
文件名：财务计算结果_{报表月份}_{时间戳}.xlsx。
项目名称与层级缩进取自输出模板 A 列（templates/summary_template.xlsx），其余样式一律不带
（无填充色、无边框、默认字体），保证输出干净、可直接用于上报。
"""
from __future__ import annotations

from decimal import Decimal
from functools import lru_cache
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .models import CalcResult

SHEET_NAME = "财务数据计算结果"
RAW_COL = 3      # C 原始值(万元)
ROUNDED_COL = 4  # D 取整后(万元)
N_ROWS = 70
TOL = 1e-6       # C 列（Decimal→float）回读容差


def template_path() -> Path:
    """输出模板路径（仅用作项目名称与层级缩进的来源）。"""
    cand = Path("templates/summary_template.xlsx")
    return cand if cand.exists() else Path(__file__).parent / "resources" / "summary_template.xlsx"


@lru_cache(maxsize=1)
def display_names() -> dict[int, str]:
    """输出模板 A 列项目名（含层级缩进与编号）→ {行次: 显示名}。

    界面与输出文件共用同一套名称，保证层级展示一致；模板缺失时返回空（调用方回退到映射名）。
    """
    try:
        ws = load_workbook(template_path(), read_only=True)[SHEET_NAME]
        out = {}
        for r in range(2, N_ROWS + 2):
            row_no = ws.cell(row=r, column=2).value
            name = ws.cell(row=r, column=1).value
            if row_no is not None and name:
                out[int(row_no)] = str(name)
        return out
    except Exception:
        return {}


def _output_path(outdir: Path, month: str, run_id: str) -> Path:
    """财务计算结果_{月份}_{时间戳}.xlsx；同秒重名时退化用完整 run_id 防覆盖。"""
    stamp = "_".join(run_id.split("_")[:2])  # YYYYMMDD_HHMMSS
    final = outdir / f"财务计算结果_{month}_{stamp}.xlsx"
    if final.exists():
        final = outdir / f"财务计算结果_{month}_{run_id}.xlsx"
    return final


def write_output(results: list[CalcResult], month: str, outdir: Path, run_id: str) -> Path:
    """全新生成工作簿 → 写临时文件 → 回读复核 → 重命名正式文件。失败清理临时文件。

    不复制模板：无颜色、无计算逻辑列，仅 项目/行次/原始值/取整后 四列。
    """
    outdir.mkdir(parents=True, exist_ok=True)
    final = _output_path(outdir, month, run_id)
    tmp = outdir / f".tmp_{run_id}.xlsx"

    hdr_raw = f"{month}本年累计 原始值(万元)"
    hdr_rounded = f"{month}本年累计 取整后(万元)"
    names = display_names()

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.cell(row=1, column=1, value="项目")
    ws.cell(row=1, column=2, value="行次")
    ws.cell(row=1, column=3, value=hdr_raw)
    ws.cell(row=1, column=4, value=hdr_rounded)

    by_row = {r.row: r for r in results}
    for row in range(1, N_ROWS + 1):
        res = by_row.get(row)
        name = names.get(row) or (res.name if res else "")
        ws.cell(row=row + 1, column=1, value=name)
        ws.cell(row=row + 1, column=2, value=row)
        if res is not None and res.raw is not None:
            c_raw = ws.cell(row=row + 1, column=RAW_COL)
            c_raw.value = res.raw if isinstance(res.raw, Decimal) else Decimal(str(res.raw))
            c_raw.number_format = "#,##0.00"  # 显示 2 位；单元格存全精度，可换算回元
        if res is not None and res.rounded is not None:
            ws.cell(row=row + 1, column=ROUNDED_COL).value = int(res.rounded)
    # 列宽与冻结表头（无任何颜色样式）
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.freeze_panes = "A2"
    wb.save(tmp)

    # 回读复核（VC-23）：D 列严格比对，C 列容差比对
    try:
        wb2 = load_workbook(tmp, data_only=False)
        ws2 = wb2[SHEET_NAME]
        bad = []
        for row in range(1, N_ROWS + 1):
            res = by_row.get(row)
            if res is None or res.rounded is None:
                continue
            actual_d = ws2.cell(row=row + 1, column=ROUNDED_COL).value
            if actual_d is None or int(actual_d) != int(res.rounded):
                bad.append(f"D{row + 1}: 期望 {res.rounded} 实际 {actual_d}")
            if res.raw is not None:
                actual_c = ws2.cell(row=row + 1, column=RAW_COL).value
                if actual_c is None or abs(float(actual_c) - float(res.raw)) > TOL:
                    bad.append(f"C{row + 1}: 期望 {res.raw} 实际 {actual_c}")
        if bad:
            raise ExportError("回读复核失败（VC-23）: " + "; ".join(bad[:5]))
        if ws2.cell(row=1, column=RAW_COL).value != hdr_raw or ws2.cell(row=1, column=ROUNDED_COL).value != hdr_rounded:
            raise ExportError("回读复核失败：C1/D1 表头不符")
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    tmp.replace(final)
    return final
