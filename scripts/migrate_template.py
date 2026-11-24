"""一次性迁移输出模板（V2.6）：Sheet 改名 + C 后插入「取整后(万元)」列。

旧布局：A=项目  B=行次  C=值  D=计算逻辑（Sheet 名「生成的数据」）
新布局：A=项目  B=行次  C=原始值(万元)  D=取整后(万元)  E=计算逻辑（Sheet 名「财务数据计算结果」）

注意：openpyxl 的 insert_cols 不会移动合并区与列宽，需手工处理。
由 scripts/make_template.py 重新生成模板后，需再跑一次本脚本。
"""
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

PATH = "templates/summary_template.xlsx"
OLD_SHEET = "生成的数据"
NEW_SHEET = "财务数据计算结果"
INSERT_AT = 4  # 在 D 列前插入（原 D「计算逻辑」整体移到 E）


def _shift_ref(ref: str) -> str:
    """D58 → E58（仅 >= INSERT_AT 的列右移一列）。"""
    col = "".join(ch for ch in ref if ch.isalpha())
    row = "".join(ch for ch in ref if ch.isdigit())
    ci = column_index_from_string(col)
    if ci >= INSERT_AT:
        ci += 1
    return f"{get_column_letter(ci)}{row}"


def main() -> None:
    wb = load_workbook(PATH)
    if NEW_SHEET in wb.sheetnames and OLD_SHEET not in wb.sheetnames:
        print("已是新版模板，跳过:", NEW_SHEET)
        return
    if OLD_SHEET not in wb.sheetnames:
        raise SystemExit(f"模板中找不到旧 Sheet「{OLD_SHEET}」，请检查 {PATH}")

    ws = wb[OLD_SHEET]

    # 1) 记录并取消合并区（insert_cols 不处理合并区）
    merges = [str(m) for m in ws.merged_cells.ranges]
    for m in merges:
        ws.unmerge_cells(m)

    # 2) 插列：原 D（计算逻辑）→ E
    ws.insert_cols(INSERT_AT)

    # 3) 恢复合并区（列号 >= INSERT_AT 的右移一列）
    for m in merges:
        a, b = m.split(":")
        ws.merge_cells(f"{_shift_ref(a)}:{_shift_ref(b)}")

    # 4) 列宽：E 继承原 D 宽，新 D 取 C 宽
    old_d_width = ws.column_dimensions["D"].width
    ws.column_dimensions["E"].width = old_d_width
    ws.column_dimensions["D"].width = ws.column_dimensions["C"].width

    # 5) 表头默认值（导出时会按月份重写）
    ws.cell(row=1, column=3).value = "原始值(万元)"
    ws.cell(row=1, column=4).value = "取整后(万元)"

    # 6) Sheet 改名
    ws.title = NEW_SHEET

    wb.save(PATH)
    print("迁移完成:", PATH, "| sheets:", wb.sheetnames,
          "| 新表头 C1/D1/E1:", ws["C1"].value, "/", ws["D1"].value, "/", str(ws["E1"].value)[:20])


if __name__ == "__main__":
    main()
