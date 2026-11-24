"""提取输出模板：保留汇总 Sheet，清除外部链接，另存 templates/summary_template.xlsx。

注意：本脚本从 doc/财务计算模板.xlsx 重新生成的是「旧版结构」模板
（Sheet 名「生成的数据」、单数值列）。生成后必须再运行
scripts/migrate_template.py 完成 V2.6 迁移（Sheet 改名「财务数据计算结果」+
C 后插入「取整后(万元)」列），否则导出将报列结构不符。
"""
from openpyxl import load_workbook

SRC = "doc/财务计算模板.xlsx"
DST = "templates/summary_template.xlsx"

wb = load_workbook(SRC)
wb._external_links = []
for name in ["资产负债表", "利润表", "现金流量表", "科目余额表"]:
    del wb[name]
wb.save(DST)
print("saved:", DST, "sheets:", wb.sheetnames)
print("请接着运行: python scripts/migrate_template.py")
