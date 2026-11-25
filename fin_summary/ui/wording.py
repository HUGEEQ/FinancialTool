"""界面文案与规则标题（财务人员友好，V2.7）。

集中管理所有面向财务用户的文字与格式。工程细节（VC 编号、run_id、映射版本）
不在此出现——它们只保留在日志与「查看底稿」对话框中。
"""
from __future__ import annotations

from decimal import Decimal

# ---------- 数字格式（千分位） ----------


def fmt_wan(v) -> str:
    """万元，千分位 2 位；None → —。"""
    if v is None:
        return "—"
    return f"{Decimal(str(v)):,.2f}"


def fmt_wan_signed(v) -> str:
    """带符号万元，如 +15.07 / -0.01。"""
    if v is None:
        return "—"
    return f"{Decimal(str(v)):+,.2f}"


def fmt_int(v) -> str:
    if v is None:
        return "—"
    return f"{int(v):,}"


def fmt_yuan(v) -> str:
    """元，千分位 2 位。"""
    if v is None:
        return "—"
    return f"{Decimal(str(v)):,.2f}"


def yuan_to_wan(v) -> Decimal:
    """元 → 万元（Decimal）。"""
    return Decimal(str(v)) / Decimal("10000")


# ---------- 窗口与步骤 ----------

APP_TITLE = "月度财务数据汇总表生成工具"

STEP_INPUT = "第一步　上传报表"
STEP_PERIOD = "第二步　报表期间"
STEP_MANUAL = "第三步　人工数据"
STEP_RESULT = "第四步　计算与核对"
STEP_EXPORT = "第五步　导出"

HELP_INPUT_HTML = (
    "<b>从财务软件导出本月报表，两种方式任选一种：</b><br>"
    "① <b>单个工作簿</b>：一个 Excel 文件含 资产负债表、利润表、现金流量表、科目余额表 4 张表；<br>"
    "② <b>文件夹 / 压缩包</b>：内含 4 个报表文件，按文件名自动识别（文件名含表名即可）。<br><br>"
    "工具按<b>项目名称</b>找数，不要求行号位置固定；文件需为 .xlsx 格式。"
)
HELP_PERIOD_HTML = (
    "<b>选择本月报表期间</b>（如 2026-10，表示 2026 年 10 月）。<br>"
    "生成的报表默认保存到「下载」目录，点「更改…」可换位置。"
)
HELP_MANUAL_HTML = (
    "<b>每月由人力部门提供，单位：人</b>；其余金额自动换算成万元。<br>"
    "0 是有效值（如本期无劳务派遣，直接填 0）。<br>"
    "系统自动带出上期填报的数，本月有变动再改。"
)
HELP_RESULT_HTML = (
    "<b>核对做什么：</b>用会计恒等式交叉验证取数——<br>"
    "　· 资产 = 负债 + 所有者权益<br>"
    "　· 营业利润 = 收入 − 成本 − 费用 ± 减值 ± 收益<br>"
    "　· 科目余额表借方 = 贷方；货币资金 = 现金期末余额 等<br>"
    "全部通过，说明取数与报表自身数字一致；有报错，说明可能串行、漏数或口径不符。<br><br>"
    "<b>公司特殊口径：</b>某条核对按本公司口径可接受时，点问题行的「<b>接受差异</b>」"
    "并填写原因即可放行（记录进底稿，可随时撤销）；不想每次点，也可在 mapping.json 的 "
    "validation 段把该规则降为 W/I 或关闭。<br><br>"
    "<b>行状态：</b>✅ 正常　⚠️ 需确认（导出前请逐项确认）　❗ 错误（导出被禁止）"
    "　⏳ 待填报　⚪ 固定值（按模板规定填 0）<br><br>"
    "<b>取数来源</b>列为 Excel 式引用（如「利润表」!E3），点击任意行可在底部明细看到完整取数过程，"
    "可对照原报表复核。"
)

# ---------- 核对问题列表 ----------

ISSUE_TBL_HEADERS = ("数据核对明细", "操作")
BTN_ACCEPT = "接受差异"
BTN_UNACCEPT = "撤销接受"
BTN_ACK_W = "已知悉"
BTN_ACCEPT_ALL_ISSUES = "全部接受差异"
ACCEPT_DLG_TITLE = "接受差异"
ACCEPT_DLG_PROMPT = "如接受该差异，请说明原因（如：本公司特殊口径，财务已确认）："
ACCEPT_ALL_DLG_PROMPT = "将一次性接受本次运行的全部 {n} 项校验差异，请填写共同原因（逐条记录进底稿）："
TIP_ACCEPT = "公司口径可接受：带原因放行（记录进底稿）"
TIP_UNACCEPT = "撤销放行，重新阻断导出"
TIP_ACK_W = "警告级，不阻断导出"
TIP_ACCEPT_ALL = "本次运行的全部校验差异一次性放行（逐条记录原因，可分别撤销）"

# ---------- 输入区 ----------

MODE_WORKBOOK = "单个工作簿：一个 .xlsx 含 资产负债表 / 利润表 / 现金流量表 / 科目余额表 4 张表"
MODE_FOLDER = "文件夹或压缩包：选择文件夹或 .zip（内含 4 张报表文件，文件名含表名即可自动识别）"
LABEL_WORKBOOK = "工作簿"
LABEL_FOLDER = "位置"
BTN_PICK_FILE = "选择…"
BTN_PICK_FOLDER = "选择文件夹…"
BTN_PICK_ZIP = "选择压缩包…"
ERR_NO_WORKBOOK = "请选择工作簿文件（一个包含 资产负债表 / 利润表 / 现金流量表 / 科目余额表 4 张表的 .xlsx）。"
ERR_NO_FOLDER = "请选择文件夹或 .zip 压缩包（内含 4 张报表文件）。"
ERR_NO_FILES = "请完整选择 4 个输入文件（资产负债表 / 利润表 / 现金流量表 / 科目余额表）。"
ERR_BAD_MONTH = "报表月份格式应为 2026-10 这样的「年-月」。"
SAVE_TO = "保存到："
BTN_CHANGE_DIR = "更改…"

# 输入区折叠（计算区更宽敞）
INPUT_TOGGLE_COLLAPSE = "收起 ▲"
INPUT_TOGGLE_EXPAND = "展开 ▼"
INPUT_COLLAPSED_FMT = "已设置：{inputs}　｜　期间 {month}　｜　{manual}"
INPUT_MANUAL_NONE = "人数未填"
INPUT_MANUAL_FMT = "人数已填 {n} 项"

# ---------- 人工数据 ----------

MANUAL_GROUP_TITLE = "人工数据（每月由人力部门提供，单位：人；其余金额自动换算成万元）"
MANUAL_MAX = 999_999  # 人数上限：百万级（万级已足够，控制输入框宽度）
MANUAL_LABELS = {
    57: "平均职工人数（人）",
    58: "其中：在岗职工（人）",
    59: "平均从业人员（人）",
    60: "其中：劳务派遣（人）",
}
MANUAL_EDIT_LIMIT_MSG = "人数超出合理范围：应在 0 ~ 999999 之间"

# ---------- 按钮与动作 ----------

BTN_RUN = "开始计算"
BTN_EXPORT = "导出 Excel 报表"
BTN_WORKPAPER = "查看底稿"
BTN_LOGS = "打开底稿目录"
RUNNING = "计算中…"
STAGE_PREFIX = "计算中… "
FAIL_TITLE = "无法计算"
DONE_FMT = "计算完成：{pass_n} 项正常，{warn_n} 项需确认，{fail_n} 项错误，{pending_n} 项待填报"
EXPORT_BLOCKED_TIP = "还有条件未满足，按钮旁有说明；把鼠标放按钮上看原因"
EXPORT_OK_TIP = "生成 Excel 报表到「{outdir}」"

# ---------- 人工确认（导出前逐项确认） ----------

CONFIRM_HINT = "人工确认（导出前需逐项确认）："
CONFIRM_PROGRESS_FMT = "已确认 {done} / {total} 项"
BTN_CONFIRM_ALL = "全部确认"
BTN_CONFIRM_NONE = "全部取消"
CONFIRM_ALL_DONE = "已全部确认"
EXPORT_BLOCKED_UNCONFIRMED = "尚有 {n} 项未人工确认"

# ---------- 结果区 ----------

SUMMARY_FMT = ("共 {total} 项：{pass_icon} {pass_n} 项正常　{warn_icon} {warn_n} 项需确认　"
               "{fail_icon} {fail_n} 项错误　{pending_icon} {pending_n} 项待填报　"
               "{const_icon} {const_n} 项固定值")
LEGEND = ("状态说明：✅ 正常　⚠️ 需确认（导出前请逐项确认）　❌ 错误（导出被禁止）"
          "　⏳ 待填报　⚪ 固定值（按模板规定填 0）")
EMPTY_SUMMARY = "尚未计算：上传报表、选好期间后，点「开始计算」。"
FILTER_ABNORMAL = "只看需要关注的行（❌ / ⚠️ / 待填报）"
CHECK_ALL_PASS = "四张报表数据核对全部通过"
DETAIL_PLACEHOLDER = "点击任意一行，这里显示该行的完整取数过程（来源表、项目、单元格、金额换算），可对照原报表复核。"

COL_STATUS = "状态"
COL_NAME = "项目"
COL_RAW = "原始值(万元)"
COL_ROUNDED = "取整后(万元)"
COL_SOURCE = "取数来源"
COL_ACK = "确认无误"

# 数据来源列：报表字段标签
FIELD_BY_TYPE = {"IS": "本年实际数", "BS": "期末余额", "CF": "本年累计"}
SHEET_BY_TYPE = {"IS": "利润表", "BS": "资产负债表", "CF": "现金流量表", "TB": "科目余额表"}
TB_FIELD_LABELS = {
    "年初借": "年初余额借方", "年初贷": "年初余额贷方",
    "期初借": "期初余额借方", "期初贷": "期初余额贷方",
    "本期借": "本期发生额借方", "本期贷": "本期发生额贷方",
    "累计借": "本年累计借方", "累计贷": "本年累计贷方",
    "期末借": "期末余额借方", "期末贷": "期末余额贷方",
}

# 取数来源引用格式（Excel 式，「表」!单元格 + 项目名 + 字段，一眼可定位）
SRC_REPORT_FMT = "「{table}」!{cell} \"{name}\" {field}"
SRC_TB_FMT = "「科目余额表」!{cell} \"{name}\" {field}"
SRC_SUM_JOIN = "  +  "

# 钻取明细
DETAIL_LINE_HEAD = "第 {row} 行　{name}　　状态：{status}"
DETAIL_SRC = "取数：{ref} · {field}（{cell}）"
DETAIL_SRC_EMPTY = "取数：{ref} · {field}（{cell}）——该单元格为空，按 0 填列"
DETAIL_CHAIN = "金额：{yuan} 元 → {wan} 万元 → 取整 {rounded} 万元"
DETAIL_FORMULA = "计算：{expression}"
DETAIL_FORMULA_VALUES = "各项取值（万元）：{values}"
DETAIL_MANUAL = "取数：人工填报（{field}）"
DETAIL_CONST = "取数：按模板固定填 {value}"
DETAIL_PENDING = "取数：待人工填报（单位：人）"

# ---------- 数据核对条（按表分组） ----------

CHECK_GROUPS = [
    # (表名, 规则编号, 全部通过时的描述)
    ("资产负债表", ("VC-13", "VC-14"), "资产 = 负债 + 所有者权益"),
    ("利润表", ("VC-15", "VC-16"), "各项目勾稽关系正常"),
    ("现金流量表", ("VC-17",), "各项目勾稽关系正常"),
    ("科目余额表", ("VC-12", "VC-04"), "借贷平衡"),
    ("跨表核对", ("VC-18", "VC-19"), "账账相符"),
    ("报表结构", ("VC-25",), "行次连续，无串行"),
]
# 汇总行：全部通过 / 有问题（问题逐行列在明细行，不用｜分隔）
CHECK_ALL_PASS = "数据核对：✅ 四张报表全部通过"
CHECK_PASS_SCOPE = ("已核对：资产负债平衡、利润表与现金流量表勾稽关系、科目余额表借贷平衡、"
                    "跨表账账相符、报表行次连续性")

# ---------- 层级缩进显示 ----------


def indent_name(name: str) -> str:
    """模板 A 列名 → 按层级缩进（模板缩进约 4 空格一档，折算为每层 1 个全角空格）。"""
    if not name:
        return name
    stripped = name.lstrip(" ")
    n = len(name) - len(stripped)
    if n <= 0:
        return stripped
    level = max(1, n // 4)
    return "　" * level + stripped


STATUS_TIP_EMPTY = "该行对应的报表单元格为空，按 0 填列（结果正确，仅提示）"
MANUAL_EDIT_TIP = "双击录入人数（单位：人）"
PENDING_VALUE_TIP = "待填报：双击直接输入（单位：人）"
ROUND_CRITICAL_TIP = "该数接近四舍五入边界（小数部分 0.45~0.55），进位方向请人工留意"

# 行级状态图标（emoji，醒目）
STATUS_ICON = {"PASS": "✅", "WARN": "⚠️", "FAIL": "❗", "PENDING": "⏳", "CONST": "⚪"}
STATUS_TEXT = {
    "PASS": "✅ 正常", "WARN": "⚠️ 需确认", "FAIL": "❗ 错误",
    "PENDING": "⏳ 待填报", "CONST": "⚪ 固定值",
}
# 数据核对条标记
CHECK_PASS_MARK = "✅"
CHECK_E_MARK = "❗"
CHECK_W_MARK = "⚠️"

# ---------- 规则编号 → 财务语言（界面显示用；编号保留在日志/底稿） ----------

RULE_TITLES = {
    "VC-01": "文件无法读取",
    "VC-02": "缺少必需的报表",
    "VC-03": "报表没有数据行",
    "VC-04": "科目余额表有重名科目",
    "VC-05": "金额列有非数字内容",
    "VC-08": "计算规则配置有误",
    "VC-09": "报表中找不到对应项目",
    "VC-10": "科目余额表中找不到对应科目",
    "VC-11": "有科目未被使用",
    "VC-12": "科目余额表借贷不平衡",
    "VC-13": "资产负债表不平衡：资产 ≠ 负债 + 所有者权益",
    "VC-14": "资产负债表合计数不符",
    "VC-15": "利润表不平：营业利润与构成项算出的结果不符",
    "VC-16": "利润表不平：利润总额/净利润与构成项不符",
    "VC-17": "现金流量表勾稽关系不符",
    "VC-18": "跨表不符：货币资金与现金期末余额不一致",
    "VC-19": "报表数与科目余额不符",
    "VC-20": "计算公式存在循环依赖",
    "VC-21": "人工数据未填报完整",
    "VC-22": "输出目录不可写",
    "VC-23": "导出文件回读复核不符",
    "VC-24": "报表格式无法识别",
    "VC-25": "报表行次不连续，可能串行",
}

# ---------- 底稿对话框 ----------

WORKPAPER_TITLE = "计算底稿"
WORKPAPER_RUN_ID = "运行编号"
WORKPAPER_MONTH = "报表期间"
WORKPAPER_INPUT = "输入文件"
WORKPAPER_OUTPUT = "输出文件"
WORKPAPER_NOT_EXPORTED = "（尚未导出）"
WORKPAPER_CHECKS = "数据核对"
WORKPAPER_ACKS = "人工确认"
WORKPAPER_ACK_NONE = "无待确认项"
WORKPAPER_LOGS = "底稿文件位置"
WORKPAPER_HINT = "运行编号可用于追溯本次计算；底稿文件供备查，日常无需查看。"
EXPORT_DONE_TITLE = "报表已生成"
EXPORT_DONE_FMT = "已生成：{filename}\n保存位置：{outdir}\n{summary}"
EXPORT_DONE_OPEN = "打开输出目录"
EXPORT_DONE_WORKPAPER = "查看底稿"
EXPORT_DONE_CLOSE = "关闭"
RECOMPUTED_TITLE = "已重新计算"
RECOMPUTED_MSG = "人工数据已更新，受影响的公式项目（增加值等）已自动重算。"
