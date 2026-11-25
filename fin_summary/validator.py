"""校验体系：表级 VC 规则 + 行级状态挂接（设计文档第 6 章）。

规则可配置（mapping.json 的 validation 段，V2.18）：
  - severity：E（阻断导出）/ W（警告）/ I（提示）——公司特殊口径可将规则降级；
  - enabled=false：关闭该规则；
  - tolerance：容差（元），覆盖默认 0.01 元。
Issue.title 使用财务语言（面向财务人员）；规则编号保留在 Issue.rule，供日志与底稿追溯。
单次差异也可在界面「接受差异」（带原因，留痕审计），不必改配置。
差异金额一律按万元展示（与工具输出口径一致）；需精确到元时，看结果表「原始值(万元)」
列（全精度，点单元格可见完整值）或钻取明细中的元金额。
"""
from __future__ import annotations

from decimal import Decimal

from .loader import Inputs
from .models import CalcResult, Issue, ItemSpec, Level, RowStatus

TOL = Decimal("0.01")
_WAN = Decimal("10000")


def issue_key(issue: Issue) -> str:
    """校验问题唯一键（用于「接受差异」跟踪）。"""
    return f"{issue.rule}|{','.join(map(str, issue.rows))}|{issue.title}"


def _rule_cfg(config: dict) -> dict:
    return ((config or {}).get("validation") or {}).get("rules") or {}


def _default_tol(config: dict) -> Decimal:
    try:
        return Decimal(str(((config or {}).get("validation") or {}).get("tolerance", TOL)))
    except Exception:
        return TOL


def _tol(rules: dict, rule: str, default: Decimal) -> Decimal:
    c = rules.get(rule) or {}
    try:
        return Decimal(str(c.get("tolerance", default)))
    except Exception:
        return default


def _wan(v) -> str:
    """元 → 万元字符串（千分位 2 位）。"""
    return f"{Decimal(v) / _WAN:,.2f}"


def _wan_signed(v) -> str:
    return f"{Decimal(v) / _WAN:+,.2f}"


def _cash_equiv_extra(inputs: Inputs, config: dict) -> tuple[Decimal, Decimal, list[str]]:
    """不计入现金等价物的货币资金科目（如其他货币资金，本公司口径）。

    返回 (期末余额, 期初余额, 科目名列表)。用于识别现金流量表/跨表的已知口径差异。
    """
    names = (config or {}).get("cash_equiv_exclude_accounts") or []
    end = Decimal("0")
    begin = Decimal("0")
    found = []
    for n in names:
        acct = inputs.tb.by_name.get(n)
        if acct is None:
            continue
        end += acct.get("期末借") - acct.get("期末贷")
        begin += acct.get("期初借") - acct.get("期初贷")
        found.append(n)
    return end, begin, found


def validate(inputs: Inputs, results: list[CalcResult], config: dict, items: list[ItemSpec]) -> list[Issue]:
    """执行全部表级校验，返回问题列表（行级挂接见 issue.rows）。"""
    rules = _rule_cfg(config)
    d_tol = _default_tol(config)
    issues = []
    issues += _trial_balance(inputs, rules, d_tol)
    issues += _bs_checks(inputs, rules, d_tol)
    issues += _is_checks(inputs, rules, d_tol)
    issues += _cf_checks(inputs, rules, d_tol, config)
    issues += _seq_checks(inputs)
    issues += _cross_checks(inputs, config, items, rules, d_tol)
    # 应用配置：关闭 / 级别覆盖
    # 语义：severity 仅允许把默认 E 级规则降级为 W/I（放宽公司口径）；W/I 级规则保持默认，不加严。
    out = []
    for i in issues:
        c = rules.get(i.rule) or {}
        if c.get("enabled") is False:
            continue
        sev = str(c.get("severity", "")).upper()
        if sev in ("W", "I") and i.level == Level.E:
            i.level = Level(sev)
        out.append(i)
    _attach_row_status(out, results)
    return out


# ---------- VC-12 试算平衡（按一级科目汇总） ----------

def _trial_balance(inputs: Inputs, rules: dict, d_tol: Decimal) -> list[Issue]:
    issues = []
    tb = inputs.tb
    tol = _tol(rules, "VC-12", d_tol)
    pairs = [("年初", "年初借", "年初贷"), ("期初", "期初借", "期初贷"), ("期末", "期末借", "期末贷")]
    for label, df, cf in pairs:
        diff = tb.top_balance(df, cf)
        if abs(diff) > tol:
            issues.append(Issue("VC-12", Level.E, f"科目余额表借贷不平衡（{label}）",
                                f"借方 − 贷方 = {_wan_signed(diff)} 万元（容差 ±{_wan(tol)} 万元）"))
    if tb.duplicate_names:
        issues.append(Issue("VC-04", Level.W, "科目余额表有重名科目",
                            "、".join(tb.duplicate_names[:10]) + "，请检查是否有科目录重了"))
    return issues


# ---------- VC-13/14 资产负债表勾稽 ----------

def _bs_checks(inputs: Inputs, rules: dict, d_tol: Decimal) -> list[Issue]:
    issues = []
    bs = inputs.bs_table

    def val(name):
        v, _, _ = bs.get(name)
        return v

    try:
        zc, fz, qy = val("资产合计"), val("负债合计"), val("所有者权益合计")
        if abs(zc - fz - qy) > _tol(rules, "VC-13", d_tol):
            issues.append(Issue("VC-13", Level.E, "资产负债表不平衡：资产 ≠ 负债 + 所有者权益",
                                f"资产 {_wan(zc)} 万元 vs 负债+权益 {_wan(fz + qy)} 万元"
                                f"（差 {_wan_signed(zc - fz - qy)} 万元）",
                                rows=[23, 32, 39]))
        ldz, fldz = val("流动资产合计"), val("非流动资产合计")
        if abs(ldz + fldz - zc) > _tol(rules, "VC-14", d_tol):
            issues.append(Issue("VC-14", Level.E, "资产负债表：流动 + 非流动资产的合计 ≠ 资产合计",
                                f"{_wan(ldz)} + {_wan(fldz)} ≠ {_wan(zc)} 万元"
                                f"（差 {_wan_signed(ldz + fldz - zc)} 万元）",
                                rows=[23, 24]))
        flfz, flz = val("流动负债合计"), val("非流动负债合计")
        if abs(flfz + flz - fz) > _tol(rules, "VC-14", d_tol):
            issues.append(Issue("VC-14", Level.E, "资产负债表：流动 + 非流动负债的合计 ≠ 负债合计",
                                f"{_wan(flfz)} + {_wan(flz)} ≠ {_wan(fz)} 万元"
                                f"（差 {_wan_signed(flfz + flz - fz)} 万元）",
                                rows=[32, 33]))
    except KeyError as e:
        issues.append(Issue("VC-13", Level.W, f"资产负债表缺少勾稽所需项目「{e.args[0]}」，跳过该组核对"))
    return issues


# ---------- VC-15/16 利润表勾稽 ----------

def _is_checks(inputs: Inputs, rules: dict, d_tol: Decimal) -> list[Issue]:
    issues = []
    is_t = inputs.is_table

    def val(name):
        v, _, _ = is_t.get(name)
        return v

    def val0(name):
        """可选构成项：缺失按 0 参与勾稽。"""
        try:
            return val(name)
        except KeyError:
            return Decimal("0")

    try:
        rev = val("营业收入")
        cost = val("营业成本")
        calc = (rev - cost
                - val0("营业税金及附加")
                - val0("销售费用") - val0("管理费用") - val0("研究与开发费")
                - val0("财务费用")
                - val0("资产减值损失") - val0("信用减值损失")
                + val0("其他收益") + val0("投资收益")
                + val0("公允价值变动收益") + val0("资产处置收益"))
        op = val("营业利润")
        if abs(calc - op) > _tol(rules, "VC-15", d_tol):
            issues.append(Issue("VC-15", Level.E,
                                "利润表不平：营业利润与「收入−成本−费用±减值±收益」算出的结果不符",
                                f"按构成计算 {_wan(calc)} 万元 vs 报表 {_wan(op)} 万元"
                                f"（差 {_wan_signed(calc - op)} 万元）",
                                rows=[17]))
        op2 = op + val0("营业外收入") - val0("营业外支出") + val0("以前年度损益调整")
        tp = val("利润总额")
        if abs(op2 - tp) > _tol(rules, "VC-16", d_tol):
            issues.append(Issue("VC-16", Level.E, "利润表不平：利润总额 ≠ 营业利润 + 营业外收入 − 营业外支出",
                                f"{_wan(op2)} ≠ {_wan(tp)} 万元（差 {_wan_signed(op2 - tp)} 万元）",
                                rows=[20]))
        np = val("净利润")
        if abs(tp - val0("所得税费用") - np) > _tol(rules, "VC-16", d_tol):
            issues.append(Issue("VC-16", Level.E, "利润表不平：净利润 ≠ 利润总额 − 所得税费用",
                                f"{_wan(tp)} − 所得税 ≠ {_wan(np)} 万元"
                                f"（差 {_wan_signed(tp - val0('所得税费用') - np)} 万元）",
                                rows=[21]))
    except KeyError as e:
        issues.append(Issue("VC-15", Level.W, f"利润表缺少勾稽所需核心项目「{e.args[0]}」，跳过该组核对"))
    return issues


# ---------- VC-17 现金流量表勾稽 ----------

def _cf_checks(inputs: Inputs, rules: dict, d_tol: Decimal, config: dict) -> list[Issue]:
    issues = []
    cf = inputs.cf_table
    tol = _tol(rules, "VC-17", d_tol)

    def val(name):
        v, _, _ = cf.get(name)
        return v

    try:
        net_op = val("经营活动产生的现金净流量")
        if abs(val("经营活动现金流入小计") - val("经营活动现金流出小计") - net_op) > tol:
            issues.append(Issue("VC-17", Level.E, "现金流量表：经营活动净额 ≠ 流入小计 − 流出小计",
                                rows=[41]))
        net_inv = val("投资活动产生的现金流量净额")
        if abs(val("投资活动现金收入小计") - val("投资活动现金流出小计") - net_inv) > tol:
            issues.append(Issue("VC-17", Level.E, "现金流量表：投资活动净额 ≠ 流入小计 − 流出小计"))
        net_fin = val("筹资活动产生的现金流量净额")
        if abs(val("筹资活动现金流入小计") - val("筹资活动现金流出小计") - net_fin) > tol:
            issues.append(Issue("VC-17", Level.E, "现金流量表：筹资活动净额 ≠ 流入小计 − 流出小计"))
        fx = val("汇率变动对现金及现金等价物的影响")
        inc_total = val("现金及现金等价物净增加额")
        if abs(net_op + net_inv + net_fin + fx - inc_total) > tol:
            issues.append(Issue("VC-17", Level.E,
                                "现金流量表：现金净增加额 ≠ 三大活动净额 + 汇率影响",
                                f"差 {_wan_signed(net_op + net_inv + net_fin + fx - inc_total)} 万元",
                                rows=[41]))
        end_cash, begin_cash = val("现金的期末余额"), val("现金的期初余额")
        diff = end_cash - begin_cash - inc_total
        if abs(diff) > tol:
            # 已知口径：差额若恰为「不计入现金等价物」的科目余额变动（如其他货币资金），
            # 属正常口径差异（现金等价物不含该科目），降为 I 级提示，不阻断。
            extra_end, extra_begin, names = _cash_equiv_extra(inputs, config)
            cands = [extra_end, extra_begin, extra_end - extra_begin]
            if names and any(abs(abs(diff) - abs(c)) <= tol for c in cands):
                issues.append(Issue(
                    "VC-17", Level.I,
                    f"现金流量表：净增加额与期末−期初的差额为「{'、'.join(names)}」（不计入现金等价物，属正常口径）",
                    f"期末 {_wan(end_cash)} − 期初 {_wan(begin_cash)} ≠ 净增加额 {_wan(inc_total)} 万元，"
                    f"差 {_wan_signed(diff)} 万元；恰为「{'、'.join(names)}」余额"
                    f"（期末 {extra_end} 元 / 期初 {extra_begin} 元）。该科目按本公司口径不计入现金及现金等价物，"
                    "不影响取数。",
                    rows=[41]))
            else:
                issues.append(Issue("VC-17", Level.E, "现金流量表：现金净增加额 ≠ 期末余额 − 期初余额",
                                    f"期末 {_wan(end_cash)} − 期初 {_wan(begin_cash)} ≠ 净增加额 {_wan(inc_total)} 万元"
                                    f"（差 {_wan_signed(diff)} 万元）",
                                    rows=[41]))
    except KeyError as e:
        issues.append(Issue("VC-17", Level.W, f"现金流量表缺少勾稽所需项目「{e.args[0]}」，跳过该组核对"))
    return issues


# ---------- VC-25 行次连续性（错行/重复行探针） ----------

def _seq_checks(inputs: Inputs) -> list[Issue]:
    """各栏「行次」列应严格递增；逆序/重复 → W（数据可能错行）。"""
    issues = []
    for t in (inputs.is_table, inputs.bs_table, inputs.cf_table):
        cols = sorted(t.seq_numbers)
        for name_col, nums in t.seq_numbers.items():
            side = "左栏" if name_col == cols[0] else "右栏"
            prev_row = prev_val = None
            for row, val, raw in nums:
                if val is None:
                    issues.append(Issue("VC-25", Level.W,
                                        f"{t.sheet_name}{side}第 {row} 行行次不是数字（{raw}）",
                                        "请检查该行是否错位"))
                    continue
                if prev_val is not None and val <= prev_val:
                    issues.append(Issue("VC-25", Level.W,
                                        f"{t.sheet_name}{side}行次不连续：第 {prev_row} 行 {prev_val} → 第 {row} 行 {val}",
                                        "行次应一行一档、逐行递增；请检查报表是否串行或重复粘贴"))
                prev_row, prev_val = row, val
    return issues


# ---------- VC-18/19 跨表勾稽（含口径白名单） ----------

def _cross_checks(inputs: Inputs, config: dict, items: list[ItemSpec],
                  rules: dict, d_tol: Decimal) -> list[Issue]:
    issues = []
    # BS 项目名 → 输出行次（用于 W 级挂接）
    bs_item_rows: dict[str, int] = {}
    for it in items:
        for s in it.sources:
            if s.type == "BS":
                bs_item_rows.setdefault(s.match_name, it.row)

    # VC-18 货币资金 vs 现金期末余额
    try:
        money, _, _ = inputs.bs_table.get("货币资金")
        end_cash, _, _ = inputs.cf_table.get("现金的期末余额")
        if abs(money - end_cash) > _tol(rules, "VC-18", d_tol):
            # 已知口径：货币资金含不计入现金等价物的科目（如其他货币资金）
            extra_end, _extra_begin, names = _cash_equiv_extra(inputs, config)
            if names and abs(abs(money - end_cash) - abs(extra_end)) <= _tol(rules, "VC-18", d_tol):
                issues.append(Issue(
                    "VC-18", Level.I,
                    f"跨表差异为不计入现金等价物的「{'、'.join(names)}」（属正常口径）",
                    f"货币资金 {_wan(money)} 万元 vs 现金期末 {_wan(end_cash)} 万元，"
                    f"差 {_wan_signed(money - end_cash)} 万元；恰为「{'、'.join(names)}」期末余额 "
                    f"{extra_end} 元（该科目不计入现金及现金等价物）。",
                    rows=[25]))
            else:
                issues.append(Issue("VC-18", Level.W,
                                    "跨表不符：资产负债表「货币资金」≠ 现金流量表「现金的期末余额」",
                                    f"{_wan(money)} vs {_wan(end_cash)} 万元（差 {_wan_signed(money - end_cash)} 万元）",
                                    rows=[25]))
    except KeyError as e:
        issues.append(Issue("VC-18", Level.W, f"跨表核对缺少项目「{e.args[0]}」"))

    # VC-19 科目余额表 vs 报表
    bs = inputs.bs_table
    tol19 = _tol(rules, "VC-19", d_tol)
    for cc in config.get("cross_checks", []):
        item = str(cc["item"])
        total = Decimal("0")
        missing = []
        for acct_name in cc["tb_accounts"]:
            acct = inputs.tb.by_name.get(acct_name)
            if acct is None:
                missing.append(acct_name)
                continue
            total += acct.get(str(cc["tb_field"]))
        if missing:
            issues.append(Issue("VC-19", Level.I, f"跨表抽查「{item}」科目余额表缺少科目: {'、'.join(missing)}"))
            continue
        try:
            report_v, _, _ = bs.get(item)
        except KeyError:
            issues.append(Issue("VC-19", Level.I, f"跨表抽查「{item}」报表无此项目"))
            continue
        diff = report_v - total
        if abs(diff) > tol19:
            detail = (f"资产负债表「{item}」{_wan(report_v)} 万元（净额）；科目余额 {_wan(total)} 万元（全额）；"
                      f"差额 {abs(diff) / _WAN:,.2f} 万元。{cc.get('note', '')}")
            if cc.get("known_diff"):
                issues.append(Issue("VC-19", Level.I,
                                    f"「{item}」按扣除坏账准备后的净额列示，与科目余额的差额即坏账准备，属正常",
                                    detail,
                                    rows=[bs_item_rows[item]] if item in bs_item_rows else []))
            else:
                issues.append(Issue("VC-19", Level.W, f"「{item}」与科目余额不符", detail,
                                    rows=[bs_item_rows[item]] if item in bs_item_rows else []))
    return issues


# ---------- 行级状态挂接 ----------

def _attach_row_status(issues: list[Issue], results: list[CalcResult]) -> None:
    """表级校验问题降级挂接到对应行：E→行 FAIL，W→行 WARN。"""
    if not issues:
        return
    by_row = {r.row: r for r in results}
    for issue in issues:
        if not issue.rows:
            continue
        for row in issue.rows:
            res = by_row.get(row)
            if res is None or res.status == RowStatus.FAIL:
                continue
            if issue.level == Level.E:
                res.status = RowStatus.FAIL
                res.messages.append(issue.title)
            elif issue.level == Level.W and res.status != RowStatus.WARN:
                if res.status == RowStatus.PASS:
                    res.status = RowStatus.WARN
                    res.messages.append(issue.title)
