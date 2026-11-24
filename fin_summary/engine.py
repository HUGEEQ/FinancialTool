"""计算引擎：graphlib 拓扑排序 + Decimal 精确计算 + 安全公式求值。"""
from __future__ import annotations

import re
from decimal import Decimal, ROUND_HALF_UP
from graphlib import CycleError, TopologicalSorter

from .loader import Inputs, ParseError, cell_ref
from .models import CalcResult, ItemSpec, RowStatus, Source, SourceHit
from .name_match import normalize

DIVISOR = Decimal("10000")
_TOKEN = re.compile(r"\s*(#(\d+)|\d+\.\d+|\d+|ROUND|MAX|ABS|[()+\-*/,])")


class FormulaError(Exception):
    pass


def _round_half_up(x: Decimal) -> Decimal:
    return x.quantize(Decimal("1"), rounding=ROUND_HALF_UP)


# ---------- 安全公式求值（递归下降，仅四则运算 + ROUND/MAX/ABS） ----------

class _Parser:
    def __init__(self, text: str, env: dict[int, Decimal]):
        self.tokens = []
        pos = 0
        while pos < len(text):
            m = _TOKEN.match(text, pos)
            if not m:
                raise FormulaError(f"非法字符: {text[pos:pos+2]!r}")
            self.tokens.append((m.group(1), int(m.group(2)) if m.group(2) else None))
            pos = m.end()
        self.i = 0
        self.env = env

    def peek(self):
        return self.tokens[self.i] if self.i < len(self.tokens) else (None, None)

    def next(self):
        tok = self.peek()
        self.i += 1
        return tok

    def parse(self) -> Decimal:
        v = self.expr()
        if self.peek()[0] is not None:
            raise FormulaError(f"多余 token: {self.peek()[0]!r}")
        return v

    def expr(self) -> Decimal:
        v = self.term()
        while self.peek()[0] in ("+", "-"):
            op, _ = self.next()
            r = self.term()
            v = v + r if op == "+" else v - r
        return v

    def term(self) -> Decimal:
        v = self.factor()
        while self.peek()[0] in ("*", "/"):
            op, _ = self.next()
            r = self.factor()
            if op == "/":
                if r == 0:
                    raise FormulaError("除零")
                v = v / r
            else:
                v = v * r
        return v

    def factor(self) -> Decimal:
        tok, num = self.next()
        if tok == "-":
            return -self.factor()
        if tok == "+":
            return self.factor()
        if tok == "(":
            v = self.expr()
            if self.next()[0] != ")":
                raise FormulaError("缺少右括号")
            return v
        if tok is not None and re.fullmatch(r"\d+(\.\d+)?|\d+", tok):
            return Decimal(tok)
        if tok is not None and tok.startswith("#") and tok[1:].isdigit():
            num = int(tok[1:])
            if num not in self.env:
                raise FormulaError(f"引用行次 {num} 无结果")
            return self.env[num]
        if tok in ("ROUND", "MAX", "ABS"):
            if self.next()[0] != "(":
                raise FormulaError(f"{tok} 缺少左括号")
            args = [self.expr()]
            while self.peek()[0] == ",":
                self.next()
                args.append(self.expr())
            if self.next()[0] != ")":
                raise FormulaError(f"{tok} 缺少右括号")
            return self._call(tok, args)
        raise FormulaError(f"意外 token: {tok!r}")

    def _call(self, fn: str, args: list[Decimal]) -> Decimal:
        if fn == "ROUND":
            digits = int(args[1]) if len(args) > 1 else 0
            q = Decimal(1).scaleb(-digits)
            return args[0].quantize(q, rounding=ROUND_HALF_UP)
        if fn == "ABS":
            return abs(args[0])
        if fn == "MAX":
            return max(args)
        raise FormulaError(f"未知函数 {fn}")


def eval_formula(text: str, env: dict[int, Decimal]) -> Decimal:
    return _Parser(text, env).parse()


# ---------- 引擎主流程 ----------

def _resolve_source(src: Source, inputs: Inputs, config: dict) -> tuple[Decimal, SourceHit, str]:
    """单个来源取数，返回 (converted 万, SourceHit, note)。未命中抛 KeyError。"""
    sheet_names = config["sheets"]
    if src.type == "TB":
        acct = inputs.tb.by_name.get(src.match_name)
        if acct is None:
            raise KeyError(f"科目余额表·{src.match_name}")
        raw = acct.get(src.field) * src.sign
        col = inputs.tb.field_cols.get(src.field)
        cell = f"科目余额表!{cell_ref(acct.row_no, col)}" if col else ""
        hit = SourceHit("TB", f"{acct.code} {acct.name}", src.field, acct.get(src.field), raw / DIVISOR, src.sign, cell=cell)
        return hit.converted, hit, ""
    table = {"IS": inputs.is_table, "BS": inputs.bs_table, "CF": inputs.cf_table}[src.type]
    value, hit_cell, note = table.get(src.match_name, src.aliases, src.match_mode)
    vcol = table.value_column(hit_cell)
    cell = f"{sheet_names[src.type]}!{cell_ref(hit_cell.row, vcol)}"
    raw = value * src.sign
    is_empty = (hit_cell.row, vcol) in table.empty_cells
    hit = SourceHit(src.type, f"{sheet_names[src.type]}·{src.match_name}", "本年实际数/期末余额/本年累计",
                    value, raw / DIVISOR, src.sign, empty=is_empty, cell=cell)
    return hit.converted, hit, note


def _formula_readable(formula: str, done: dict[int, CalcResult]) -> str:
    """公式可读化：#行次 → 规范化项目名（如 #64+#65 → 营业盈余+生产税净额）。"""
    return re.sub(r"#(\d+)", lambda m: normalize(spec_name(done, int(m.group(1)))), formula)


def compute(items: list[ItemSpec], inputs: Inputs, config: dict, manual: dict[int, Decimal]) -> list[CalcResult]:
    """按依赖序计算全部 70 项。manual: 行次 → 人工值（已按原单位提供）。"""
    specs = {it.row: it for it in items}

    # VC-20 循环依赖检测
    sorter = TopologicalSorter({it.row: set(it.deps) for it in items})
    try:
        order = list(sorter.static_order())
    except CycleError as e:
        raise ParseError("VC-20", f"公式循环依赖: {e}") from e

    results: dict[int, CalcResult] = {}
    for row in order:
        spec = specs[row]
        results[row] = _compute_one(spec, inputs, config, manual, results)

    return [results[r] for r in sorted(results)]


def _compute_one(spec: ItemSpec, inputs: Inputs, config: dict,
                 manual: dict[int, Decimal], done: dict[int, CalcResult]) -> CalcResult:
    res = CalcResult(row=spec.row, name=spec.name, status=RowStatus.PASS)

    if spec.kind == "CONST":
        res.raw = spec.const_value
        res.rounded = int(_round_half_up(spec.const_value))
        res.status = RowStatus.CONST
        res.expression = f"按模板固定填 {spec.const_value}"
        return res

    if spec.kind == "MANUAL":
        if spec.row in manual and manual[spec.row] is not None:
            v = manual[spec.row]
            res.raw = v
            res.rounded = int(_round_half_up(v)) if spec.divide else int(v)
            res.expression = "人工填报"
            res.sources = [SourceHit("MANUAL", "人工录入", "人力部门提供", v, v, 1)]
        else:
            res.status = RowStatus.PENDING
            res.expression = "待人工填报"
        return res

    if spec.kind == "FORMULA":
        env = {r: Decimal(done[r].rounded) if done[r].rounded is not None else Decimal("0")
               for r in spec.deps}
        try:
            v = eval_formula(spec.formula, env)
        except FormulaError as e:
            res.status = RowStatus.FAIL
            res.messages.append(f"公式求值失败: {e}")
            return res
        res.raw = v
        res.rounded = int(_round_half_up(v))
        res.expression = _formula_readable(spec.formula, done)
        res.formula_detail = "各项取值（万元）：" + "，".join(
            f"{normalize(spec_name(done, r))} {env[r]:,}" for r in spec.deps)
        return res

    # VALUE / SUM：来源取数
    parts, hits, notes, failed = [], [], [], False
    total = Decimal("0")
    for src in spec.sources:
        try:
            converted, hit, note = _resolve_source(src, inputs, config)
        except KeyError as e:
            if src.missing_ok:
                notes.append(f"报表中无「{e.args[0]}」项目，按 0 填列")
                continue
            res.status = RowStatus.FAIL
            res.messages.append(f"未命中来源: {e.args[0]}（VC-09/10）")
            failed = True
            continue
        except ParseError as e:  # VC-05 非数值
            res.status = RowStatus.FAIL
            res.messages.append(e.message)
            failed = True
            continue
        total += converted
        hits.append(hit)
        sign_s = "×1" if src.sign == 1 else "×-1"
        parts.append(f"{hit.ref}[{src.field or '值'}]{sign_s}={hit.raw_amount}")
        if note and "最左列" in note:
            notes.append(f"{src.match_name}: {note}")

    if failed:
        return res
    res.sources = hits
    if parts:
        res.expression = ("SUM(" + ", ".join(parts) + ")" if spec.kind == "SUM" else parts[0]) + " ÷10000"
    else:
        res.expression = "按 0 填列（来源缺失）"
    res.messages.extend(notes)

    # 下限规则（万单位，舍入前应用）
    if spec.max_floor is not None and total < spec.max_floor:
        res.messages.append(f"按公式算得 {total:,.2f} 万元，低于模板规定的下限 "
                            f"{spec.max_floor:,.0f} 万元，按 {spec.max_floor:,.0f} 万元填列")
        total = spec.max_floor

    res.raw = total
    res.rounded = int(_round_half_up(total))
    # 真空单元格（区别于显式 0）→ I 级提示：空=0 结果正确，不标异常、不阻断导出（V2.4）
    if any(h.empty for h in hits):
        res.messages.append("对应报表单元格为空，按 0 填列")
    return res


def spec_name(done: dict[int, CalcResult], row: int) -> str:
    r = done.get(row)
    return r.name if r else str(row)
