"""四张输入表解析：利润表/资产负债表/现金流量表（双栏项目定位）+ 科目余额表（层级科目）。

列定位机制（V2.3）：名称列/值列不硬编码行号列号，按第 1 行表头文本动态定位——
报表表：名称列锚定「行次」表头（取其左一列），值列锚定期望表头文本
（利润表=本年实际数、资产负债表=期末余额、现金流量表=本年累计金额）；
科目余额表：按两级表头（第 1 行分组 + 第 2 行借/贷）定位 10 个金额字段列。
定位失败抛 ParseError(VC-24)，消息附实际表头全文，便于人工修正输入表格格式。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook

from .name_match import NameHit, build_index, match_unique, normalize

SHEET_IS = "利润表"
SHEET_BS = "资产负债表"
SHEET_CF = "现金流量表"
SHEET_TB = "科目余额表"
ALL_SHEETS = (SHEET_IS, SHEET_BS, SHEET_CF, SHEET_TB)

# 表头锚点：名称列 = 「行次」表头列 − 1（三张报表第 1 行均有，左右栏各一处）
SEQ_HEADER = "行次"
# 值列锚点：各报表期望的值列表头文本（第 1 行，左右栏各一处）
VALUE_HEADERS = {
    SHEET_IS: "本年实际数",
    SHEET_BS: "期末余额",
    SHEET_CF: "本年累计金额",
}
# 科目余额表两级表头：(第 1 行分组表头, 字段前缀) × (第 2 行子表头, 字段后缀)
TB_GROUPS = (
    ("年初余额", "年初"), ("期初余额", "期初"), ("本期发生额", "本期"),
    ("本年累计", "累计"), ("期末余额", "期末"),
)
TB_SUBS = (("借方", "借"), ("贷方", "贷"))


class ParseError(Exception):
    """结构/数据错误，携带 VC 规则编号与定位（消息为财务语言，便于人工修正输入表格）。"""

    def __init__(self, rule: str, message: str):
        super().__init__(f"[{rule}] {message}")
        self.rule = rule
        self.message = message


def to_decimal(value, where: str) -> Decimal:
    """单元格 → Decimal；空/空白为 0；非数值抛 ParseError(VC-05)。"""
    if value is None:
        return Decimal("0")
    if isinstance(value, int) and not isinstance(value, bool):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    s = str(value).strip().replace(",", "")
    if s == "":
        return Decimal("0")
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ParseError("VC-05", f"{where} 的金额不是数字（{value!r}），请检查该单元格") from None


@dataclass
class ReportTable:
    """一张双栏财务报表的解析结果。"""
    sheet_name: str
    layout: dict
    index: dict[str, list[NameHit]] = field(default_factory=dict)
    values: dict[tuple[int, int], Decimal] = field(default_factory=dict)  # (row,col)->金额
    empty_cells: set[tuple[int, int]] = field(default_factory=set)        # 真空单元格（None/空白）
    seq_numbers: dict[int, list[tuple[int, int | None, str]]] = field(default_factory=dict)  # 名称列→[(行号,行次,原文)]
    n_rows: int = 0

    def value_column(self, hit) -> int:
        """按命中名称所在栏取对应值列（表头定位结果）。"""
        name_cols = self.layout["name_cols"]
        if hit.name_col in name_cols:
            return self.layout["values"][name_cols.index(hit.name_col)]
        return self.layout["values"][0]

    def get(self, match_name: str, aliases=None, mode="normalized_exact"):
        """返回 (value: Decimal, hit: NameHit, note: str)；未命中抛 KeyError。"""
        hit, cands, note = match_unique(self.index, match_name, aliases, mode)
        if hit is None:
            raise KeyError(match_name)
        vcol = self.value_column(hit)
        return self.values.get((hit.row, vcol), Decimal("0")), hit, note


@dataclass
class AccountRow:
    row_no: int
    code: str
    name: str
    amounts: dict[str, Decimal]  # 字段名 → 金额

    def get(self, field_name: str) -> Decimal:
        return self.amounts.get(field_name, Decimal("0"))


@dataclass
class TrialBalance:
    sheet_name: str
    accounts: list[AccountRow] = field(default_factory=list)
    by_name: dict[str, AccountRow] = field(default_factory=dict)
    duplicate_names: list[str] = field(default_factory=list)
    n_rows: int = 0
    field_cols: dict[str, int] = field(default_factory=dict)  # 字段名 → 列号（表头定位结果）

    def top_level(self) -> list[AccountRow]:
        return [a for a in self.accounts if "." not in a.code]

    def top_balance(self, d_field: str, c_field: str) -> Decimal:
        """一级科目借方合计 − 贷方合计。"""
        s = Decimal("0")
        for a in self.top_level():
            s += a.get(d_field) - a.get(c_field)
        return s


def _iter_cells(ws, name_cols):
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max(name_cols)):
        for col in name_cols:
            cell = row[col - 1]
            yield cell


def _header_row(ws, row: int) -> dict[int, str]:
    """指定行非空表头：列号 → 规范化文本。"""
    out = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=row, column=col).value
        if v is None or str(v).strip("\u3000 \t") == "":
            continue
        out[col] = normalize(str(v))
    return out


def _header_dump(ws, row: int) -> str:
    """实际表头全文转储（VC-24 错误消息用，便于人工对照修正表格）。"""
    parts = []
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=row, column=col).value
        if v is None or str(v).strip("\u3000 \t") == "":
            continue
        parts.append(f"{cell_ref(row, col)}={v}")
    return "，".join(parts) if parts else "（空）"


def locate_report_columns(ws, sheet_name: str) -> dict:
    """按第 1 行表头定位双栏结构：名称列 = 「行次」列 − 1；值列 = 期望表头文本所在列。

    返回 {"name_cols": (左, 右), "values": (左, 右), "value_header": 文本}。
    定位失败抛 ParseError(VC-24)，消息附实际表头全文。
    """
    expected = VALUE_HEADERS[sheet_name]
    headers = _header_row(ws, 1)
    seq_cols = [c for c, t in headers.items() if t == SEQ_HEADER]
    value_cols = [c for c, t in headers.items() if t == expected]
    name_cols = [c - 1 for c in seq_cols]

    problems = []
    if len(seq_cols) != 2:
        problems.append(f"「{SEQ_HEADER}」表头找到 {len(seq_cols)} 处（期望 2 处，左右栏各一）")
    if len(value_cols) != 2:
        problems.append(f"「{expected}」表头找到 {len(value_cols)} 处（期望 2 处，左右栏各一）")
    if len(seq_cols) == 2 and len(value_cols) == 2:
        n0, n1 = name_cols
        v0, v1 = value_cols
        if not (n0 < v0 < n1 < v1):
            problems.append("列序不符合「名称列 → 值列 → 名称列 → 值列」的双栏结构")
    if problems:
        raise ParseError(
            "VC-24",
            f"「{sheet_name}」格式无法识别：" + "；".join(problems)
            + f"。第 1 行应为：项目｜行次｜上年实际数｜本月实际数｜{expected}（左右两栏）"
            + f"。实际读到：{_header_dump(ws, 1)}"
            + "。请从财务软件按标准格式重新导出。",
        )
    return {"name_cols": tuple(name_cols), "values": tuple(value_cols), "value_header": expected}


def locate_tb_columns(ws) -> dict[str, int]:
    """科目余额表两级表头定位 10 个金额字段列。失败抛 ParseError(VC-24)。"""
    row1 = _header_row(ws, 1)
    row2 = _header_row(ws, 2)
    fields: dict[str, int] = {}
    problems = []
    for group_text, prefix in TB_GROUPS:
        starts = [c for c, t in row1.items() if t == group_text]
        if len(starts) != 1:
            problems.append(f"第 1 行分组表头「{group_text}」找到 {len(starts)} 处（期望 1 处）")
            continue
        gstart = starts[0]
        later = [c for c in row1 if c > gstart]
        gend = min(later) if later else ws.max_column + 1
        for sub_text, suffix in TB_SUBS:
            cols = [c for c in range(gstart, gend) if row2.get(c) == sub_text]
            if len(cols) != 1:
                problems.append(f"「{group_text}」的「{sub_text}」列找到 {len(cols)} 处（期望 1 处）")
                continue
            fields[prefix + suffix] = cols[0]
    if problems:
        raise ParseError(
            "VC-24",
            "「科目余额表」格式无法识别：" + "；".join(problems)
            + "。第 1 行应为分组表头（年初余额/期初余额/本期发生额/本年累计/期末余额），"
            + f"第 2 行实际表头：{_header_dump(ws, 2)}"
            + "。请从财务软件按标准格式重新导出。",
        )
    return fields


def _collect_seq_numbers(ws, layout: dict) -> dict[int, list[tuple[int, int | None, str]]]:
    """采集各栏「行次」列序列（VC-25 错行探针用）：名称列 → [(行号, 行次或None, 原文)]。"""
    out: dict[int, list[tuple[int, int | None, str]]] = {}
    for name_col in layout["name_cols"]:
        nums = []
        for row in range(3, ws.max_row + 1):
            v = ws.cell(row=row, column=name_col + 1).value
            if v is None or str(v).strip("\u3000 \t") == "":
                continue
            raw = str(v).strip()
            try:
                nums.append((row, int(raw), raw))
            except ValueError:
                nums.append((row, None, raw))
        out[name_col] = nums
    return out


def parse_report(wb, sheet_name: str) -> ReportTable:
    if sheet_name not in wb.sheetnames:
        raise ParseError("VC-02", f"工作簿中找不到「{sheet_name}」这张表")
    ws = wb[sheet_name]
    if ws.max_row < 3:
        raise ParseError("VC-03", f"「{sheet_name}」没有数据行")
    layout = locate_report_columns(ws, sheet_name)

    cells = []
    for cell in _iter_cells(ws, layout["name_cols"]):
        if cell.value is None or cell.row <= 2:
            continue
        text = str(cell.value)
        if text.strip("\u3000 \t") == "":
            continue
        cells.append(NameHit(row=cell.row, name_col=cell.column, text=text))

    table = ReportTable(sheet_name=sheet_name, layout=layout, index=build_index(cells), n_rows=ws.max_row)
    table.seq_numbers = _collect_seq_numbers(ws, layout)
    # 惰性取值：只读用到的单元格（在 get 中按需缓存）
    ws_values = {}

    def value_at(row, col):
        key = (row, col)
        if key not in ws_values:
            raw = ws.cell(row=row, column=col).value
            where = f"{sheet_name}!{cell_ref(row, col)}"
            val = to_decimal(raw, where)
            ws_values[key] = val
            if raw is None or str(raw).strip() == "":
                table.empty_cells.add(key)
        return ws_values[key]

    def get(match_name, aliases=None, mode="normalized_exact"):
        hit, cands, note = match_unique(table.index, match_name, aliases, mode)
        if hit is None:
            raise KeyError(match_name)
        vcol = table.value_column(hit)
        return value_at(hit.row, vcol), hit, note

    table.get = get  # type: ignore[method-assign]
    return table


def cell_ref(row: int, col: int) -> str:
    from openpyxl.utils import get_column_letter
    return f"{get_column_letter(col)}{row}"


def parse_trial_balance(wb) -> TrialBalance:
    if SHEET_TB not in wb.sheetnames:
        raise ParseError("VC-02", f"工作簿中找不到「{SHEET_TB}」这张表")
    ws = wb[SHEET_TB]
    if ws.max_row < 3:
        raise ParseError("VC-03", f"「{SHEET_TB}」没有数据行")

    tb = TrialBalance(sheet_name=SHEET_TB)
    field_cols = locate_tb_columns(ws)
    tb.field_cols = field_cols
    seen_codes: dict[str, int] = {}
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=12):
        code_cell = row[0]
        if code_cell.value is None or str(code_cell.value).strip() == "":
            continue
        code = str(code_cell.value).strip()
        name_cell = row[1].value
        name = "" if name_cell is None else str(name_cell).strip()
        if code in seen_codes:
            raise ParseError("VC-04", f"科目代码重复：{code}（第 {seen_codes[code]} 行与第 {code_cell.row} 行），请检查")
        seen_codes[code] = code_cell.row
        amounts = {}
        for field_name, col in field_cols.items():
            amounts[field_name] = to_decimal(row[col - 1].value, f"{SHEET_TB}!{cell_ref(code_cell.row, col)}")
        acct = AccountRow(row_no=code_cell.row, code=code, name=name, amounts=amounts)
        tb.accounts.append(acct)
        if name and name not in tb.by_name:
            tb.by_name[name] = acct
    tb.n_rows = len(tb.accounts)
    tb.duplicate_names = sorted(
        {a.name for a in tb.accounts if a.name and sum(1 for x in tb.accounts if x.name == a.name) > 1}
    )
    return tb


@dataclass
class Inputs:
    """一次运行解析完成的全部输入。"""
    workbook_path: str | None
    is_table: ReportTable
    bs_table: ReportTable
    cf_table: ReportTable
    tb: TrialBalance
    sha256: str
    duplicate_tb_names: list[str] = field(default_factory=list)


def load_inputs(paths: dict) -> Inputs:
    """paths: {"workbook": path} / {"folder": 目录} / {"zip": 压缩包} / 4 个文件路径。

    文件夹/压缩包模式：按文件名关键词（表名）自动匹配 4 张报表文件；
    单个文件内优先取与表名相同的 Sheet，否则取第一个 Sheet。
    """
    import hashlib

    tmpdir = None
    try:
        if paths.get("zip"):
            tmpdir = _extract_zip(Path(paths["zip"]))
            paths = resolve_paths_from_folder(tmpdir)
        elif paths.get("folder"):
            paths = resolve_paths_from_folder(Path(paths["folder"]))

        if paths.get("workbook"):
            fp = Path(paths["workbook"])
            if not fp.exists():
                raise ParseError("VC-01", f"找不到文件：{fp}")
            if fp.suffix.lower() != ".xlsx":
                raise ParseError("VC-01", f"仅支持 .xlsx（.xls 请先在 Excel 中另存为 .xlsx）：{fp.name}")
            sha = hashlib.sha256(fp.read_bytes()).hexdigest()[:16]
            try:
                wb = load_workbook(fp, read_only=True, data_only=True)
            except Exception as e:  # noqa: BLE001
                raise ParseError("VC-01", f"文件无法打开（若文件含公式请先在 Excel 中打开并保存）: {e}") from e
            is_t = parse_report(wb, SHEET_IS)
            bs_t = parse_report(wb, SHEET_BS)
            cf_t = parse_report(wb, SHEET_CF)
            tb = parse_trial_balance(wb)
            return Inputs(str(fp), is_t, bs_t, cf_t, tb, sha)

        # 四个独立文件（含文件夹/压缩包解析结果）→ 合并为一个工作簿内存对象
        from openpyxl import Workbook
        merged = Workbook()
        merged.remove(merged.active)
        sha_src = b""
        for sheet_name in ALL_SHEETS:
            p = paths.get(sheet_name)
            if not p:
                raise ParseError("VC-02", f"未选择「{sheet_name}」文件")
            fp = Path(p)
            if not fp.exists():
                raise ParseError("VC-01", f"找不到文件：{fp}")
            if fp.suffix.lower() != ".xlsx":
                raise ParseError("VC-01", f"仅支持 .xlsx：{fp.name}")
            sha_src += hashlib.sha256(fp.read_bytes()).digest()
            src = load_workbook(fp, read_only=True, data_only=True)
            # 优先取与表名相同的 Sheet，否则取第一个 Sheet
            s = src[sheet_name] if sheet_name in src.sheetnames else src[src.sheetnames[0]]
            new = merged.create_sheet(title=sheet_name)
            for r in s.iter_rows():
                for c in r:
                    if c.value is not None:
                        new.cell(row=c.row, column=c.column, value=c.value)
            src.close()
        is_t = parse_report(merged, SHEET_IS)
        bs_t = parse_report(merged, SHEET_BS)
        cf_t = parse_report(merged, SHEET_CF)
        tb = parse_trial_balance(merged)
        return Inputs(None, is_t, bs_t, cf_t, tb, hashlib.sha256(sha_src).hexdigest()[:16])
    finally:
        if tmpdir is not None:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)


def resolve_paths_from_folder(folder: Path) -> dict:
    """扫描文件夹，按文件名关键词匹配 4 张报表；单个含全部 4 张表的文件也可识别为工作簿。"""
    files = [p for p in folder.rglob("*")
             if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith(("~$", "."))]
    if not files:
        raise ParseError("VC-02", f"文件夹里没有找到 .xlsx 文件：{folder}")

    found: dict[str, str] = {}
    for sheet in ALL_SHEETS:
        for p in files:
            if sheet in p.name:
                found[sheet] = str(p)
                break
    if len(found) == len(ALL_SHEETS):
        return found

    # 兜底：若某个文件本身包含全部 4 张表，按工作簿模式使用
    for p in files:
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            names = set(wb.sheetnames)
            wb.close()
        except Exception:  # noqa: BLE001
            continue
        if set(ALL_SHEETS) <= names:
            return {"workbook": str(p)}

    missing = [s for s in ALL_SHEETS if s not in found]
    raise ParseError(
        "VC-02",
        f"文件夹里找不到「{'、'.join(missing)}」对应的文件。请确认文件名中包含表名"
        f"（如「资产负债表.xlsx」），或该文件内含全部 4 张表。已找到：{'、'.join(found) or '无'}",
    )


def _extract_zip(zip_path: Path) -> Path:
    """解压 zip 到临时目录（防路径穿越），返回临时目录路径。"""
    import tempfile
    import zipfile

    if not zip_path.exists():
        raise ParseError("VC-01", f"找不到文件：{zip_path}")
    if zip_path.suffix.lower() != ".zip":
        raise ParseError("VC-01", f"压缩包需为 .zip 格式：{zip_path.name}")
    tmp = Path(tempfile.mkdtemp(prefix="finsum_zip_"))
    try:
        with zipfile.ZipFile(zip_path) as z:
            for info in z.infolist():
                target = (tmp / info.filename).resolve()
                if not str(target).startswith(str(tmp.resolve())):
                    raise ParseError("VC-01", f"压缩包含非法路径：{info.filename}")
                if info.is_dir():
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with z.open(info) as srcf, open(target, "wb") as dstf:
                    import shutil
                    shutil.copyfileobj(srcf, dstf)
    except ParseError:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    except Exception as e:  # noqa: BLE001
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
        raise ParseError("VC-01", f"压缩包无法解压：{e}") from e
    return tmp


def layout_snapshot(inputs: Inputs) -> dict:
    """结构快照（写入审计摘要）：各表本次运行实际定位到的名称列/值列/表头。"""
    snap: dict = {}
    for t in (inputs.is_table, inputs.bs_table, inputs.cf_table):
        snap[t.sheet_name] = {
            "name_cols": list(t.layout["name_cols"]),
            "value_cols": list(t.layout["values"]),
            "value_header": t.layout["value_header"],
        }
    snap[SHEET_TB] = {"field_cols": dict(inputs.tb.field_cols)}
    return snap
