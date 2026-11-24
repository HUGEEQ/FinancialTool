"""命令行入口：

  python -m fin_summary.cli compute -w <工作簿.xlsx> -m 2026-10 [--manual 1346,1327,1328,1] [--export] [--outdir output]

说明：计算过程与核对结论见终端输出与 logs/ 底稿；导出文件为 Excel 报表
（财务计算结果_报表月份_时间戳.xlsx，含原始值/取整后双列）。
"""
from __future__ import annotations

import argparse
import sys
from decimal import Decimal

from . import service
from .loader import ParseError
from .config_loader import MappingError

MANUAL_ROWS = (57, 58, 59, 60)
MANUAL_NAMES = {57: "平均职工人数", 58: "平均在岗职工人数", 59: "平均从业人员人数", 60: "平均劳务派遣人数"}


def _build_paths(args: argparse.Namespace) -> dict:
    if args.workbook:
        return {"workbook": args.workbook}
    if getattr(args, "folder", None):
        return {"folder": args.folder}
    if getattr(args, "zip", None):
        return {"zip": args.zip}
    return {"资产负债表": args.bs, "利润表": args.is_x, "现金流量表": args.cf, "科目余额表": args.tb}


def _parse_manual(s: str | None) -> dict[int, Decimal]:
    if not s:
        return {}
    parts = [p.strip() for p in s.split(",")]
    if len(parts) != 4:
        raise SystemExit("--manual 需要 4 个值（平均职工人数,在岗,从业,派遣），如 1346,1327,1328,1")
    vals = []
    for p in parts:
        if p == "":
            vals.append(None)
            continue
        try:
            vals.append(Decimal(p))
        except Exception:
            raise SystemExit(f"--manual 含非数值: {p!r}")
    return {row: v for row, v in zip(MANUAL_ROWS, vals) if v is not None}


def _print_results(outcome) -> None:
    icon = {"PASS": "✓", "WARN": "⚠", "FAIL": "✗", "PENDING": "—", "CONST": "○"}
    print(f"\nrun_id: {outcome.run_id}   映射版本: {outcome.mapping_version}")
    print(f"{'行':>4} {'状态':^4} {'原始值(万)':>14} {'取整值':>12}  项目")
    for r in outcome.results:
        raw = f"{r.raw:.2f}" if r.raw is not None else "—"
        val = str(r.rounded) if r.rounded is not None else "—"
        msgs = "；".join(r.messages[:2])
        print(f"{r.row:>4} {icon[r.status.value]:^4} {raw:>14} {val:>12}  {r.name}" + (f"   [{msgs}]" if msgs else ""))
    c = outcome.counts()
    print(f"\n汇总: ✓{c['PASS']}  ⚠{c['WARN']}  ✗{c['FAIL']}  待录入{c['PENDING']}  常量{c['CONST']}")
    for i in outcome.issues:
        print(f"  [{i.level.value}] {i.rule} {i.title}" + (f" —— {i.detail}" if i.detail else ""))
    print(f"计算日志: {outcome.calc_log}")
    print(f"校验日志: {outcome.validation_log}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="fin_summary.cli")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add_common(p):
        p.add_argument("-w", "--workbook", help="包含 4 个 Sheet 的工作簿路径")
        p.add_argument("--folder", help="报表文件夹（内含 4 张报表文件，文件名含表名）")
        p.add_argument("--zip", help="报表压缩包（.zip，内含 4 张报表文件）")
        p.add_argument("--bs", help="资产负债表 xlsx")
        p.add_argument("--is", dest="is_x", help="利润表 xlsx")
        p.add_argument("--cf", help="现金流量表 xlsx")
        p.add_argument("--tb", help="科目余额表 xlsx")
        p.add_argument("-m", "--month", required=True, help="报表月份 YYYY-MM")
        p.add_argument("--mapping", help="mapping.json 路径（默认内置）")
        p.add_argument("--log-dir", default=None, help="日志根目录（默认 ./logs）")

    p1 = sub.add_parser("compute", help="计算并预览（可选导出）")
    add_common(p1)
    p1.add_argument("--manual", help="4 项人工数据，逗号分隔")
    p1.add_argument("--ack-all", action="store_true", help="已逐项人工确认全部计算结果（记录确认）")
    p1.add_argument("--export", action="store_true", help="通过导出条件后写输出文件")
    p1.add_argument("--outdir", default="output")

    args = ap.parse_args(argv)
    paths = _build_paths(args)
    has_single = any(paths.get(k) for k in ("workbook", "folder", "zip"))
    if not has_single and not all(paths.get(k) for k in ("资产负债表", "利润表", "现金流量表", "科目余额表")):
        ap.error("需提供 -w 工作簿、--folder 文件夹、--zip 压缩包，或 --bs/--is/--cf/--tb 四个文件")

    try:
        outcome = service.compute_run(paths, args.month, _parse_manual(getattr(args, "manual", None)),
                                      mapping_path=args.mapping, log_dir=args.log_dir)
    except (ParseError, MappingError) as e:
        print(f"[失败] {e}", file=sys.stderr)
        return 2

    if getattr(args, "ack_all", False):
        from .models import RowStatus
        confirmable = {r.row for r in outcome.results
                       if r.status in (RowStatus.PASS, RowStatus.WARN, RowStatus.CONST)}
        outcome.acknowledged = confirmable

    _print_results(outcome)

    if args.export:
        ok, reason = outcome.can_export()
        if not ok:
            print(f"[未导出] {reason}", file=sys.stderr)
            return 3
        path = service.export_run(outcome, args.outdir)
        print(f"已导出: {path.resolve()}")
        if outcome.audit_summary:
            print(f"审计摘要: {outcome.audit_summary}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
