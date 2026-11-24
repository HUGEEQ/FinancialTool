"""主界面：左侧输入设置 + 右侧「计算与核对」核心区（设计文档 9.1，V2.10 布局重构）。

布局原则：计算与核对是核心功能，占据界面主要位置（右侧可拉伸区域）；
输入设置等辅助功能收在左侧窄栏（分隔条可拖拽调宽窄）。
导出前须对每一项人工确认（全部确认方可导出）。
"""
from __future__ import annotations

import json
import webbrowser
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QGridLayout, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox,
    QPlainTextEdit, QPushButton, QRadioButton, QSpinBox, QSplitter, QStatusBar,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from .. import service
from ..exporter import display_names
from ..models import Level, RowStatus
from . import wording as W
from .worker import ComputeThread

MONTHS_BACK = 24
FILE_SHEETS = ("资产负债表", "利润表", "现金流量表", "科目余额表")
STATE_PATH = Path.home() / ".fin_summary_state.json"
LEFT_PANEL_MAX_W = 430

COLOR = {
    RowStatus.PASS: QColor("#f0fff0"),
    RowStatus.WARN: QColor("#fff3e0"),
    RowStatus.FAIL: QColor("#fdecea"),
    RowStatus.PENDING: QColor("#f5f5f5"),
    RowStatus.CONST: QColor("#f7fcff"),
}
CRITICAL_BG = QColor("#fff9c4")  # 舍入临界高亮


def _default_outdir() -> str:
    """输出目录默认取本机下载目录，不存在时退到 ./output。"""
    downloads = Path.home() / "Downloads"
    return str(downloads if downloads.is_dir() else Path("output").resolve())


def _load_last_manual() -> dict[int, int]:
    """读取上期填报的人数（本地记忆，本月有变动再改；0 也如实恢复）。"""
    try:
        data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        return {int(k): int(v) for k, v in (data.get("last_manual") or {}).items()}
    except Exception:
        return {}


def _save_last_manual(manual: dict[int, Decimal]) -> None:
    try:
        STATE_PATH.write_text(
            json.dumps({"last_manual": {str(k): int(v) for k, v in manual.items()}}, ensure_ascii=False),
            encoding="utf-8")
    except Exception:
        pass


def _source_brief(r) -> str:
    """取数来源列：Excel 式引用「表」!单元格 "项目" 字段，一眼可在源表中定位。

    报表类：「利润表」!E3 "营业收入" 本年实际数
    科目类：「科目余额表」!L904 "长期借款_银行贷款" 期末余额贷方
    多来源用「 + 」连接；空单元格标注（空）。
    """
    if not r.sources:
        return r.expression  # 公式可读化 / 按模板固定填 0 / 待人工填报
    parts = []
    for s in r.sources:
        if s.type == "MANUAL":
            parts.append('人工填报"人数"（人力部门提供，单位：人）')
            continue
        if s.type == "TB":
            name = s.ref.split(" ", 1)[1] if " " in s.ref else s.ref  # 去掉科目代码前缀
            field = W.TB_FIELD_LABELS.get(s.field, s.field)
            tpl = W.SRC_TB_FMT
        else:
            name = s.ref.split("·", 1)[1] if "·" in s.ref else s.ref
            field = W.FIELD_BY_TYPE.get(s.type, s.field)
            tpl = W.SRC_REPORT_FMT.replace("{table}", W.SHEET_BY_TYPE.get(s.type, s.type))
        _, _, addr = s.cell.partition("!") if s.cell else (None, None, "")
        txt = (tpl.replace("{cell}", addr)
                  .replace("{name}", name)
                  .replace("{field}", field))
        if s.empty:
            txt += "（空）"
        parts.append(txt)
    return W.SRC_SUM_JOIN.join(parts) if len(parts) > 1 else parts[0]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(W.APP_TITLE)
        self.resize(1360, 940)
        # 全局字号：紧凑可读，避免左栏文字截断
        app = QApplication.instance()
        if app is not None:
            f = app.font()
            f.setPixelSize(12)
            app.setFont(f)
            # 悬停提示浮框样式：清晰排版、足够停留时间
            app.setStyleSheet(
                "QToolTip { background-color: #fffdea; color: #333333;"
                " border: 1px solid #bbbbbb; padding: 6px 8px; font-size: 12px; }")
        self.outcome = None
        self._thread = None
        self._outdir = _default_outdir()
        self._build_ui()
        self._refresh()

    # ---------- UI 骨架：左输入 / 右核心 ----------

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        h = QHBoxLayout(central)
        h.setContentsMargins(8, 8, 8, 8)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([LEFT_PANEL_MAX_W, 900])
        h.addWidget(splitter)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("就绪")

    def _build_left_panel(self) -> QWidget:
        """左侧窄栏：输入设置（上传/期间/人工数据）+ 开始计算 + 查看底稿。"""
        panel = QWidget()
        panel.setMaximumWidth(LEFT_PANEL_MAX_W)
        panel.setMinimumWidth(340)
        v = QVBoxLayout(panel)
        v.setContentsMargins(0, 0, 0, 0)

        g = QGroupBox()
        outer = QVBoxLayout(g)
        head = QHBoxLayout()
        lb_title = QLabel("<b>输入设置</b>")
        head.addWidget(lb_title)
        head.addStretch(1)
        outer.addLayout(head)

        # ---- 第一步：上传报表 ----
        outer.addWidget(QLabel(f"<b>{W.STEP_INPUT}</b>"))
        mode_row = QHBoxLayout()
        self.rb_workbook = QRadioButton("单个工作簿")
        self.rb_folder = QRadioButton("文件夹 / 压缩包")
        self.rb_workbook.setChecked(True)
        self.rb_workbook.setToolTip(W.MODE_WORKBOOK)
        self.rb_folder.setToolTip(W.MODE_FOLDER)
        mode_row.addWidget(self.rb_workbook)
        mode_row.addWidget(self.rb_folder)
        mode_row.addStretch(1)
        outer.addLayout(mode_row)

        self.w_wb = QWidget()
        h_wb = QHBoxLayout(self.w_wb)
        h_wb.setContentsMargins(0, 0, 0, 0)
        h_wb.addWidget(QLabel(W.LABEL_WORKBOOK))
        self.ed_workbook = QLineEdit()
        btn_wb = QPushButton(W.BTN_PICK_FILE)
        btn_wb.clicked.connect(lambda: self._pick(self.ed_workbook, "工作簿 (*.xlsx)"))
        h_wb.addWidget(self.ed_workbook)
        h_wb.addWidget(btn_wb)
        outer.addWidget(self.w_wb)

        self.w_folder = QWidget()
        h_fd = QHBoxLayout(self.w_folder)
        h_fd.setContentsMargins(0, 0, 0, 0)
        h_fd.addWidget(QLabel(W.LABEL_FOLDER))
        self.ed_folder = QLineEdit()
        self.ed_folder.setPlaceholderText("含 4 张报表的文件夹，或 .zip")
        btn_dir = QPushButton(W.BTN_PICK_FOLDER)
        btn_dir.clicked.connect(self._pick_folder)
        btn_zip = QPushButton(W.BTN_PICK_ZIP)
        btn_zip.clicked.connect(self._pick_zip)
        h_fd.addWidget(self.ed_folder)
        h_fd.addWidget(btn_dir)
        h_fd.addWidget(btn_zip)
        outer.addWidget(self.w_folder)
        self.w_folder.setVisible(False)

        self.rb_workbook.toggled.connect(self._on_input_mode)
        self.rb_folder.toggled.connect(self._on_input_mode)
        self._on_input_mode()

        # ---- 第二步：报表期间 ----
        outer.addWidget(QLabel(f"<b>{W.STEP_PERIOD}</b>"))
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("报表月份"))
        self.cb_month = QComboBox()
        self.cb_month.setEditable(True)
        today = datetime.now()
        for k in range(MONTHS_BACK):
            y, m = divmod((today.year * 12 + today.month - 1 - k), 12)
            self.cb_month.addItem(f"{y}-{m + 1:02d}")
        self.cb_month.setToolTip("报表期间，格式如 2026-10")
        h1.addWidget(self.cb_month)
        h1.addStretch(1)
        outer.addLayout(h1)

        h1b = QHBoxLayout()
        h1b.addWidget(QLabel(W.SAVE_TO))
        self.lb_outdir = QLabel(self._outdir)
        self.lb_outdir.setStyleSheet("color:#666;")
        self.lb_outdir.setWordWrap(True)
        btn_o = QPushButton(W.BTN_CHANGE_DIR)
        btn_o.clicked.connect(self._pick_dir)
        h1b.addWidget(self.lb_outdir)
        h1b.addWidget(btn_o)
        outer.addLayout(h1b)

        # ---- 第三步：人工数据（2×2，紧凑） ----
        lb_manual = QLabel(f"<b>{W.STEP_MANUAL}</b>")
        lb_manual.setToolTip(W.HELP_MANUAL_HTML)
        outer.addWidget(lb_manual)
        self.manual_spins = {}
        last = _load_last_manual()
        grid = QGridLayout()
        positions = [(57, 0, 0), (58, 0, 2), (59, 1, 0), (60, 1, 2)]
        for row, r_i, c_i in positions:
            grid.addWidget(QLabel(W.MANUAL_LABELS[row]), r_i, c_i)
            sp = QSpinBox()
            sp.setRange(0, W.MANUAL_MAX)
            sp.setMaximumWidth(76)  # 6 位数足够（人数上限 999999），把宽度留给标签
            sp.setValue(last.get(row, 0))
            sp.valueChanged.connect(self._update_manual_hint)
            self.manual_spins[row] = sp
            grid.addWidget(sp, r_i, c_i + 1)
        grid.setColumnStretch(3, 1)
        outer.addLayout(grid)
        self.lb_manual_hint = QLabel()
        self.lb_manual_hint.setWordWrap(True)
        self.lb_manual_hint.setStyleSheet("color:#e65100;")
        outer.addWidget(self.lb_manual_hint)
        self._update_manual_hint()

        # ---- 动作 ----
        h_run = QHBoxLayout()
        self.btn_run = QPushButton(W.BTN_RUN)
        self.btn_run.setDefault(True)
        self.btn_run.setMinimumHeight(34)
        self.btn_run.clicked.connect(self._on_run)
        h_run.addWidget(self.btn_run)
        outer.addLayout(h_run)

        self.btn_workpaper = QPushButton(W.BTN_WORKPAPER)
        self.btn_workpaper.clicked.connect(self._show_workpaper)
        self.btn_workpaper.setEnabled(False)
        outer.addWidget(self.btn_workpaper)

        v.addWidget(g)
        v.addStretch(1)
        return panel

    def _build_right_panel(self) -> QWidget:
        """右侧核心区：计算与核对（汇总/核对/确认/结果表/明细/导出）。

        垂直空间优先给结果表：标题并入汇总行、图例改为悬停提示、明细面板默认收起。
        """
        panel = QWidget()
        panel.setMinimumWidth(640)
        v = QVBoxLayout(panel)
        v.setContentsMargins(4, 2, 0, 2)
        v.setSpacing(4)

        self.lb_summary = QLabel()
        self.lb_summary.setStyleSheet("font-weight:bold;")
        self.lb_summary.setToolTip(W.HELP_RESULT_HTML)  # 悬停显示核对说明与行状态图例
        v.addWidget(self.lb_summary)
        self.lb_checks = QLabel()
        self.lb_checks.setStyleSheet("font-weight:bold;")
        v.addWidget(self.lb_checks)
        self.lb_checks_detail = QLabel()
        self.lb_checks_detail.setWordWrap(True)
        self.lb_checks_detail.setStyleSheet("color:#666;")
        v.addWidget(self.lb_checks_detail)

        # ---- 人工确认工具条：按钮右对齐到「确认无误」列正上方 ----
        h_confirm = QHBoxLayout()
        self.ck_abnormal = QCheckBox(W.FILTER_ABNORMAL)
        self.ck_abnormal.toggled.connect(self._refresh)
        h_confirm.addWidget(self.ck_abnormal)
        h_confirm.addStretch(1)
        self.lb_confirm_progress = QLabel()
        self.lb_confirm_progress.setToolTip(W.CONFIRM_HINT)
        h_confirm.addWidget(self.lb_confirm_progress)
        self.btn_confirm_all = QPushButton(W.BTN_CONFIRM_ALL)
        self.btn_confirm_all.clicked.connect(self._confirm_all)
        self.btn_confirm_none = QPushButton(W.BTN_CONFIRM_NONE)
        self.btn_confirm_none.clicked.connect(self._confirm_none)
        h_confirm.addWidget(self.btn_confirm_all)
        h_confirm.addWidget(self.btn_confirm_none)
        v.addLayout(h_confirm)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            [W.COL_STATUS, W.COL_NAME, W.COL_RAW, W.COL_ROUNDED, W.COL_SOURCE, W.COL_ACK])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 44)   # 状态
        self.table.setColumnWidth(2, 84)   # 原始值
        self.table.setColumnWidth(3, 84)   # 取整后
        self.table.setColumnWidth(5, 64)   # 确认无误（固定窄列，与上方按钮对齐）
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.DoubleClicked)
        self.table.itemChanged.connect(self._on_item_changed)
        self.table.itemSelectionChanged.connect(self._on_select)
        v.addWidget(self.table, 1)

        self.detail = QPlainTextEdit()
        self.detail.setReadOnly(True)
        self.detail.setPlaceholderText(W.DETAIL_PLACEHOLDER)
        self.detail.setVisible(False)  # 默认收起，点击行自动展开
        self._detail_visible = False
        v.addWidget(self.detail)

        h_btn = QHBoxLayout()
        self.btn_export = QPushButton(W.BTN_EXPORT)
        self.btn_export.setMinimumHeight(32)
        self.btn_export.setMinimumWidth(160)
        self.btn_export.clicked.connect(self._export)
        self.btn_export.setEnabled(False)
        h_btn.addWidget(self.btn_export)
        h_btn.addStretch(1)
        self.btn_detail_toggle = QPushButton("取数明细 ▼")
        self.btn_detail_toggle.setToolTip("展开/收起选中行的取数过程（来源表、单元格、金额换算）")
        self.btn_detail_toggle.clicked.connect(self._toggle_detail)
        h_btn.addWidget(self.btn_detail_toggle)
        v.addLayout(h_btn)
        return panel

    def _toggle_detail(self):
        """展开/收起钻取明细面板。"""
        self._detail_visible = not self._detail_visible
        self.detail.setVisible(self._detail_visible)
        self.btn_detail_toggle.setText("取数明细 ▲" if self._detail_visible else "取数明细 ▼")

    # ---------- 输入辅助 ----------

    def _pick(self, edit: QLineEdit, pattern: str):
        path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", pattern)
        if path:
            edit.setText(path)

    def _pick_folder(self):
        path = QFileDialog.getExistingDirectory(self, "选择报表文件夹", "")
        if path:
            self.ed_folder.setText(path)

    def _pick_zip(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择压缩包", "", "压缩包 (*.zip)")
        if path:
            self.ed_folder.setText(path)

    def _pick_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择保存位置", self._outdir)
        if path:
            self._outdir = path
            self.lb_outdir.setText(path)

    def _on_input_mode(self):
        self.w_wb.setVisible(self.rb_workbook.isChecked())
        self.w_folder.setVisible(self.rb_folder.isChecked())

    def _paths(self) -> dict | None:
        """按当前输入方式收集路径；不完整时弹提示返回 None。"""
        if self.rb_workbook.isChecked():
            wb = self.ed_workbook.text().strip()
            if not wb:
                QMessageBox.warning(self, "输入不完整", W.ERR_NO_WORKBOOK)
                return None
            return {"workbook": wb}
        fd = self.ed_folder.text().strip()
        if not fd:
            QMessageBox.warning(self, "输入不完整", W.ERR_NO_FOLDER)
            return None
        if fd.lower().endswith(".zip"):
            return {"zip": fd}
        return {"folder": fd}

    def _month(self) -> str | None:
        m = self.cb_month.currentText().strip()
        parts = m.split("-")
        if len(parts) != 2 or not all(p.isdigit() for p in parts):
            QMessageBox.warning(self, "月份格式", W.ERR_BAD_MONTH)
            return None
        return m

    def _manual(self) -> dict:
        """4 项人数原样提交（0 是有效值，不等于未填）。

        「未填」只发生在 CLI 未提供 --manual 的场景；界面端输入框始终有值，
        是否已审核由 V2.10 的逐项人工确认把关。
        """
        return {row: Decimal(sp.value()) for row, sp in self.manual_spins.items()}

    def _update_manual_hint(self):
        """四个人数全为 0 时给一句非阻断提示（0 有效，但提醒确认是否已取得人力数据）。"""
        if all(sp.value() == 0 for sp in self.manual_spins.values()):
            self.lb_manual_hint.setText(
                "四个人数均为 0（0 是有效值）。如本期尚未取得人力部门数据，请先获取再计算。")
        else:
            self.lb_manual_hint.setText("")

    # ---------- 计算 ----------

    def _on_run(self):
        paths = self._paths()
        if paths is None:
            return
        month = self._month()
        if month is None:
            return
        self.btn_run.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.status.showMessage(W.RUNNING)
        self._thread = ComputeThread(paths, month, self._manual())
        self._thread.progress.connect(lambda s: self.status.showMessage(W.STAGE_PREFIX + s))
        self._thread.done.connect(self._done)
        self._thread.failed.connect(self._failed)
        self._thread.start()

    def _done(self, outcome):
        self.btn_run.setEnabled(True)
        self.outcome = outcome
        _save_last_manual(self._manual())
        c = outcome.counts()
        self.status.showMessage(W.DONE_FMT.format(pass_n=c["PASS"], warn_n=c["WARN"],
                                                  fail_n=c["FAIL"], pending_n=c["PENDING"]))
        self._refresh()

    def _failed(self, msg: str):
        self.btn_run.setEnabled(True)
        self.status.showMessage("计算失败")
        QMessageBox.critical(self, W.FAIL_TITLE, msg)

    # ---------- 人工确认 ----

    def _confirm_all(self):
        if self.outcome is None:
            return
        self.outcome.acknowledged |= self.outcome.confirmable_rows()
        self._refresh()

    def _confirm_none(self):
        if self.outcome is None:
            return
        self.outcome.acknowledged.clear()
        self._refresh()

    def _ack(self, row: int, on: bool):
        if on:
            self.outcome.acknowledged.add(row)
        else:
            self.outcome.acknowledged.discard(row)
        self._refresh()

    # ---------- 结果填充 ----------

    def _display_name(self, r) -> str:
        """项目名：优先用输出模板 A 列（含层级缩进与编号），回退到映射名。"""
        name = display_names().get(r.row, r.name)
        return W.indent_name(name)

    def _refresh(self):
        oc = self.outcome
        if oc is None:
            self.lb_summary.setText(W.EMPTY_SUMMARY)
            self.lb_checks.setText("")
            self.lb_checks_detail.setText("")
            self.table.setRowCount(0)
            self.btn_export.setEnabled(False)
            self.btn_workpaper.setEnabled(False)
            self.btn_confirm_all.setEnabled(False)
            self.btn_confirm_none.setEnabled(False)
            self.lb_confirm_progress.setText("")
            return

        c = oc.counts()
        self.lb_summary.setText(
            f"第四步　计算与核对　　" + W.SUMMARY_FMT.format(
                total=len(oc.results), pass_icon=W.STATUS_ICON["PASS"], pass_n=c["PASS"],
                warn_icon=W.STATUS_ICON["WARN"], warn_n=c["WARN"],
                fail_icon=W.STATUS_ICON["FAIL"], fail_n=c["FAIL"],
                pending_icon=W.STATUS_ICON["PENDING"], pending_n=c["PENDING"],
                const_icon=W.STATUS_ICON["CONST"], const_n=c["CONST"]))
        summary, detail = self._check_panels(oc)
        self.lb_checks.setText(summary)
        self.lb_checks_detail.setText(detail)
        has_problem = summary != W.CHECK_ALL_PASS
        self.lb_checks_detail.setStyleSheet("color:#b71c1c;" if has_problem else "color:#666;")

        # 人工确认进度
        done, total = oc.confirm_progress()
        self.lb_confirm_progress.setText(
            W.CONFIRM_PROGRESS_FMT.format(done=done, total=total)
            + (f"（{W.CONFIRM_ALL_DONE}）" if total and done == total else ""))
        self.btn_confirm_all.setEnabled(done < total)
        self.btn_confirm_none.setEnabled(done > 0)

        only_abnormal = self.ck_abnormal.isChecked()
        rows = [r for r in oc.results
                if not only_abnormal or r.status in (RowStatus.FAIL, RowStatus.WARN, RowStatus.PENDING)]

        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self._fill_row(i, r)
        self.table.blockSignals(False)
        self.btn_workpaper.setEnabled(True)
        self._update_export_btn()

    def _check_panels(self, oc) -> tuple[str, str]:
        """数据核对：返回 (汇总行, 明细行)。

        全部通过：汇总行一句话，明细行说明核对范围（另有 ℹ️ 提示则附后）；
        有问题：汇总行报错误/警告数量，明细行逐行列出（每行一条，不用｜分隔）。
        """
        problems, infos = [], []  # (mark, title, detail)
        for _name, rules, _pass_text in W.CHECK_GROUPS:
            rel = [i for i in oc.issues if i.rule in rules]
            if not rel:
                continue
            e = next((i for i in rel if i.level == Level.E), None)
            w = next((i for i in rel if i.level == Level.W), None)
            first = e or w or rel[0]
            mark = W.CHECK_E_MARK if e else (W.CHECK_W_MARK if w else "ℹ️")
            (problems if (e or w) else infos).append((mark, first.title, first.detail))
        info_lines = [f"ℹ️ {t}" + (f"（{d}）" if d else "") for _m, t, d in infos]
        if not problems:
            detail = W.CHECK_PASS_SCOPE + (("\n" + "\n".join(info_lines)) if info_lines else "")
            return W.CHECK_ALL_PASS, detail
        n_e = sum(1 for p in problems if p[0] == W.CHECK_E_MARK)
        n_w = sum(1 for p in problems if p[0] == W.CHECK_W_MARK)
        parts = []
        if n_e:
            parts.append(f"{W.CHECK_E_MARK} {n_e} 项错误")
        if n_w:
            parts.append(f"{W.CHECK_W_MARK} {n_w} 项警告")
        summary = "数据核对：" + "　".join(parts)
        lines = [f"{m} {t}" + (f"（{d}）" if d else "") for m, t, d in problems]
        return summary, "\n".join(lines + info_lines)

    def _fill_row(self, i, r):
        oc = self.outcome

        it_status = QTableWidgetItem(W.STATUS_ICON[r.status.value])
        it_status.setTextAlignment(Qt.AlignCenter)
        it_status.setToolTip(W.STATUS_TEXT[r.status.value])
        if any(s.empty for s in r.sources):
            it_status.setToolTip(W.STATUS_TEXT[r.status.value] + "；" + W.STATUS_TIP_EMPTY)
        self.table.setItem(i, 0, it_status)

        it_name = QTableWidgetItem(self._display_name(r))
        it_name.setData(Qt.UserRole, r.row)  # 行次存入角色数据（界面不单独占一列）
        self.table.setItem(i, 1, it_name)

        it_raw = QTableWidgetItem(W.fmt_wan(r.raw))
        it_raw.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if r.raw is not None:
            frac = abs(Decimal(str(r.raw))) % Decimal("1")
            if Decimal("0.45") <= frac <= Decimal("0.55"):
                it_raw.setBackground(CRITICAL_BG)
                it_raw.setToolTip(W.ROUND_CRITICAL_TIP)
        self.table.setItem(i, 2, it_raw)

        if r.status == RowStatus.PENDING:
            it_val = QTableWidgetItem("待填报")
            it_val.setToolTip(W.PENDING_VALUE_TIP)
        else:
            it_val = QTableWidgetItem(W.fmt_int(r.rounded))
        it_val.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        if r.status == RowStatus.PENDING:
            it_val.setFlags(it_val.flags() | Qt.ItemIsEditable)
        self.table.setItem(i, 3, it_val)

        reason = _source_brief(r)
        if r.messages:
            reason += "　｜　" + "；".join(r.messages)
        self.table.setItem(i, 4, QTableWidgetItem(reason))

        # 确认无误：可确认状态的行才可勾选（错误行须先修正，待填报行须先补数）
        ck = QCheckBox()
        confirmable = r.status in (RowStatus.PASS, RowStatus.WARN, RowStatus.CONST)
        ck.setChecked(r.row in oc.acknowledged)
        ck.setEnabled(confirmable)
        if confirmable:
            ck.toggled.connect(lambda on, row=r.row: self._ack(row, on))
        elif r.status == RowStatus.PENDING:
            ck.setToolTip("先填报该项人数")
        else:
            ck.setToolTip("先修正该行错误")
        self.table.setCellWidget(i, 5, ck)

        for col in range(5):
            item = self.table.item(i, col)
            if item is not None:
                item.setBackground(COLOR[r.status])
                if r.status == RowStatus.FAIL:
                    item.setForeground(QColor("#b71c1c"))

    # ---------- 交互 ----------

    def _result_at(self, table_row: int):
        item = self.table.item(table_row, 1)
        if item is None:
            return None
        row_no = item.data(Qt.UserRole)
        if row_no is None:
            return None
        return next((r for r in self.outcome.results if r.row == int(row_no)), None)

    def _on_select(self):
        items = self.table.selectedItems()
        if not items:
            return
        res = self._result_at(items[0].row())
        if res is None:
            return
        if not self._detail_visible:  # 点击行自动展开明细（可再手动收起）
            self._toggle_detail()
        self.detail.setPlainText(self._detail_text(res))

    def _detail_text(self, res) -> str:
        """钻取明细：完整取数链路，财务语言。"""
        lines = [W.DETAIL_LINE_HEAD.format(row=res.row, name=self._display_name(res).strip(),
                                           status=W.STATUS_TEXT[res.status.value])]
        if res.sources:
            for s in res.sources:
                if s.type == "MANUAL":
                    lines.append("取数：人工填报\"人数\"（人力部门提供，单位：人）")
                    lines.append(f"数值：{W.fmt_int(s.raw_amount)} 人（人数不换算，直接填列）")
                    continue
                if s.type == "TB":
                    name = s.ref.split(" ", 1)[1] if " " in s.ref else s.ref
                    field = W.TB_FIELD_LABELS.get(s.field, s.field)
                    tpl = W.SRC_TB_FMT
                else:
                    name = s.ref.split("·", 1)[1] if "·" in s.ref else s.ref
                    field = W.FIELD_BY_TYPE.get(s.type, s.field)
                    tpl = W.SRC_REPORT_FMT.replace("{table}", W.SHEET_BY_TYPE.get(s.type, s.type))
                _, _, addr = s.cell.partition("!") if s.cell else (None, None, "")
                src = "取数：" + (tpl.replace("{cell}", addr)
                                       .replace("{name}", name)
                                       .replace("{field}", field))
                if s.empty:
                    src += "——该单元格为空，按 0 填列"
                lines.append(src)
                lines.append(f"　　　{W.fmt_yuan(s.raw_amount)} 元 → {W.fmt_wan(s.converted)} 万元")
            if len(res.sources) > 1:
                lines.append(f"合计：{W.fmt_wan(res.raw)} 万元 → 取整 {W.fmt_int(res.rounded)} 万元")
            elif res.sources[0].type != "MANUAL":
                lines.append(f"取整：{W.fmt_int(res.rounded)} 万元")
        elif res.status == RowStatus.PENDING:
            lines.append("取数：待人工填报（单位：人，在左侧「人工数据」或本行双击录入）")
        elif res.expression.startswith("按模板固定填"):
            lines.append(f"取数：{res.expression}")
        else:
            lines.append(f"计算：{res.expression}")
            if res.formula_detail:
                lines.append(res.formula_detail)
            lines.append(f"结果：{W.fmt_wan(res.raw)} 万元 → 取整 {W.fmt_int(res.rounded)} 万元")
        if res.messages:
            lines.append("提示：" + "；".join(res.messages))
        return "\n".join(lines)

    def _on_item_changed(self, item):
        """人工项就地编辑 → 同步输入框、重算，并取消数值发生变化行的确认。"""
        if item.column() != 3:
            return
        res = self._result_at(item.row())
        if res is None:
            return
        spec = next((s for s in self.outcome.items if s.row == res.row), None)
        if spec is None or spec.kind != "MANUAL":
            return
        try:
            v = Decimal(item.text().strip())
        except Exception:  # noqa: BLE001
            QMessageBox.warning(self, "数值非法", f"「{item.text()}」不是数值")
            self._refresh()
            return
        if v < 0 or v > W.MANUAL_MAX:
            QMessageBox.warning(self, "数值超出范围", W.MANUAL_EDIT_LIMIT_MSG)
            self._refresh()
            return
        before = {r.row: r.rounded for r in self.outcome.results}
        manual = dict(self.outcome.manual)
        manual[res.row] = v
        if res.row in self.manual_spins:
            self.manual_spins[res.row].setValue(int(v))
        self.outcome.recompute(manual)
        # 数值发生变化的行（含依赖它的公式行）取消确认，须重新核对
        changed = {r.row for r in self.outcome.results if before.get(r.row) != r.rounded}
        self.outcome.acknowledged -= changed
        self._refresh()
        QMessageBox.information(self, W.RECOMPUTED_TITLE,
                                W.RECOMPUTED_MSG + f"\n数值变化的 {len(changed)} 行已取消确认，请重新核对。")

    def _update_export_btn(self):
        if self.outcome is None:
            self.btn_export.setEnabled(False)
            return
        ok, reason = self.outcome.can_export()
        self.btn_export.setEnabled(ok)
        if ok:
            self.btn_export.setToolTip(W.EXPORT_OK_TIP.format(outdir=self._outdir))
            self.btn_export.setText(W.BTN_EXPORT)
        else:
            self.btn_export.setToolTip(reason)
            self.btn_export.setText(W.BTN_EXPORT + "（条件未满足）")

    # ---------- 动作 ----------

    def _show_workpaper(self):
        """计算底稿：运行编号/期间/输入输出文件/核对结论/确认记录（财务语言）。"""
        if self.outcome is None:
            QMessageBox.information(self, W.WORKPAPER_TITLE, "尚未计算。请先点「开始计算」。")
            return
        oc = self.outcome
        inp = Path(oc.inputs.workbook_path).name if oc.inputs.workbook_path else "文件夹 / 压缩包（4 张报表）"
        out = Path(oc.exported_path).name if oc.exported_path else W.WORKPAPER_NOT_EXPORTED

        check_summary, check_detail = self._check_panels(oc)
        done, total = oc.confirm_progress()
        ack_text = (W.WORKPAPER_ACK_NONE if done == 0 else
                    (f"全部 {total} 项已确认" if done == total else f"已确认 {done}/{total} 项"))

        lines = [
            f"{W.WORKPAPER_RUN_ID}：{oc.run_id}（可用于追溯本次计算）",
            f"{W.WORKPAPER_MONTH}：{oc.month}",
            f"{W.WORKPAPER_INPUT}：{inp}",
            f"{W.WORKPAPER_OUTPUT}：{out}",
            f"{W.WORKPAPER_CHECKS}：{check_summary}",
        ]
        for ln in check_detail.split("\n"):
            lines.append(f"  {ln}")
        lines += [
            f"{W.WORKPAPER_ACKS}：{ack_text}",
            "",
            W.WORKPAPER_HINT,
            f"{W.WORKPAPER_LOGS}：",
            f"  计算过程　{oc.calc_log}",
            f"  核对记录　{oc.validation_log}",
        ]
        dlg = QDialog(self)
        dlg.setWindowTitle(W.WORKPAPER_TITLE)
        dlg.resize(680, 460)
        v = QVBoxLayout(dlg)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setFont(QFont("Courier New", 12))
        text.setPlainText("\n".join(lines))
        v.addWidget(text)
        h = QHBoxLayout()
        h.addStretch(1)
        btn_logs = QPushButton(W.BTN_LOGS)
        from ..logger import default_log_root
        btn_logs.clicked.connect(lambda: webbrowser.open(default_log_root().resolve().as_uri()))
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(dlg.accept)
        h.addWidget(btn_logs)
        h.addWidget(btn_close)
        v.addLayout(h)
        dlg.exec()

    def _export(self):
        if self.outcome is None:
            return
        ok, reason = self.outcome.can_export()
        if not ok:
            QMessageBox.warning(self, "还不能导出", reason)
            return
        outdir = Path(self._outdir)
        try:
            confirmations = [
                {"row": r.row, "item": r.name, "status": r.status.value}
                for r in self.outcome.results if r.row in self.outcome.acknowledged
            ]
            path = service.export_run(self.outcome, str(outdir), confirmations=confirmations)
        except PermissionError as e:
            QMessageBox.warning(self, "还不能导出", str(e))
            return
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "导出失败", f"{e}\n（临时文件已清理，不产生半成品）")
            return

        c = self.outcome.counts()
        done, total = self.outcome.confirm_progress()
        box = QMessageBox(self)
        box.setWindowTitle(W.EXPORT_DONE_TITLE)
        box.setText(W.EXPORT_DONE_FMT.format(
            filename=path.name, outdir=str(outdir),
            summary=f"共 {len(self.outcome.results)} 项：{c['PASS']} 项正常，{c['WARN']} 项需确认；"
                    f"已人工确认 {done}/{total} 项"))
        open_btn = box.addButton(W.EXPORT_DONE_OPEN, QMessageBox.AcceptRole)
        wp_btn = box.addButton(W.EXPORT_DONE_WORKPAPER, QMessageBox.ActionRole)
        box.addButton(W.EXPORT_DONE_CLOSE, QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is open_btn:
            webbrowser.open(path.parent.resolve().as_uri())
        elif box.clickedButton() is wp_btn:
            self._show_workpaper()
        self._update_export_btn()
