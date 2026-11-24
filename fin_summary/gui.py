"""GUI 入口：python -m fin_summary.gui"""
from __future__ import annotations

import sys


def main() -> int:
    try:
        from PySide6.QtWidgets import QApplication
    except ImportError:
        print("未安装 PySide6，请执行: pip install PySide6", file=sys.stderr)
        return 1
    from .ui.main_window import MainWindow

    app = QApplication(sys.argv)
    app.setApplicationName("月度财务数据汇总表生成工具")
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
