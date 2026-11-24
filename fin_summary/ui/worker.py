"""后台计算线程：避免 UI 假死。"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Signal

from .. import service
from ..loader import ParseError
from ..config_loader import MappingError


class ComputeThread(QThread):
    progress = Signal(str)
    done = Signal(object)       # Outcome
    failed = Signal(str)

    def __init__(self, inputs_paths: dict, month: str, manual: dict, log_dir: str | None = None):
        super().__init__()
        self.inputs_paths = inputs_paths
        self.month = month
        self.manual = manual
        self.log_dir = log_dir

    def run(self):
        try:
            outcome = service.compute_run(
                self.inputs_paths, self.month, self.manual,
                log_dir=self.log_dir,
                progress=lambda s: self.progress.emit(s),
            )
            self.done.emit(outcome)
        except ParseError as e:
            # 财务语言消息，不带规则编号（编号保留在日志/CLI）
            self.failed.emit(e.message)
        except MappingError as e:
            self.failed.emit(str(e))
        except Exception as e:  # noqa: BLE001
            import traceback
            self.failed.emit(f"{e}\n{traceback.format_exc(limit=3)}")


def build_paths(workbook: str | None, files: dict) -> dict | None:
    """工作簿优先；否则 4 文件。"""
    if workbook:
        return {"workbook": workbook}
    if all(files.get(k) for k in ("资产负债表", "利润表", "现金流量表", "科目余额表")):
        return {k: v for k, v in files.items()}
    return None
