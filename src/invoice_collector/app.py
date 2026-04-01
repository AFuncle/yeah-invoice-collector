from __future__ import annotations

import sys
import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from .config import load_config
from .database import InvoiceDatabase
from .paths import BASE_DIR, LOG_DIR, ensure_runtime_dirs
from .ui.main_window import MainWindow


def _install_exception_hook() -> None:
    def handler(exc_type, exc_value, exc_tb):
        error_text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        try:
            log_file = LOG_DIR / "crash.log"
            log_file.parent.mkdir(parents=True, exist_ok=True)
            with log_file.open("a", encoding="utf-8") as f:
                f.write(f"\n{'=' * 60}\n[{datetime.now().isoformat()}]\n{error_text}\n")
        except Exception:
            pass
        try:
            QMessageBox.critical(
                None,
                "程序异常",
                f"发生未处理的异常，详情已记录到 logs/crash.log。\n\n{error_text[-1500:]}",
            )
        except Exception:
            pass

    sys.excepthook = handler


def main() -> None:
    _install_exception_hook()
    ensure_runtime_dirs()
    config = load_config()
    db_path = Path(config.storage.database_path)
    if not db_path.is_absolute():
        db_path = BASE_DIR / db_path
    database = InvoiceDatabase(db_path)

    app = QApplication(sys.argv)
    window = MainWindow(database)
    window.show()
    sys.exit(app.exec())
