from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from .config import load_config
from .database import InvoiceDatabase
from .paths import BASE_DIR, ensure_runtime_dirs
from .ui.main_window import MainWindow


def main() -> None:
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
