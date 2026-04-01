from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = BASE_DIR / "config"
DATA_DIR = BASE_DIR / "data"
LOG_DIR = BASE_DIR / "logs"


def get_user_home() -> Path:
    return Path.home()


def get_default_save_dir() -> Path:
    return get_user_home() / "Documents" / "YeahInvoiceCollector" / "attachments"


def get_default_export_dir() -> Path:
    return get_user_home() / "Documents" / "YeahInvoiceCollector" / "exports"


def normalize_path(value: str | Path) -> Path:
    return Path(value).expanduser().resolve()


def ensure_dir(path: str | Path) -> Path:
    target = Path(path).expanduser()
    target.mkdir(parents=True, exist_ok=True)
    return target.resolve()


def open_folder(path: str | Path) -> None:
    target = normalize_path(path)
    if not target.exists():
        raise FileNotFoundError(f"目录不存在: {target}")
    if sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=True)
        return
    if os.name == "nt":
        subprocess.run(["explorer", str(target)], check=True)
        return
    subprocess.run(["xdg-open", str(target)], check=True)


def ensure_runtime_dirs() -> None:
    for path in (CONFIG_DIR, DATA_DIR, LOG_DIR, get_default_save_dir(), get_default_export_dir()):
        ensure_dir(path)
