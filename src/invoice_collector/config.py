from __future__ import annotations

import json
from dataclasses import asdict
from datetime import datetime, timedelta
from pathlib import Path

from .models import AppConfig, EmailConfig, MailFilterConfig, ParsingConfig, SettlementGroupConfig, StorageConfig, UiPreferences
from .paths import BUNDLE_DIR, CONFIG_DIR, ensure_runtime_dirs, get_default_export_dir, get_default_save_dir


DEFAULT_CONFIG_PATH = CONFIG_DIR / "config.json"
EXAMPLE_CONFIG_PATH = BUNDLE_DIR / "config" / "config.example.json"


def _read_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def default_recent_criteria(days: int = 30) -> str:
    dt = datetime.now() - timedelta(days=days)
    return f"SINCE {dt.strftime('%d-%b-%Y')}"


def _default_config_dict() -> dict:
    return {
        "email": {
            "email_address": "",
            "auth_code": "",
            "imap_host": "imap.yeah.net",
            "imap_port": 993,
            "use_ssl": True,
            "mail_folder": "INBOX",
            "search_criteria": default_recent_criteria(30),
        },
        "mail_filters": {
            "sender_keywords": ["中国电信", "189.cn"],
            "subject_keywords": ["电子发票", "中国电信"],
        },
        "parsing": {
            "phone_regex": r"(?<!\d)(1\d{10})(?!\d)",
            "billing_period_patterns": [
                r"(20\d{2})年(0?[1-9]|1[0-2])月",
                r"(20\d{2})-(0?[1-9]|1[0-2])",
                r"(20\d{2})(0[1-9]|1[0-2])",
            ],
        },
        "settlement_groups": {
            "default": "大盘",
            "mappings": {
                "19120056416": "墨秀",
                "19120076109": "墨秀",
                "19120076054": "墨秀",
            },
        },
        "storage": {
            "database_path": "data/invoices.db",
            "save_root": str(get_default_save_dir()),
            "export_root": str(get_default_export_dir()),
            "folder_rule": "{account_name}/{settlement_group}/{billing_month}",
        },
        "ui_preferences": {
            "recent_configs": [],
            "last_save_root": str(get_default_save_dir()),
            "log_mode": "simple",
            "default_export_scope": "all",
            "window_width": 1360,
            "window_height": 860,
        },
    }


def _merge_defaults(data: dict) -> dict:
    email_data = dict(data.get("email", {}))
    if "username" in email_data and "email_address" not in email_data:
        email_data["email_address"] = email_data.pop("username")
    if "password" in email_data and "auth_code" not in email_data:
        email_data["auth_code"] = email_data.pop("password")
    if "folder" in email_data and "mail_folder" not in email_data:
        email_data["mail_folder"] = email_data.pop("folder")

    storage_data = dict(data.get("storage", {}))
    if "attachments_root" in storage_data and "save_root" not in storage_data:
        storage_data["save_root"] = storage_data.pop("attachments_root")

    data = dict(data)
    data["email"] = email_data
    data["storage"] = storage_data

    default = _default_config_dict()
    merged = default | data
    for key in ("email", "mail_filters", "parsing", "settlement_groups", "storage", "ui_preferences"):
        merged[key] = default[key] | data.get(key, {})
    if not merged["storage"].get("save_root"):
        merged["storage"]["save_root"] = str(get_default_save_dir())
    if not merged["storage"].get("export_root"):
        merged["storage"]["export_root"] = str(get_default_export_dir())
    if not merged["ui_preferences"].get("last_save_root"):
        merged["ui_preferences"]["last_save_root"] = merged["storage"]["save_root"]
    return merged


def load_config(path: str | Path | None = None) -> AppConfig:
    ensure_runtime_dirs()
    config_path = Path(path) if path else DEFAULT_CONFIG_PATH
    if not config_path.exists():
        save_raw_config(_default_config_dict(), config_path)
    data = _merge_defaults(_read_json(config_path))
    return AppConfig(
        email=EmailConfig(**data["email"]),
        mail_filters=MailFilterConfig(**data["mail_filters"]),
        parsing=ParsingConfig(**data["parsing"]),
        settlement_groups=SettlementGroupConfig(**data["settlement_groups"]),
        storage=StorageConfig(**data["storage"]),
        ui_preferences=UiPreferences(**data["ui_preferences"]),
    )


def save_config(config: AppConfig, path: str | Path | None = None) -> Path:
    config_path = Path(path) if path else DEFAULT_CONFIG_PATH
    save_raw_config(asdict(config), config_path)
    return config_path


def update_recent_configs(config: AppConfig, config_path: str | Path) -> AppConfig:
    path_text = str(Path(config_path).expanduser())
    recent = [path_text]
    recent.extend(item for item in config.ui_preferences.recent_configs if item != path_text)
    config.ui_preferences.recent_configs = recent[:8]
    return config


def save_raw_config(data: dict, path: str | Path) -> None:
    target = Path(path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    merged = _merge_defaults(data)
    with target.open("w", encoding="utf-8") as file:
        json.dump(merged, file, ensure_ascii=False, indent=2)
