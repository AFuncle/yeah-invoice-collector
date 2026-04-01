from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


@dataclass(slots=True)
class EmailConfig:
    email_address: str = ""
    auth_code: str = ""
    imap_host: str = "imap.yeah.net"
    imap_port: int = 993
    use_ssl: bool = True
    mail_folder: str = "INBOX"
    search_criteria: str = "ALL"


@dataclass(slots=True)
class MailFilterConfig:
    sender_keywords: list[str] = field(default_factory=list)
    subject_keywords: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ParsingConfig:
    phone_regex: str = r"(?<!\d)(1\d{10})(?!\d)"
    billing_period_patterns: list[str] = field(default_factory=list)


@dataclass(slots=True)
class SettlementGroupConfig:
    default: str = "大盘"
    mappings: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class StorageConfig:
    database_path: str = "data/invoices.db"
    save_root: str = ""
    export_root: str = ""
    folder_rule: str = "{account_name}/{settlement_group}/{billing_month}"


@dataclass(slots=True)
class UiPreferences:
    recent_configs: list[str] = field(default_factory=list)
    last_save_root: str = ""
    log_mode: str = "simple"
    default_export_scope: str = "all"
    window_width: int = 1360
    window_height: int = 860


@dataclass(slots=True)
class AppConfig:
    email: EmailConfig = field(default_factory=EmailConfig)
    mail_filters: MailFilterConfig = field(default_factory=MailFilterConfig)
    parsing: ParsingConfig = field(default_factory=ParsingConfig)
    settlement_groups: SettlementGroupConfig = field(default_factory=SettlementGroupConfig)
    storage: StorageConfig = field(default_factory=StorageConfig)
    ui_preferences: UiPreferences = field(default_factory=UiPreferences)


@dataclass(slots=True)
class MailAttachment:
    filename: str
    payload: bytes
    size: int


@dataclass(slots=True)
class ParsedMail:
    message_uid: str
    message_id: str
    sender: str
    subject: str
    received_at: str
    phone_number: str
    billing_period: str
    settlement_group: str
    attachments: list[MailAttachment] = field(default_factory=list)


@dataclass(slots=True)
class InvoiceRecord:
    message_uid: str
    message_id: str
    sender: str
    subject: str
    phone_number: str
    billing_period: str
    settlement_group: str
    attachment_name: str
    attachment_path: str
    attachment_size: int
    received_at: str
    amount: str = ""
    account_name: str = ""
    status: str = "saved"
    error_message: str = ""
    collected_at: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    @property
    def attachment_path_obj(self) -> Path:
        return Path(self.attachment_path)


@dataclass(slots=True)
class CollectSummary:
    matched_mails: int = 0
    processed_mails: int = 0
    total_attachments: int = 0
    saved_records: int = 0
    duplicate_records: int = 0
    skipped_records: int = 0
    failed_records: int = 0
    invalid_attachments: int = 0


@dataclass(slots=True)
class ConnectionTestResult:
    server_ok: bool
    login_ok: bool
    folder_ok: bool
    readable_count: int | None
    message: str
