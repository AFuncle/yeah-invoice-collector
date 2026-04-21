from __future__ import annotations

import re
from pathlib import Path

from .database import InvoiceDatabase
from .imap_client import ImapMailClient
from .models import AppConfig, CollectSummary, InvoiceRecord
from .parser import build_parsed_mail
from .pdf_parser import extract_invoice_amount
from .paths import ensure_dir


def sanitize_filename(filename: str) -> str:
    return re.sub(r"[\\/:*?\"<>|]+", "_", filename)


class InvoiceCollectorService:
    def __init__(self, config: AppConfig, database: InvoiceDatabase) -> None:
        self.config = config
        self.database = database

    def collect(self, logger, progress_callback=None, stop_requested=None) -> CollectSummary:
        summary = CollectSummary()
        with ImapMailClient(self.config) as client:
            messages = client.fetch_target_messages(
                logger=logger,
                progress_callback=progress_callback,
                stop_requested=stop_requested,
            )
            summary.matched_mails = len(messages)
            for index, message in enumerate(messages, start=1):
                if stop_requested and stop_requested():
                    raise InterruptedError("采集已手动停止。")
                if progress_callback:
                    progress_callback(
                        {
                            "step": "正在写入结果",
                            "current": index,
                            "total": len(messages),
                            "saved": summary.saved_records,
                            "duplicate": summary.duplicate_records,
                            "failed": summary.failed_records,
                        }
                    )
                summary.processed_mails = index
                try:
                    parsed = build_parsed_mail(
                        message_uid=message["uid"],
                        message_id=message["message_id"],
                        sender=message["sender"],
                        subject=message["subject"],
                        received_at=message["received_at"],
                        attachments=message["attachments"],
                        parsing=self.config.parsing,
                        settlement_groups=self.config.settlement_groups,
                    )
                except Exception as exc:
                    summary.failed_records += 1
                    logger("error", f"跳过 UID={message['uid']}，原因：{exc}")
                    continue

                if not parsed.attachments:
                    summary.invalid_attachments += 1
                    logger("detail", f"跳过 UID={parsed.message_uid}，原因：没有附件。")
                    continue

                for attachment in parsed.attachments:
                    summary.total_attachments += 1
                    if not attachment.filename.lower().endswith(".pdf"):
                        summary.invalid_attachments += 1
                        self.database.insert_invoice(
                            InvoiceRecord(
                                message_uid=parsed.message_uid,
                                message_id=parsed.message_id,
                                sender=parsed.sender,
                                subject=parsed.subject,
                                phone_number=parsed.phone_number,
                                billing_period=parsed.billing_period,
                                settlement_group=parsed.settlement_group,
                                account_name=self.config.email.email_address,
                                attachment_name=attachment.filename,
                                attachment_path="",
                                attachment_size=attachment.size,
                                amount="",
                                received_at=parsed.received_at,
                                status="invalid_attachment",
                                error_message="仅支持 PDF 附件",
                            )
                        )
                        logger("detail", f"跳过附件 {attachment.filename}，原因：仅支持 PDF。")
                        continue
                    try:
                        if self.database.record_exists(parsed.message_uid, attachment.filename):
                            summary.duplicate_records += 1
                            logger("detail", f"重复附件，已跳过 {attachment.filename} (UID={parsed.message_uid})")
                            continue
                        record = self._save_attachment_record(parsed, attachment.filename, attachment.payload, attachment.size)
                        inserted = self.database.insert_invoice(record)
                        if inserted:
                            summary.saved_records += 1
                            logger("detail", f"已保存 {attachment.filename} -> {record.attachment_path}")
                        else:
                            summary.duplicate_records += 1
                            Path(record.attachment_path).unlink(missing_ok=True)
                            logger("detail", f"重复附件，已跳过 {attachment.filename} (UID={parsed.message_uid})")
                    except Exception as exc:
                        summary.failed_records += 1
                        self.database.insert_invoice(
                            InvoiceRecord(
                                message_uid=parsed.message_uid,
                                message_id=parsed.message_id,
                                sender=parsed.sender,
                                subject=parsed.subject,
                                phone_number=parsed.phone_number,
                                billing_period=parsed.billing_period,
                                settlement_group=parsed.settlement_group,
                                account_name=self.config.email.email_address,
                                attachment_name=attachment.filename,
                                attachment_path="",
                                attachment_size=attachment.size,
                                amount="",
                                received_at=parsed.received_at,
                                status="failed",
                                error_message=str(exc),
                            )
                        )
                        logger("error", f"保存附件失败 {attachment.filename}：{exc}")
        summary.skipped_records = summary.duplicate_records + summary.invalid_attachments
        return summary

    def _save_attachment_record(self, parsed, attachment_name: str, payload: bytes, attachment_size: int) -> InvoiceRecord:
        save_root = ensure_dir(self.config.storage.save_root)
        account_name = _safe_folder_name(self.config.email.email_address or "default_account")
        folder_rule = self.config.storage.folder_rule or "{account_name}/{settlement_group}/{billing_month}"
        relative_dir = folder_rule.format(
            account_name=account_name,
            settlement_group=_safe_folder_name(parsed.settlement_group or "未分组"),
            billing_month=parsed.billing_period,
        )
        archive_dir = ensure_dir(save_root / relative_dir)
        display_name = _build_display_filename(parsed.billing_period, parsed.phone_number)
        target = self._build_unique_path(archive_dir, display_name, parsed.message_uid)
        target.write_bytes(payload)
        amount = extract_invoice_amount(payload)
        return InvoiceRecord(
            message_uid=parsed.message_uid,
            message_id=parsed.message_id,
            sender=parsed.sender,
            subject=parsed.subject,
            phone_number=parsed.phone_number,
            billing_period=parsed.billing_period,
            settlement_group=parsed.settlement_group,
            account_name=self.config.email.email_address,
            attachment_name=attachment_name,
            attachment_path=str(target),
            attachment_size=attachment_size,
            amount=amount,
            received_at=parsed.received_at,
            status="saved",
        )

    def _build_unique_path(self, archive_dir: Path, attachment_name: str, message_uid: str) -> Path:
        safe_name = sanitize_filename(attachment_name)
        base = Path(safe_name).stem
        suffix = Path(safe_name).suffix or ".pdf"
        candidate = archive_dir / f"{base}{suffix}"
        if not candidate.exists():
            return candidate
        uid_candidate = archive_dir / f"{base}_{message_uid}{suffix}"
        if not uid_candidate.exists():
            return uid_candidate
        index = 1
        while True:
            indexed = archive_dir / f"{base}_{message_uid}_{index}{suffix}"
            if not indexed.exists():
                return indexed
            index += 1


def _build_display_filename(billing_period: str, phone_number: str) -> str:
    year, month = "", ""
    if billing_period and "-" in billing_period:
        parts = billing_period.split("-", 1)
        year = parts[0]
        month = str(int(parts[1]))
    phone = phone_number or "未知号码"
    return f"{year}年【中国电信】代表号码为{phone}，账期为{month}月的电子发票.pdf"


def _safe_folder_name(value: str) -> str:
    cleaned = sanitize_filename(value.strip())
    return cleaned or "unknown"
