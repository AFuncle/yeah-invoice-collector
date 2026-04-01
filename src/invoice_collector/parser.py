from __future__ import annotations

import re
from datetime import datetime
from email.header import decode_header, make_header

from .models import MailAttachment, MailFilterConfig, ParsedMail, ParsingConfig, SettlementGroupConfig


def decode_mime_text(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def is_target_mail(sender: str, subject: str, filters: MailFilterConfig) -> bool:
    sender_text = sender.lower()
    subject_text = subject.lower()
    sender_ok = not filters.sender_keywords or any(keyword.lower() in sender_text for keyword in filters.sender_keywords)
    subject_ok = not filters.subject_keywords or any(keyword.lower() in subject_text for keyword in filters.subject_keywords)
    return sender_ok and subject_ok


def extract_phone_number(subject: str, parsing: ParsingConfig) -> str:
    matched = re.search(parsing.phone_regex, subject)
    if not matched:
        matched = re.search(r"(?<!\d)(1\d{10})(?!\d)", subject)
    if not matched:
        return ""
    return matched.group(1) if matched.lastindex else matched.group(0)


def extract_billing_period(subject: str, parsing: ParsingConfig, received_at: str) -> str:
    for pattern in parsing.billing_period_patterns:
        matched = re.search(pattern, subject)
        if matched:
            year = matched.group(1)
            month = matched.group(2).zfill(2)
            return f"{year}-{month}"
    month_only = re.search(r"账期为(0?[1-9]|1[0-2])月", subject) or re.search(r"(0?[1-9]|1[0-2])月", subject)
    if month_only:
        month = int(month_only.group(1))
        year = infer_billing_year(month, received_at)
        return f"{year}-{month:02d}"
    raise ValueError(f"无法从主题提取账期: {subject}")


def infer_billing_year(month: int, received_at: str) -> int:
    if received_at:
        try:
            received_dt = datetime.fromisoformat(received_at)
            year = received_dt.year
            if month > received_dt.month:
                year -= 1
            return year
        except ValueError:
            pass
    today = datetime.now()
    year = today.year
    if month > today.month:
        year -= 1
    return year


def resolve_settlement_group(phone_number: str, settlement_groups: SettlementGroupConfig) -> str:
    if not phone_number:
        return settlement_groups.default
    return settlement_groups.mappings.get(phone_number, settlement_groups.default)


def build_parsed_mail(
    message_uid: str,
    message_id: str,
    sender: str,
    subject: str,
    received_at: str,
    attachments: list[MailAttachment],
    parsing: ParsingConfig,
    settlement_groups: SettlementGroupConfig,
) -> ParsedMail:
    phone_number = extract_phone_number(subject, parsing)
    billing_period = extract_billing_period(subject, parsing, received_at)
    return ParsedMail(
        message_uid=message_uid,
        message_id=message_id,
        sender=sender,
        subject=subject,
        received_at=received_at,
        phone_number=phone_number,
        billing_period=billing_period,
        settlement_group=resolve_settlement_group(phone_number, settlement_groups),
        attachments=attachments,
    )
