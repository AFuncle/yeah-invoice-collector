from __future__ import annotations

import email
import imaplib
from email.message import Message
from email.utils import parsedate_to_datetime

from .models import AppConfig, ConnectionTestResult, MailAttachment
from .parser import decode_mime_text, is_target_mail


class ImapMailClient:
    def __init__(self, config: AppConfig) -> None:
        self.config = config
        self._client: imaplib.IMAP4_SSL | imaplib.IMAP4 | None = None

    def __enter__(self) -> "ImapMailClient":
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def connect(self) -> None:
        email_cfg = self.config.email
        try:
            if email_cfg.use_ssl:
                self._client = imaplib.IMAP4_SSL(email_cfg.imap_host, email_cfg.imap_port)
            else:
                self._client = imaplib.IMAP4(email_cfg.imap_host, email_cfg.imap_port)
        except Exception as exc:
            raise RuntimeError(
                f"无法连接 IMAP 服务器 {email_cfg.imap_host}:{email_cfg.imap_port}，请检查服务器地址、端口和网络。原始错误：{exc}"
            ) from exc

        try:
            self._client.login(email_cfg.email_address, email_cfg.auth_code)
        except imaplib.IMAP4.error as exc:
            raise RuntimeError("IMAP 登录失败，请检查邮箱账号、密码/授权码，以及邮箱是否已开启 IMAP。") from exc

        status, _ = self._client.select(email_cfg.mail_folder)
        if status != "OK":
            raise RuntimeError(f"无法打开邮箱目录：{email_cfg.mail_folder}")

    def close(self) -> None:
        if not self._client:
            return
        try:
            self._client.close()
        except Exception:
            pass
        try:
            self._client.logout()
        except Exception:
            pass
        self._client = None

    def test_connection(self) -> ConnectionTestResult:
        email_cfg = self.config.email
        client: imaplib.IMAP4_SSL | imaplib.IMAP4 | None = None
        server_ok = False
        login_ok = False
        folder_ok = False
        readable_count: int | None = None
        try:
            if email_cfg.use_ssl:
                client = imaplib.IMAP4_SSL(email_cfg.imap_host, email_cfg.imap_port)
            else:
                client = imaplib.IMAP4(email_cfg.imap_host, email_cfg.imap_port)
            server_ok = True
            client.login(email_cfg.email_address, email_cfg.auth_code)
            login_ok = True
            status, _ = client.select(email_cfg.mail_folder)
            folder_ok = status == "OK"
            if folder_ok:
                status, data = client.uid("search", None, "ALL")
                if status == "OK":
                    readable_count = len(data[0].split())
            return ConnectionTestResult(
                server_ok=server_ok,
                login_ok=login_ok,
                folder_ok=folder_ok,
                readable_count=readable_count,
                message="连接测试成功。" if folder_ok else "邮箱目录不可用。",
            )
        except imaplib.IMAP4.error as exc:
            return ConnectionTestResult(server_ok, login_ok, folder_ok, readable_count, f"IMAP 登录失败：{exc}")
        except Exception as exc:
            return ConnectionTestResult(server_ok, login_ok, folder_ok, readable_count, f"连接测试失败：{exc}")
        finally:
            if client:
                try:
                    client.close()
                except Exception:
                    pass
                try:
                    client.logout()
                except Exception:
                    pass

    def fetch_target_messages(self, logger, progress_callback=None, stop_requested=None) -> list[dict]:
        if not self._client:
            raise RuntimeError("IMAP client is not connected")

        criteria = self._build_search_criteria()
        logger("info", f"正在连接 {self.config.email.imap_host}:{self.config.email.imap_port}，目录={self.config.email.mail_folder}")
        logger("info", f"正在搜索邮件，条件：{' '.join(criteria)}")
        status, data = self._client.uid("search", None, *criteria)
        if status != "OK":
            raise RuntimeError(f"IMAP 搜索失败，条件：{' '.join(criteria)}")

        results: list[dict] = []
        uid_list = data[0].split()
        total = len(uid_list)
        logger("info", f"检索到 {total} 封邮件，开始筛选。")
        for index, raw_uid in enumerate(uid_list, start=1):
            if stop_requested and stop_requested():
                raise InterruptedError("采集已手动停止。")
            uid = raw_uid.decode("utf-8", errors="ignore")
            if progress_callback:
                progress_callback(
                    {
                        "step": "正在解析邮件",
                        "current": index,
                        "total": total,
                        "matched": len(results),
                    }
                )
            status, msg_data = self._client.uid("fetch", raw_uid, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                logger("error", f"跳过 UID={uid}，原因：邮件内容读取失败。")
                continue
            raw_message = msg_data[0][1]
            message = email.message_from_bytes(raw_message)
            sender = decode_mime_text(message.get("From"))
            subject = decode_mime_text(message.get("Subject"))
            if not is_target_mail(sender, subject, self.config.mail_filters):
                continue
            results.append(
                {
                    "uid": uid,
                    "message_id": decode_mime_text(message.get("Message-ID")),
                    "sender": sender,
                    "subject": subject,
                    "received_at": self._parse_received_at(message),
                    "attachments": self._extract_attachments(message),
                }
            )
        logger("info", f"筛选后命中 {len(results)} 封中国电信电子发票邮件。")
        return results

    def _build_search_criteria(self) -> list[str]:
        raw = (self.config.email.search_criteria or "ALL").strip()
        return raw.split() if raw else ["ALL"]

    @staticmethod
    def _parse_received_at(message: Message) -> str:
        raw_date = message.get("Date")
        if not raw_date:
            return ""
        try:
            return parsedate_to_datetime(raw_date).isoformat(timespec="seconds")
        except Exception:
            return raw_date

    @staticmethod
    def _extract_attachments(message: Message) -> list[MailAttachment]:
        attachments: list[MailAttachment] = []
        for part in message.walk():
            filename = decode_mime_text(part.get_filename())
            if part.get_content_disposition() != "attachment" or not filename:
                continue
            payload = part.get_payload(decode=True) or b""
            attachments.append(MailAttachment(filename=filename, payload=payload, size=len(payload)))
        return attachments
