from __future__ import annotations

import sqlite3
from pathlib import Path

from .models import InvoiceRecord


TABLE_COLUMNS: dict[str, str] = {
    "id": "INTEGER PRIMARY KEY AUTOINCREMENT",
    "message_uid": "TEXT NOT NULL",
    "message_id": "TEXT",
    "sender": "TEXT",
    "subject": "TEXT",
    "phone_number": "TEXT",
    "billing_period": "TEXT",
    "settlement_group": "TEXT",
    "account_name": "TEXT",
    "attachment_name": "TEXT NOT NULL",
    "attachment_path": "TEXT NOT NULL",
    "attachment_size": "INTEGER DEFAULT 0",
    "amount": "TEXT DEFAULT ''",
    "received_at": "TEXT",
    "collected_at": "TEXT",
    "status": "TEXT DEFAULT 'saved'",
    "error_message": "TEXT DEFAULT ''",
}


class InvoiceDatabase:
    def __init__(self, db_path: str | Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self) -> None:
        with self._connect() as conn:
            columns_sql = ", ".join(f"{name} {definition}" for name, definition in TABLE_COLUMNS.items())
            conn.execute(f"CREATE TABLE IF NOT EXISTS invoices ({columns_sql});")
            self._ensure_columns(conn)
            conn.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS uq_invoices_uid_attachment ON invoices (message_uid, attachment_name);"
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_invoices_period_group_status ON invoices (billing_period, settlement_group, status);"
            )
            conn.commit()

    def _ensure_columns(self, conn: sqlite3.Connection) -> None:
        existing = {row["name"] for row in conn.execute("PRAGMA table_info(invoices)")}
        for name, definition in TABLE_COLUMNS.items():
            if name not in existing:
                conn.execute(f"ALTER TABLE invoices ADD COLUMN {name} {definition}")

    def record_exists(self, message_uid: str, attachment_name: str) -> bool:
        sql = "SELECT 1 FROM invoices WHERE message_uid = ? AND attachment_name = ? LIMIT 1"
        with self._connect() as conn:
            return conn.execute(sql, (message_uid, attachment_name)).fetchone() is not None

    def insert_invoice(self, record: InvoiceRecord) -> bool:
        sql = """
        INSERT OR IGNORE INTO invoices (
            message_uid, message_id, sender, subject, phone_number, billing_period,
            settlement_group, account_name, attachment_name, attachment_path, attachment_size, amount,
            received_at, collected_at, status, error_message
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        with self._connect() as conn:
            cursor = conn.execute(
                sql,
                (
                    record.message_uid,
                    record.message_id,
                    record.sender,
                    record.subject,
                    record.phone_number,
                    record.billing_period,
                    record.settlement_group,
                    record.account_name,
                    record.attachment_name,
                    record.attachment_path,
                    record.attachment_size,
                    record.amount,
                    record.received_at,
                    record.collected_at,
                    record.status,
                    record.error_message,
                ),
            )
            conn.commit()
            return cursor.rowcount > 0

    def fetch_invoices(
        self,
        settlement_group: str | None = None,
        billing_period: str | None = None,
        status: str | None = None,
    ) -> list[sqlite3.Row]:
        sql = """
        SELECT
            id, message_uid, message_id, sender, subject, phone_number, billing_period,
            settlement_group, account_name, attachment_name, attachment_path, attachment_size, amount,
            received_at, collected_at, status, error_message
        FROM invoices
        WHERE 1=1
        """
        params: list[str] = []
        if settlement_group and settlement_group != "全部":
            sql += " AND settlement_group = ?"
            params.append(settlement_group)
        if billing_period and billing_period != "全部":
            sql += " AND billing_period = ?"
            params.append(billing_period)
        if status and status != "全部":
            sql += " AND status = ?"
            params.append(status)
        sql += " ORDER BY received_at DESC, id DESC"
        with self._connect() as conn:
            return list(conn.execute(sql, params))

    def fetch_distinct_values(self, column: str) -> list[str]:
        if column not in {"settlement_group", "billing_period", "status"}:
            raise ValueError(f"Unsupported column: {column}")
        sql = f"SELECT DISTINCT {column} FROM invoices WHERE {column} IS NOT NULL AND {column} != '' ORDER BY {column} DESC"
        with self._connect() as conn:
            return [row[0] for row in conn.execute(sql)]

    def get_stats(self) -> dict[str, int]:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT
                    COUNT(*) AS total_records,
                    SUM(CASE WHEN status = 'saved' THEN 1 ELSE 0 END) AS saved_records,
                    SUM(CASE WHEN status = 'duplicate' THEN 1 ELSE 0 END) AS duplicate_records,
                    SUM(CASE WHEN status = 'failed' THEN 1 ELSE 0 END) AS failed_records,
                    COUNT(DISTINCT settlement_group) AS settlement_groups,
                    COUNT(DISTINCT billing_period) AS billing_periods
                FROM invoices
                """
            ).fetchone()
        return {
            "total_records": row["total_records"] or 0,
            "saved_records": row["saved_records"] or 0,
            "duplicate_records": row["duplicate_records"] or 0,
            "failed_records": row["failed_records"] or 0,
            "settlement_groups": row["settlement_groups"] or 0,
            "billing_periods": row["billing_periods"] or 0,
        }

    def clear_all_records(self) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM invoices")
            conn.commit()
