from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from .database import InvoiceDatabase
from .paths import ensure_dir


class ExcelExporter:
    def __init__(self, database: InvoiceDatabase, export_root: str | Path) -> None:
        self.database = database
        self.export_root = ensure_dir(export_root)

    def export_rows(self, rows: list[dict], scope_name: str) -> Path:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Invoices"
        sheet.append(
            [
                "ID",
                "状态",
                "账期",
                "结算组",
                "代表号码",
                "附件名",
                "金额",
                "发件人",
                "主题",
                "接收时间",
                "保存路径",
                "错误信息",
                "采集时间",
            ]
        )
        for row in rows:
            sheet.append(
                [
                    row.get("id", ""),
                    row.get("status", ""),
                    row.get("billing_period", ""),
                    row.get("settlement_group", ""),
                    row.get("phone_number", ""),
                    row.get("attachment_name", ""),
                    row.get("amount", ""),
                    row.get("sender", ""),
                    row.get("subject", ""),
                    row.get("received_at", ""),
                    row.get("attachment_path", ""),
                    row.get("error_message", ""),
                    row.get("collected_at", ""),
                ]
            )
        filename = f"invoices_{scope_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        target = self.export_root / filename
        workbook.save(target)
        return target

    def export_all(self) -> Path:
        rows = [dict(row) for row in self.database.fetch_invoices()]
        return self.export_rows(rows, "all")
