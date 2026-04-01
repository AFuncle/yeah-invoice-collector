from __future__ import annotations

import shutil
import traceback
from dataclasses import asdict
from datetime import datetime, timedelta
from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QAbstractSpinBox,
    QComboBox,
    QDateEdit,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)

from ..collector import InvoiceCollectorService
from ..config import DEFAULT_CONFIG_PATH, default_recent_criteria, load_config, save_config, update_recent_configs
from ..database import InvoiceDatabase
from ..exporter import ExcelExporter
from ..imap_client import ImapMailClient
from ..models import AppConfig
from ..paths import ensure_dir, get_default_save_dir, normalize_path, open_folder


class ConnectionTestWorker(QThread):
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, config: AppConfig) -> None:
        super().__init__()
        self.config = config

    def run(self) -> None:
        try:
            result = ImapMailClient(self.config).test_connection()
            self.finished_ok.emit(
                {
                    "server_ok": result.server_ok,
                    "login_ok": result.login_ok,
                    "folder_ok": result.folder_ok,
                    "readable_count": result.readable_count,
                    "message": result.message,
                }
            )
        except Exception:
            self.failed.emit(traceback.format_exc())


class CollectWorker(QThread):
    log = Signal(str, str)
    progress = Signal(dict)
    finished_ok = Signal(dict)
    failed = Signal(str)

    def __init__(self, config: AppConfig, database: InvoiceDatabase) -> None:
        super().__init__()
        self.config = config
        self.database = database
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def is_stop_requested(self) -> bool:
        return self._stop_requested

    def run(self) -> None:
        try:
            service = InvoiceCollectorService(self.config, self.database)
            summary = service.collect(self.log.emit, self.progress.emit, self.is_stop_requested)
            self.finished_ok.emit(asdict(summary))
        except InterruptedError as exc:
            self.failed.emit(str(exc))
        except Exception:
            self.failed.emit(traceback.format_exc())


class MainWindow(QMainWindow):
    def __init__(self, database: InvoiceDatabase) -> None:
        super().__init__()
        self.database = database
        self.config_path = DEFAULT_CONFIG_PATH
        self.config = load_config(self.config_path)
        self.worker: CollectWorker | None = None
        self.test_worker: ConnectionTestWorker | None = None
        self.current_rows: list[dict] = []
        self.log_entries: list[tuple[str, str]] = []
        self.last_summary = {"matched_mails": 0, "saved_records": 0, "skipped_records": 0, "failed_records": 0}

        self._build_widgets()
        self._apply_adaptive_control_sizes()
        self._setup_ui()
        self._bind_events()
        self._apply_config_to_form()
        self._refresh_results()

        self.resize(self.config.ui_preferences.window_width, self.config.ui_preferences.window_height)
        self.setWindowTitle("邮箱电子发票采集器")

    def _build_widgets(self) -> None:
        self.save_root_toolbar_input = QLineEdit()
        self.email_input = QLineEdit()
        self.email_input.setMinimumWidth(420)
        self.email_input.setMaximumWidth(520)
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setMinimumWidth(420)
        self.password_input.setMaximumWidth(520)
        self.save_root_input = QLineEdit()
        self.server_status_value = QLabel("未测试")
        self.login_status_value = QLabel("未测试")
        self.folder_status_value = QLabel("未测试")
        self.mail_count_value = QLabel("-")
        self.server_status_label = QLabel("服务器连接")
        self.login_status_label = QLabel("账号登录")
        self.folder_status_label = QLabel("邮箱目录")
        self.mail_count_label = QLabel("可读取邮件数")
        self.range_combo = QComboBox()
        self.range_combo.addItems(["最近一个月", "最近三个月", "最近半年", "全部邮件", "自定义"])
        self.range_combo.setMinimumWidth(150)
        self.custom_start_month_input = QDateEdit()
        self.custom_start_month_input.setDisplayFormat("yyyy-MM-dd")
        self.custom_start_month_input.setCalendarPopup(True)
        self.custom_start_month_input.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self.custom_end_month_input = QDateEdit()
        self.custom_end_month_input.setDisplayFormat("yyyy-MM-dd")
        self.custom_end_month_input.setCalendarPopup(True)
        self.custom_end_month_input.setButtonSymbols(QAbstractSpinBox.NoButtons)
        self._configure_date_editors()
        self.host_input = QLineEdit()
        self.port_input = QLineEdit()
        self.folder_input = QLineEdit()
        self.criteria_input = QLineEdit()

        self.choose_save_root_button = QPushButton("选择保存目录")
        self.open_save_root_button = QPushButton("打开保存目录")
        self.test_button = QPushButton("测试连接")
        self.collect_button = QPushButton("开始采集")
        self.export_button = QPushButton("导出 Excel")
        self.clear_cache_button = QPushButton("清理缓存")
        self.export_button.setEnabled(False)

        self.password_toggle = QToolButton()
        self.password_toggle.setText("显示")
        self.password_toggle.setCheckable(True)

        self.advanced_toggle = QToolButton()
        self.advanced_toggle.setText("显示高级设置")
        self.advanced_toggle.setCheckable(True)
        self.advanced_toggle.setChecked(False)
        self.advanced_toggle.setFixedHeight(32)

        self.log_toggle = QToolButton()
        self.log_toggle.setText("展开详细日志")
        self.log_toggle.setCheckable(True)
        self.log_toggle.setChecked(False)
        self.clear_log_button = QPushButton("清理日志")
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["白色主题", "黑色主题"])

        self.progress_label = QLabel("未开始采集")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)

        self.summary_label = QLabel("命中邮件数：0    新增记录数：0    跳过数：0    失败数：0")
        self.group_filter_combo = QComboBox()
        self.group_filter_combo.addItem("全部结算组")
        self.group_filter_combo.setMinimumWidth(160)
        self.month_filter_combo = QComboBox()
        self.month_filter_combo.addItem("全部月份")
        self.month_filter_combo.setMinimumWidth(150)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["账期", "结算组", "代表号码", "附件名", "金额", "收件时间", "状态"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSortingEnabled(False)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.apply_table_column_policy()

        self.table_empty_label = QLabel("暂无采集结果。\n请先完成配置并点击“开始采集”。")
        self.table_empty_label.setAlignment(Qt.AlignCenter)
        self.table_empty_label.setMinimumHeight(260)
        self.result_stack = QStackedWidget()

        self.detail_empty_label = QLabel("请选择一条记录查看详情")
        self.detail_empty_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        self.detail_sender = QLabel("-")
        self.detail_subject = QLabel("-")
        self.detail_path = QLabel("-")
        self.detail_amount = QLabel("-")
        self.detail_status = QLabel("-")
        self.detail_error = QLabel("-")
        for widget in (
            self.detail_sender,
            self.detail_subject,
            self.detail_path,
            self.detail_amount,
            self.detail_status,
            self.detail_error,
        ):
            widget.setWordWrap(True)
            widget.setTextInteractionFlags(Qt.TextSelectableByMouse)
            widget.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        for widget in (
            self.server_status_label,
            self.login_status_label,
            self.folder_status_label,
            self.mail_count_label,
        ):
            widget.setObjectName("statusKeyLabel")

        for widget in (
            self.server_status_value,
            self.login_status_value,
            self.folder_status_value,
            self.mail_count_value,
        ):
            widget.setObjectName("statusValueLabel")

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setVisible(False)

    def _apply_adaptive_control_sizes(self) -> None:
        combo_boxes = [
            self.range_combo,
            self.group_filter_combo,
            self.month_filter_combo,
            self.theme_combo,
        ]
        for combo in combo_boxes:
            combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)
            combo.setMinimumContentsLength(max(len(combo.itemText(i)) for i in range(combo.count())) + 2)
            combo.setMinimumWidth(max(combo.minimumWidth(), combo.sizeHint().width() + 24))

        buttons = [
            self.choose_save_root_button,
            self.open_save_root_button,
            self.test_button,
            self.collect_button,
            self.export_button,
        ]
        for button in buttons:
            button.setMinimumWidth(max(button.minimumWidth(), button.sizeHint().width() + 20))

        tool_buttons = [
            self.password_toggle,
            self.advanced_toggle,
            self.log_toggle,
        ]
        for button in tool_buttons:
            button.setMinimumWidth(max(button.minimumWidth(), button.sizeHint().width() + 16))

    def apply_table_column_policy(self) -> None:
        header = self.table.horizontalHeader()
        from PySide6.QtWidgets import QHeaderView

        fixed_columns = {
            0: 92,   # 账期
            1: 96,   # 结算组
            2: 118,  # 代表号码
            4: 86,   # 金额
            5: 150,  # 收件时间
            6: 84,   # 状态
        }

        for column, width in fixed_columns.items():
            header.setSectionResizeMode(column, QHeaderView.Fixed)
            self.table.setColumnWidth(column, width)

        header.setSectionResizeMode(3, QHeaderView.Stretch)  # 附件名

    def _configure_date_editors(self) -> None:
        from PySide6.QtCore import QDate

        current_day = datetime.now()
        previous_day = current_day - timedelta(days=30)
        self._date_guard_enabled = False
        self.custom_start_month_input.setDate(QDate(previous_day.year, previous_day.month, previous_day.day))
        self.custom_end_month_input.setDate(QDate(current_day.year, current_day.month, current_day.day))
        self.custom_start_month_input.setMinimumWidth(140)
        self.custom_end_month_input.setMinimumWidth(140)
        self._date_guard_enabled = True

    def _setup_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        root.addWidget(self._build_task_bar())
        root.addWidget(self._build_basic_config())
        root.addWidget(self._build_result_area(), 1)
        root.addWidget(self._build_log_area())
        root.setStretch(0, 0)
        root.setStretch(1, 0)
        root.setStretch(2, 1)
        root.setStretch(3, 0)

        self.setCentralWidget(central)
        self._apply_styles()

    def _build_task_bar(self) -> QWidget:
        group = QGroupBox("顶部任务栏")
        layout = QGridLayout(group)

        layout.addWidget(QLabel("保存目录"), 0, 0)
        layout.addWidget(self.save_root_toolbar_input, 0, 1, 1, 4)
        layout.addWidget(self.choose_save_root_button, 0, 5)
        layout.addWidget(self.open_save_root_button, 0, 6)
        layout.addWidget(self.test_button, 0, 7)
        layout.addWidget(self.collect_button, 0, 8)
        layout.addWidget(self.export_button, 0, 9)
        layout.addWidget(QLabel("界面主题"), 0, 10)
        layout.addWidget(self.theme_combo, 0, 11)
        layout.setColumnStretch(1, 1)
        return group

    def _build_basic_config(self) -> QWidget:
        group = QGroupBox("基础配置")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 6, 10, 8)
        layout.setSpacing(6)

        label_width = 76
        field_width = 420

        self.email_input.setMinimumWidth(field_width)
        self.email_input.setMaximumWidth(field_width)
        self.password_input.setMinimumWidth(field_width)
        self.password_input.setMaximumWidth(field_width)
        self.range_combo.setMaximumWidth(180)

        row1 = QHBoxLayout()
        row1.setContentsMargins(0, 0, 0, 0)
        row1.setSpacing(8)
        label_account = QLabel("邮箱账号")
        label_account.setFixedWidth(label_width)
        row1.addWidget(label_account)
        row1.addWidget(self.email_input)
        row1.addStretch(1)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.setContentsMargins(0, 0, 0, 0)
        row2.setSpacing(8)
        label_password = QLabel("授权码")
        label_password.setFixedWidth(label_width)
        row2.addWidget(label_password)
        row2.addWidget(self.password_input)
        row2.addWidget(self.password_toggle)
        row2.addStretch(1)
        layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.setContentsMargins(0, 0, 0, 0)
        row3.setSpacing(8)
        label_range = QLabel("采集范围")
        label_range.setFixedWidth(label_width)
        row3.addWidget(label_range)
        row3.addWidget(self.range_combo)
        row3.addStretch(1)
        layout.addLayout(row3)

        self.custom_date_row_widget = QWidget()
        custom_date_row_layout = QHBoxLayout(self.custom_date_row_widget)
        custom_date_row_layout.setContentsMargins(label_width + 8, 0, 0, 0)
        custom_date_row_layout.setSpacing(8)
        custom_date_row_layout.addWidget(QLabel("自定义日期"))
        custom_date_row_layout.addWidget(QLabel("开始日期"))
        custom_date_row_layout.addWidget(self.custom_start_month_input)
        custom_date_row_layout.addWidget(QLabel("结束日期"))
        custom_date_row_layout.addWidget(self.custom_end_month_input)
        custom_date_row_layout.addStretch(1)
        self.custom_date_row_widget.setVisible(False)
        layout.addWidget(self.custom_date_row_widget)

        row4 = QHBoxLayout()
        row4.setContentsMargins(0, 0, 0, 0)
        row4.setSpacing(8)
        label_status = QLabel("连接状态")
        label_status.setFixedWidth(label_width)
        row4.addWidget(label_status)
        status_container = QWidget()
        status_container.setObjectName("connectionStatusContainer")
        status_layout = QGridLayout(status_container)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setHorizontalSpacing(18)
        status_layout.setVerticalSpacing(4)
        status_layout.addWidget(self.server_status_label, 0, 0)
        status_layout.addWidget(self.server_status_value, 0, 1)
        status_layout.addWidget(self.login_status_label, 0, 2)
        status_layout.addWidget(self.login_status_value, 0, 3)
        status_layout.addWidget(self.folder_status_label, 0, 4)
        status_layout.addWidget(self.folder_status_value, 0, 5)
        status_layout.addWidget(self.mail_count_label, 0, 6)
        status_layout.addWidget(self.mail_count_value, 0, 7)
        status_layout.setColumnStretch(8, 1)
        row4.addWidget(status_container)
        row4.addStretch(1)
        layout.addLayout(row4)

        self.advanced_group = QGroupBox("高级设置")
        self.advanced_group.setContentsMargins(6, 6, 6, 6)
        advanced_form = QFormLayout(self.advanced_group)
        advanced_form.setHorizontalSpacing(12)
        advanced_form.setVerticalSpacing(5)
        advanced_form.addRow("IMAP Host", self.host_input)
        advanced_form.addRow("端口", self.port_input)
        advanced_form.addRow("邮箱目录", self.folder_input)
        advanced_form.addRow("搜索条件", self.criteria_input)
        self.advanced_group.setVisible(False)
        layout.addWidget(self.advanced_group)

        row5 = QHBoxLayout()
        row5.setContentsMargins(0, 0, 0, 0)
        row5.setSpacing(8)
        label_progress = QLabel("当前状态")
        label_progress.setFixedWidth(label_width)
        row5.addWidget(label_progress)
        row5.addWidget(self.progress_label)
        row5.addWidget(self.progress_bar, 1)
        row5.addWidget(self.advanced_toggle)
        layout.addLayout(row5)
        return group

    def _build_result_area(self) -> QWidget:
        group = QGroupBox("采集结果")
        group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout = QHBoxLayout(group)
        layout.setContentsMargins(8, 12, 8, 8)
        layout.setSpacing(8)

        left_panel = QWidget()
        left_panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("按结算组筛选"))
        filter_row.addWidget(self.group_filter_combo)
        filter_row.addSpacing(12)
        filter_row.addWidget(QLabel("按月份筛选"))
        filter_row.addWidget(self.month_filter_combo)
        filter_row.addStretch(1)
        left_layout.addLayout(filter_row)
        self.table_empty_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.result_stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.result_stack.addWidget(self.table_empty_label)
        self.result_stack.addWidget(self.table)
        left_layout.addWidget(self.result_stack, 1)

        right_panel = QGroupBox("详情面板")
        right_panel.setMinimumWidth(300)
        right_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setAlignment(Qt.AlignTop)
        right_layout.addWidget(self.detail_empty_label)
        detail_form = QFormLayout()
        detail_form.setLabelAlignment(Qt.AlignLeft | Qt.AlignTop)
        detail_form.setFormAlignment(Qt.AlignTop | Qt.AlignLeft)
        detail_form.addRow("发件人", self.detail_sender)
        detail_form.addRow("邮件主题", self.detail_subject)
        detail_form.addRow("保存路径", self.detail_path)
        detail_form.addRow("发票金额", self.detail_amount)
        detail_form.addRow("下载状态", self.detail_status)
        detail_form.addRow("错误信息", self.detail_error)
        right_layout.addLayout(detail_form)

        self.result_splitter = QSplitter(Qt.Horizontal)
        self.result_splitter.setChildrenCollapsible(False)
        self.result_splitter.addWidget(left_panel)
        self.result_splitter.addWidget(right_panel)
        self.result_splitter.setStretchFactor(0, 4)
        self.result_splitter.setStretchFactor(1, 1)
        self.result_splitter.setSizes([980, 320])
        self.apply_table_column_policy()
        layout.addWidget(self.result_splitter, 1)
        return group

    def _build_log_area(self) -> QWidget:
        group = QGroupBox("底部日志摘要区")
        layout = QVBoxLayout(group)
        summary_row = QHBoxLayout()
        summary_row.addWidget(self.summary_label)
        summary_row.addStretch(1)
        summary_row.addWidget(self.clear_cache_button)
        summary_row.addWidget(self.clear_log_button)
        summary_row.addWidget(self.log_toggle)
        layout.addLayout(summary_row)
        layout.addWidget(self.log_text)
        return group

    def _bind_events(self) -> None:
        self.choose_save_root_button.clicked.connect(self.choose_save_root)
        self.open_save_root_button.clicked.connect(self.open_save_root)
        self.test_button.clicked.connect(self.test_connection)
        self.collect_button.clicked.connect(self.toggle_collection)
        self.export_button.clicked.connect(self.export_excel)
        self.clear_cache_button.clicked.connect(self.confirm_clear_cache)
        self.password_toggle.toggled.connect(self.toggle_password_visibility)
        self.advanced_toggle.toggled.connect(self.toggle_advanced_settings)
        self.log_toggle.toggled.connect(self.toggle_log_area)
        self.clear_log_button.clicked.connect(self.clear_logs)
        self.theme_combo.currentTextChanged.connect(self.apply_theme)
        self.group_filter_combo.currentTextChanged.connect(self._refresh_results)
        self.month_filter_combo.currentTextChanged.connect(self._refresh_results)
        self.table.itemSelectionChanged.connect(self.update_detail_panel)
        self.range_combo.currentTextChanged.connect(self.on_range_changed)
        self.custom_start_month_input.dateChanged.connect(self.on_custom_date_changed)
        self.custom_end_month_input.dateChanged.connect(self.on_custom_date_changed)

    def _apply_styles(self) -> None:
        self.collect_button.setObjectName("primaryButton")
        self.apply_theme(self.theme_combo.currentText())

    def apply_theme(self, theme_name: str) -> None:
        if theme_name == "黑色主题":
            self.setStyleSheet(self.build_stylesheet("dark"))
            return
        self.setStyleSheet(self.build_stylesheet("light"))

    @staticmethod
    def build_stylesheet(theme_mode: str = "light") -> str:
        if theme_mode == "dark":
            return """
            QWidget {
                background: #1f2329;
                color: #f3f4f6;
                font-size: 13px;
            }

            QMainWindow {
                background: #1f2329;
            }

            QLabel {
                color: #f3f4f6;
                background: transparent;
            }

            QLabel#statusKeyLabel {
                color: #9ca3af;
                font-weight: 500;
            }

            QLabel#statusValueLabel {
                color: #f9fafb;
                font-weight: 700;
            }

            QWidget#connectionStatusContainer {
                background: transparent;
                border: none;
            }

            QGroupBox {
                color: #f3f4f6;
                background: #2b313a;
                border: 1px solid #434c59;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
                font-weight: 600;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 10px;
                top: 2px;
                padding: 0 4px;
                color: #f9fafb;
                background: #2b313a;
            }

            QLineEdit,
            QDateEdit,
            QDateTimeEdit,
            QTextEdit,
            QPlainTextEdit,
            QComboBox,
            QTableWidget,
            QTableView {
                background: #262b33;
                color: #f3f4f6;
                border: 1px solid #4b5563;
                border-radius: 4px;
                selection-background-color: #1d4ed8;
                selection-color: #ffffff;
            }

            QDateEdit,
            QDateTimeEdit {
                padding-right: 28px;
            }

            QDateEdit:focus,
            QDateTimeEdit:focus {
                border: 1px solid #60a5fa;
            }

            QDateEdit::drop-down,
            QDateTimeEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 26px;
                border: none;
                border-left: 1px solid #4b5563;
                background: #262b33;
                border-top-right-radius: 4px;
                border-bottom-right-radius: 4px;
            }

            QDateEdit::drop-down:hover,
            QDateTimeEdit::drop-down:hover,
            QDateEdit::drop-down:focus,
            QDateTimeEdit::drop-down:focus {
                background: #313844;
            }

            QDateEdit::down-arrow,
            QDateTimeEdit::down-arrow {
                image: none;
                width: 0;
                height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #d1d5db;
                margin-right: 8px;
            }

            QLineEdit::placeholder,
            QTextEdit::placeholder,
            QPlainTextEdit::placeholder {
                color: #9ca3af;
            }

            QPushButton {
                min-height: 32px;
                padding: 0 12px;
                color: #f3f4f6;
                background: #2b313a;
                border: 1px solid #4b5563;
                border-radius: 4px;
            }

            QPushButton:hover {
                background: #343b45;
            }

            QPushButton:pressed {
                background: #3b4450;
            }

            QPushButton:disabled {
                background: #2a2f37;
                color: #9ca3af;
                border-color: #3f4652;
            }

            QPushButton#primaryButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-weight: 700;
            }

            QPushButton#primaryButton:hover {
                background: #1d4ed8;
                border-color: #1d4ed8;
            }

            QToolButton {
                color: #f3f4f6;
                background: #2b313a;
                border: 1px solid #4b5563;
                border-radius: 4px;
                padding: 4px 10px;
            }

            QHeaderView::section {
                background: #313844;
                color: #f3f4f6;
                border: none;
                border-right: 1px solid #4b5563;
                border-bottom: 1px solid #4b5563;
                padding: 6px 8px;
                font-weight: 600;
            }

            QTableWidget {
                gridline-color: #3f4652;
                alternate-background-color: #2d333c;
            }

            QTableWidget::item {
                color: #f3f4f6;
                background: #262b33;
                padding: 4px;
            }

            QTableWidget::item:selected {
                background: #1d4ed8;
                color: #ffffff;
            }

            QProgressBar {
                background: #262b33;
                color: #f3f4f6;
                border: 1px solid #4b5563;
                border-radius: 4px;
                text-align: center;
            }

            QProgressBar::chunk {
                background: #2563eb;
                border-radius: 3px;
            }
            """
        return """
        QWidget {
            background: #f5f6f8;
            color: #222222;
            font-size: 13px;
        }

        QMainWindow {
            background: #f5f6f8;
        }

        QLabel {
            color: #222222;
            background: transparent;
        }

        QLabel#statusKeyLabel {
            color: #667085;
            font-weight: 500;
        }

        QLabel#statusValueLabel {
            color: #1f1f1f;
            font-weight: 700;
        }

        QWidget#connectionStatusContainer {
            background: transparent;
            border: none;
        }

        QGroupBox {
            color: #222222;
            background: #ffffff;
            border: 1px solid #d9dde5;
            border-radius: 6px;
            margin-top: 10px;
            padding-top: 10px;
            font-weight: 600;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            left: 10px;
            top: 2px;
            padding: 0 4px;
            color: #1f1f1f;
            background: #ffffff;
        }

        QLineEdit,
        QDateEdit,
        QDateTimeEdit,
        QTextEdit,
        QPlainTextEdit,
        QComboBox,
        QTableWidget,
        QTableView {
            background: #ffffff;
            color: #222222;
            border: 1px solid #cfd6df;
            border-radius: 4px;
            selection-background-color: #dbeafe;
            selection-color: #1f1f1f;
        }

        QDateEdit,
        QDateTimeEdit {
            background: #ffffff;
            color: #222222;
            padding-right: 28px;
        }

        QDateEdit:focus,
        QDateTimeEdit:focus {
            border: 1px solid #60a5fa;
        }

        QDateEdit::drop-down,
        QDateTimeEdit::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 26px;
            border: none;
            border-left: 1px solid #cfd6df;
            background: #f8fafc;
            border-top-right-radius: 4px;
            border-bottom-right-radius: 4px;
        }

        QDateEdit::drop-down:hover,
        QDateTimeEdit::drop-down:hover,
        QDateEdit::drop-down:focus,
        QDateTimeEdit::drop-down:focus {
            background: #eef2f7;
        }

        QDateEdit::down-arrow,
        QDateTimeEdit::down-arrow {
            image: none;
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 6px solid #667085;
            margin-right: 8px;
        }

        QLineEdit:disabled,
        QDateEdit:disabled,
        QDateTimeEdit:disabled,
        QTextEdit:disabled,
        QPlainTextEdit:disabled,
        QComboBox:disabled,
        QTableWidget:disabled,
        QTableView:disabled {
            background: #f1f3f5;
            color: #666666;
        }

        QLineEdit[echoMode="2"] {
            lineedit-password-character: 9679;
        }

        QLineEdit::placeholder,
        QTextEdit::placeholder,
        QPlainTextEdit::placeholder {
            color: #8a8f99;
        }

        QPushButton {
            min-height: 32px;
            padding: 0 12px;
            color: #222222;
            background: #ffffff;
            border: 1px solid #cfd6df;
            border-radius: 4px;
        }

        QPushButton:hover {
            background: #f3f6fb;
        }

        QPushButton:pressed {
            background: #e8edf5;
        }

        QPushButton:disabled {
            background: #f1f3f5;
            color: #8a8f99;
            border-color: #dde2e8;
        }

        QPushButton#primaryButton {
            background: #1769e0;
            color: #ffffff;
            border: 1px solid #1769e0;
            font-weight: 700;
        }

        QPushButton#primaryButton:hover {
            background: #1258bd;
            border-color: #1258bd;
        }

        QPushButton#primaryButton:pressed {
            background: #104ca4;
            border-color: #104ca4;
        }

        QToolButton {
            color: #222222;
            background: #ffffff;
            border: 1px solid #cfd6df;
            border-radius: 4px;
            padding: 4px 10px;
        }

        QToolButton:hover {
            background: #f3f6fb;
        }

        QHeaderView::section {
            background: #f2f4f7;
            color: #222222;
            border: none;
            border-right: 1px solid #dde2e8;
            border-bottom: 1px solid #dde2e8;
            padding: 6px 8px;
            font-weight: 600;
        }

        QTableWidget {
            gridline-color: #e5e7eb;
            alternate-background-color: #fafbfc;
        }

        QTableWidget::item {
            color: #222222;
            background: #ffffff;
            padding: 4px;
        }

        QTableWidget::item:selected {
            background: #dbeafe;
            color: #1f1f1f;
        }

        QCheckBox,
        QRadioButton {
            color: #222222;
            background: transparent;
        }

        QCheckBox:disabled,
        QRadioButton:disabled {
            color: #8a8f99;
        }

        QProgressBar {
            background: #ffffff;
            color: #222222;
            border: 1px solid #cfd6df;
            border-radius: 4px;
            text-align: center;
        }

        QProgressBar::chunk {
            background: #1769e0;
            border-radius: 3px;
        }
        """

    def _apply_config_to_form(self) -> None:
        save_root = self.config.storage.save_root or str(get_default_save_dir())
        self.save_root_toolbar_input.setText(save_root)
        self.save_root_input.setText(save_root)
        self.email_input.setText(self.config.email.email_address)
        self.password_input.setText(self.config.email.auth_code)
        self.host_input.setText(self.config.email.imap_host)
        self.port_input.setText(str(self.config.email.imap_port))
        self.folder_input.setText(self.config.email.mail_folder)
        self.criteria_input.setText(self.config.email.search_criteria or default_recent_criteria(30))
        self.criteria_input.setPlaceholderText("默认采集最近 30 天，例如：SINCE 01-Mar-2026")
        self._apply_range_from_criteria(self.criteria_input.text().strip())

    def _apply_form_to_config(self) -> AppConfig:
        save_root = self.save_root_toolbar_input.text().strip() or self.save_root_input.text().strip() or str(get_default_save_dir())
        self.config.email.email_address = self.email_input.text().strip()
        self.config.email.auth_code = self.password_input.text().strip()
        self.config.email.imap_host = self.host_input.text().strip() or "imap.yeah.net"
        self.config.email.imap_port = int(self.port_input.text().strip() or "993")
        self.config.email.mail_folder = self.folder_input.text().strip() or "INBOX"
        self.config.email.search_criteria = self._criteria_from_range() or self.criteria_input.text().strip() or default_recent_criteria(30)
        self.config.storage.save_root = save_root
        self.config.ui_preferences.last_save_root = save_root
        self.config.ui_preferences.window_width = self.width()
        self.config.ui_preferences.window_height = self.height()
        self.save_root_toolbar_input.setText(save_root)
        return self.config

    def persist_config(self) -> None:
        try:
            self._apply_form_to_config()
            self.config_path = Path(self.config_path or DEFAULT_CONFIG_PATH).expanduser()
            self.config = update_recent_configs(self.config, self.config_path)
            save_config(self.config, self.config_path)
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", f"无法保存配置。\n\n{exc}")

    def choose_save_root(self) -> None:
        start_dir = self.save_root_input.text().strip() or self.config.ui_preferences.last_save_root or str(get_default_save_dir())
        path = QFileDialog.getExistingDirectory(self, "选择发票保存目录", start_dir, QFileDialog.ShowDirsOnly)
        if path:
            self.save_root_toolbar_input.setText(path)
            self.persist_config()

    def open_save_root(self) -> None:
        try:
            target = ensure_dir(self.save_root_toolbar_input.text().strip() or get_default_save_dir())
            open_folder(target)
        except Exception as exc:
            QMessageBox.warning(self, "打开失败", f"无法打开保存目录。\n\n{exc}")

    def toggle_password_visibility(self, checked: bool) -> None:
        self.password_input.setEchoMode(QLineEdit.Normal if checked else QLineEdit.Password)
        self.password_toggle.setText("隐藏" if checked else "显示")

    def toggle_advanced_settings(self, checked: bool) -> None:
        self.advanced_group.setVisible(checked)
        self.advanced_toggle.setText("隐藏高级设置" if checked else "显示高级设置")

    def toggle_log_area(self, checked: bool) -> None:
        self.log_text.setVisible(checked)
        self.log_toggle.setText("收起详细日志" if checked else "展开详细日志")

    def on_range_changed(self, _: str) -> None:
        is_custom = self.range_combo.currentText() == "自定义"
        self.custom_date_row_widget.setVisible(is_custom)
        criteria = self._criteria_from_range()
        if criteria:
            self.criteria_input.setText(criteria)

    def on_custom_date_changed(self) -> None:
        if self.range_combo.currentText() != "自定义":
            return
        if getattr(self, "_date_guard_enabled", True):
            current_day = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            start_day = self._parse_date_value(self._date_editor_value(self.custom_start_month_input))
            end_day = self._parse_date_value(self._date_editor_value(self.custom_end_month_input))
            if (start_day and start_day > current_day) or (end_day and end_day > current_day):
                QMessageBox.warning(self, "日期超出范围", "自定义日期不能超过今天，请重新选择。")
                self._configure_date_editors()
                return
        criteria = self._criteria_from_range()
        if criteria:
            self.criteria_input.setText(criteria)

    def _criteria_from_range(self) -> str:
        text = self.range_combo.currentText()
        if text == "最近一个月":
            return self._build_since_criteria(30)
        if text == "最近三个月":
            return self._build_since_criteria(90)
        if text == "最近半年":
            return self._build_since_criteria(180)
        if text == "全部邮件":
            return "ALL"
        if text == "自定义":
            return self._build_date_range_criteria(
                self._date_editor_value(self.custom_start_month_input),
                self._date_editor_value(self.custom_end_month_input),
            )
        return ""

    def _apply_range_from_criteria(self, criteria: str) -> None:
        normalized = criteria.strip().upper()
        if normalized == "ALL":
            self.range_combo.setCurrentText("全部邮件")
            self.custom_date_row_widget.setVisible(False)
            return
        date_range = self._parse_date_range(criteria)
        if date_range:
            start_date, end_date = date_range
            self.range_combo.setCurrentText("自定义")
            self._set_date_editor_value(self.custom_start_month_input, start_date)
            self._set_date_editor_value(self.custom_end_month_input, end_date)
            self.custom_date_row_widget.setVisible(True)
            return
        days = self._parse_since_days(criteria)
        if days is None:
            self.range_combo.setCurrentText("自定义")
            self.custom_date_row_widget.setVisible(True)
            return
        if days <= 35:
            self.range_combo.setCurrentText("最近一个月")
        elif days <= 100:
            self.range_combo.setCurrentText("最近三个月")
        elif days <= 190:
            self.range_combo.setCurrentText("最近半年")
        else:
            self.range_combo.setCurrentText("自定义")
        self.custom_date_row_widget.setVisible(self.range_combo.currentText() == "自定义")

    def _parse_since_days(self, criteria: str) -> int | None:
        criteria = criteria.strip()
        if not criteria.upper().startswith("SINCE "):
            return None
        date_text = criteria[6:].strip()
        for fmt in ("%d-%b-%Y", "%d-%B-%Y"):
            try:
                dt = datetime.strptime(date_text, fmt)
                return max((datetime.now() - dt).days, 0)
            except ValueError:
                continue
        return None

    def _build_since_criteria(self, days: int) -> str:
        dt = datetime.now() - timedelta(days=days)
        return f"SINCE {dt.strftime('%d-%b-%Y')}"

    def _build_date_range_criteria(self, start_date: str, end_date: str) -> str:
        if not start_date or not end_date:
            return ""
        start_dt = self._parse_date_value(start_date)
        end_dt = self._parse_date_value(end_date)
        if not start_dt or not end_dt:
            return ""
        if start_dt > end_dt:
            start_dt, end_dt = end_dt, start_dt
        next_day = end_dt + timedelta(days=1)
        return f"SINCE {start_dt.strftime('%d-%b-%Y')} BEFORE {next_day.strftime('%d-%b-%Y')}"

    def _parse_date_value(self, value: str) -> datetime | None:
        try:
            return datetime.strptime(value, "%Y-%m-%d")
        except ValueError:
            return None

    def _date_editor_value(self, editor: QDateEdit) -> str:
        return editor.date().toString("yyyy-MM-dd")

    def _set_date_editor_value(self, editor: QDateEdit, value: str) -> None:
        from PySide6.QtCore import QDate

        dt = self._parse_date_value(value)
        if not dt:
            return
        self._date_guard_enabled = False
        editor.setDate(QDate(dt.year, dt.month, dt.day))
        self._date_guard_enabled = True

    def _parse_date_range(self, criteria: str) -> tuple[str, str] | None:
        parts = criteria.strip().split()
        if len(parts) != 4:
            return None
        if parts[0].upper() != "SINCE" or parts[2].upper() != "BEFORE":
            return None
        start_dt = self._parse_imap_date(parts[1])
        before_dt = self._parse_imap_date(parts[3])
        if not start_dt or not before_dt:
            return None
        end_dt = before_dt - timedelta(days=1)
        return start_dt.strftime("%Y-%m-%d"), end_dt.strftime("%Y-%m-%d")

    def _parse_imap_date(self, value: str) -> datetime | None:
        for fmt in ("%d-%b-%Y", "%d-%B-%Y"):
            try:
                return datetime.strptime(value, fmt)
            except ValueError:
                continue
        return None

    def test_connection(self) -> None:
        self._apply_form_to_config()
        if not self.config.email.email_address or not self.config.email.auth_code:
            QMessageBox.warning(self, "配置不完整", "请先填写邮箱账号和授权码。")
            return
        self.persist_config()
        self.test_button.setEnabled(False)
        self.server_status_value.setText("测试中")
        self.login_status_value.setText("测试中")
        self.folder_status_value.setText("测试中")
        self.mail_count_value.setText("-")
        self.append_log("info", "开始测试邮箱连接。")
        self.test_worker = ConnectionTestWorker(self.config)
        self.test_worker.finished_ok.connect(self.on_test_connection_finished)
        self.test_worker.failed.connect(self.on_test_connection_failed)
        self.test_worker.start()

    def on_test_connection_finished(self, result: dict) -> None:
        self.test_button.setEnabled(True)
        self.server_status_value.setText("成功" if result["server_ok"] else "失败")
        self.login_status_value.setText("成功" if result["login_ok"] else "失败")
        self.folder_status_value.setText("成功" if result["folder_ok"] else "失败")
        self.mail_count_value.setText(str(result["readable_count"]) if result["readable_count"] is not None else "-")
        message = (
            f"服务器连接：{self.server_status_value.text()}，"
            f"账号登录：{self.login_status_value.text()}，"
            f"邮箱目录：{self.folder_status_value.text()}，"
            f"可读取邮件数：{self.mail_count_value.text()}"
        )
        self.append_log("info", message)

    def on_test_connection_failed(self, error_text: str) -> None:
        self.test_button.setEnabled(True)
        self.server_status_value.setText("失败")
        self.login_status_value.setText("失败")
        self.folder_status_value.setText("失败")
        self.mail_count_value.setText("-")
        self.append_log("error", error_text)
        QMessageBox.critical(self, "连接测试失败", error_text[-1200:])

    def toggle_collection(self) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.collect_button.setEnabled(False)
            self.progress_label.setText("正在停止采集")
            return
        self.start_collection()

    def start_collection(self) -> None:
        self._apply_form_to_config()
        if not self.config.email.email_address or not self.config.email.auth_code:
            QMessageBox.warning(self, "配置不完整", "请先填写邮箱账号和授权码。")
            return
        self.persist_config()
        try:
            ensure_dir(self.config.storage.save_root)
        except Exception as exc:
            QMessageBox.critical(self, "目录不可用", f"无法创建保存目录。\n\n{exc}")
            return
        self.collect_button.setText("停止采集")
        self.collect_button.setEnabled(True)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFormat("采集中")
        self.progress_label.setText("正在连接邮箱")
        self.append_log("info", "开始执行采集任务。")
        self.worker = CollectWorker(self.config, self.database)
        self.worker.log.connect(self.append_log)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished_ok.connect(self.on_collection_finished)
        self.worker.failed.connect(self.on_collection_failed)
        self.worker.start()

    def update_progress(self, payload: dict) -> None:
        total = payload.get("total", 0)
        current = payload.get("current", 0)
        step = payload.get("step", "处理中")
        if total > 0:
            self.progress_bar.setRange(0, total)
            self.progress_bar.setValue(current)
        else:
            self.progress_bar.setRange(0, 0)
        parts = [step]
        if total > 0:
            parts.append(f"已处理 {current}/{total}")
        if "saved" in payload:
            parts.append(f"新增 {payload['saved']}")
        if "duplicate" in payload:
            parts.append(f"跳过 {payload['duplicate']}")
        if "failed" in payload:
            parts.append(f"失败 {payload['failed']}")
        self.progress_label.setText("    ".join(parts))

    def on_collection_finished(self, summary: dict) -> None:
        self.worker = None
        self.collect_button.setText("开始采集")
        self.collect_button.setEnabled(True)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(1)
        self.progress_bar.setFormat("采集完成")
        self.progress_label.setText("采集完成")
        self.last_summary = summary
        self.update_summary_label()
        self.append_log(
            "info",
            f"采集完成。命中邮件 {summary['matched_mails']}，新增 {summary['saved_records']}，跳过 {summary['skipped_records']}，失败 {summary['failed_records']}。",
        )
        self._refresh_results()

    def on_collection_failed(self, error_text: str) -> None:
        self.worker = None
        self.collect_button.setText("开始采集")
        self.collect_button.setEnabled(True)
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("已停止")
        if "手动停止" in error_text:
            self.progress_label.setText("采集已停止")
            self.append_log("info", error_text)
            return
        self.progress_label.setText("采集失败")
        self.append_log("error", error_text)
        QMessageBox.critical(self, "采集失败", error_text[-1200:])

    def update_summary_label(self) -> None:
        self.summary_label.setText(
            f"命中邮件数：{self.last_summary.get('matched_mails', 0)}    "
            f"新增记录数：{self.last_summary.get('saved_records', 0)}    "
            f"跳过数：{self.last_summary.get('skipped_records', 0)}    "
            f"失败数：{self.last_summary.get('failed_records', 0)}"
        )

    def append_log(self, level: str, message: str) -> None:
        self.log_entries.append((level, message))
        lines = []
        for item_level, item_message in self.log_entries:
            prefix = "[错误]" if item_level == "error" else "[日志]"
            lines.append(f"{prefix} {item_message}")
        self.log_text.setPlainText("\n".join(lines))
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_text.setTextCursor(cursor)

    def clear_logs(self) -> None:
        self.log_entries.clear()
        self.log_text.clear()

    def confirm_clear_cache(self) -> None:
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "提示", "请先停止采集任务，再执行清理缓存。")
            return

        message_box = QMessageBox(self)
        message_box.setWindowTitle("清理缓存")
        message_box.setText("请选择清理方式。")
        message_box.setInformativeText("可清理采集结果缓存、导出文件，以及可选地清理已下载的 PDF 文件。")
        cache_only_button = message_box.addButton("仅清理缓存", QMessageBox.AcceptRole)
        all_button = message_box.addButton("清理缓存和PDF", QMessageBox.DestructiveRole)
        cancel_button = message_box.addButton("取消", QMessageBox.RejectRole)
        message_box.exec()

        clicked = message_box.clickedButton()
        if clicked == cancel_button or clicked is None:
            return
        include_downloads = clicked == all_button
        self.clear_cache(include_downloads=include_downloads)

    def clear_cache(self, include_downloads: bool) -> None:
        try:
            self.database.clear_all_records()
            self._delete_children(Path(self.config.storage.export_root)) if self.config.storage.export_root else None
            if include_downloads and self.config.storage.save_root:
                self._delete_children(Path(self.config.storage.save_root))
            self.last_summary = {
                "matched_mails": 0,
                "saved_records": 0,
                "skipped_records": 0,
                "failed_records": 0,
            }
            self.update_summary_label()
            self.clear_logs()
            self._refresh_results()
            notice = "已清理缓存和已下载文件。" if include_downloads else "已清理缓存。"
            self.append_log("info", notice)
            QMessageBox.information(self, "清理完成", notice)
        except Exception as exc:
            QMessageBox.critical(self, "清理失败", f"清理缓存时出错。\n\n{exc}")

    def _delete_children(self, path: Path) -> None:
        target = path.expanduser()
        if not target.exists():
            return
        for child in target.iterdir():
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=False)
            else:
                child.unlink(missing_ok=True)

    def _refresh_results(self) -> None:
        self.current_rows = sorted(
            [
                dict(row)
                for row in self.database.fetch_invoices(
                    settlement_group=self._selected_group_filter(),
                    billing_period=self._selected_month_filter(),
                )
                if any(
                    [
                        row["billing_period"],
                        row["attachment_name"],
                        row["received_at"],
                        row["status"],
                        row["subject"],
                    ]
                )
            ],
            key=lambda row: (
                row.get("billing_period", "") or "",
                row.get("received_at", "") or "",
                row.get("collected_at", "") or "",
            ),
            reverse=True,
        )
        self.table.setRowCount(len(self.current_rows))
        for row_index, row in enumerate(self.current_rows):
            values = [
                row.get("billing_period", "") or "-",
                row.get("settlement_group", "") or "-",
                row.get("phone_number", "") or "-",
                row.get("attachment_name", "") or "-",
                row.get("amount", "") or "未识别",
                self._display_received_time(row.get("received_at", "") or ""),
                self._display_status(row.get("status", "") or "-"),
            ]
            for column_index, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                item.setToolTip(str(value or ""))
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_index, column_index, item)
        has_rows = bool(self.current_rows)
        self.result_stack.setCurrentWidget(self.table if has_rows else self.table_empty_label)
        self.export_button.setEnabled(has_rows)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self._refresh_filter_options()
        self.table.viewport().update()
        self.update_detail_panel()

    def _refresh_filter_options(self) -> None:
        self._reload_filter_combo(
            self.group_filter_combo,
            "全部结算组",
            self.database.fetch_distinct_values("settlement_group"),
        )
        self._reload_filter_combo(
            self.month_filter_combo,
            "全部月份",
            self.database.fetch_distinct_values("billing_period"),
        )

    def _reload_filter_combo(self, combo: QComboBox, all_label: str, values: list[str]) -> None:
        current = combo.currentText() or all_label
        combo.blockSignals(True)
        combo.clear()
        combo.addItem(all_label)
        combo.addItems(values)
        index = combo.findText(current)
        combo.setCurrentIndex(index if index >= 0 else 0)
        combo.blockSignals(False)

    def _selected_group_filter(self) -> str | None:
        value = self.group_filter_combo.currentText().strip()
        return None if not value or value == "全部结算组" else value

    def _selected_month_filter(self) -> str | None:
        value = self.month_filter_combo.currentText().strip()
        return None if not value or value == "全部月份" else value

    def update_detail_panel(self) -> None:
        row_index = self.table.currentRow()
        if row_index < 0 or row_index >= len(self.current_rows):
            self.detail_empty_label.setVisible(True)
            self.detail_sender.setText("-")
            self.detail_subject.setText("-")
            self.detail_path.setText("-")
            self.detail_amount.setText("-")
            self.detail_status.setText("-")
            self.detail_error.setText("-")
            return
        row = self.current_rows[row_index]
        self.detail_empty_label.setVisible(False)
        self.detail_sender.setText(row.get("sender", "-") or "-")
        self.detail_subject.setText(row.get("subject", "-") or "-")
        self.detail_path.setText(row.get("attachment_path", "-") or "-")
        self.detail_amount.setText(row.get("amount", "") or "未识别")
        self.detail_status.setText(self._display_status(row.get("status", "-") or "-"))
        self.detail_error.setText(row.get("error_message", "-") or "-")

    def _display_status(self, value: str) -> str:
        mapping = {
            "saved": "已下载",
            "downloaded": "已下载",
            "duplicate": "已下载",
            "failed": "未下载",
            "invalid_attachment": "未下载",
        }
        return mapping.get(value, value or "未下载")

    def _display_received_time(self, value: str) -> str:
        if not value:
            return "-"
        try:
            dt = datetime.fromisoformat(value)
            return dt.strftime("%Y-%m-%d %H:%M")
        except ValueError:
            return value

    def export_excel(self) -> None:
        if not self.current_rows:
            return
        try:
            export_root = self.config.storage.export_root or str(Path(self.config.storage.save_root).parent / "exports")
            exporter = ExcelExporter(self.database, export_root)
            output = exporter.export_rows(self.current_rows, "current")
            self.append_log("info", f"Excel 导出完成：{output}")
            QMessageBox.information(self, "导出成功", f"文件已生成：\n{output}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", f"无法导出 Excel。\n\n{exc}")

    def closeEvent(self, event) -> None:
        try:
            self.persist_config()
        except Exception:
            pass
        super().closeEvent(event)
