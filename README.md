# Yeah.net 中国电信电子发票采集器

## 项目目录结构

```text
邮件采集/
├── attachments/                 # 附件归档根目录
├── config/
│   └── config.example.json      # 配置文件示例
├── data/
│   └── invoices.db              # SQLite 数据库（运行后生成）
├── exports/                     # Excel 导出目录
├── logs/                        # 日志目录
├── src/
│   └── invoice_collector/
│       ├── ui/
│       │   └── main_window.py   # PySide6 主界面
│       ├── __init__.py
│       ├── app.py               # 应用组装入口
│       ├── collector.py         # 采集流程编排
│       ├── config.py            # 配置加载与保存
│       ├── database.py          # SQLite 封装
│       ├── exporter.py          # Excel 导出
│       ├── imap_client.py       # IMAP 读取
│       ├── models.py            # 数据模型
│       ├── parser.py            # 邮件筛选与主题解析
│       └── paths.py             # 路径常量
├── main.py                      # 启动入口
├── requirements.txt             # 依赖清单
└── .gitignore
```

## 配置文件示例

配置文件采用 JSON，采集规则、归档路径、结算组映射都在配置中维护，不写死在代码里。

见 [config/config.example.json](/Users/afuncle/Desktop/邮件采集/config/config.example.json)

## 数据库设计

SQLite 数据库默认位于 `data/invoices.db`，核心表结构如下：

### `invoices`

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| id | INTEGER PK | 自增主键 |
| message_uid | TEXT UNIQUE | IMAP UID，防重复 |
| message_id | TEXT | 邮件 Message-ID |
| sender | TEXT | 发件人 |
| subject | TEXT | 邮件主题 |
| phone_number | TEXT | 主题提取出的 11 位号码 |
| billing_period | TEXT | 账期，统一为 `YYYY-MM` |
| settlement_group | TEXT | 结算组 |
| attachment_name | TEXT | 原始附件名 |
| attachment_path | TEXT | 本地归档路径 |
| attachment_size | INTEGER | 附件字节数 |
| received_at | TEXT | 邮件接收时间 |
| collected_at | TEXT | 采集入库时间 |
| status | TEXT | 记录状态，默认 `downloaded` |

### 索引设计

- `UNIQUE(message_uid, attachment_name)`：避免重复下载同一封邮件中的同名附件
- `INDEX idx_invoices_period_group`：按结算组、账期查询更快

## MVP 范围

- yeah.net IMAP SSL 登录
- 按配置筛选中国电信电子发票邮件
- 从主题解析号码和账期
- 根据配置映射结算组
- 附件下载并按 `settlement_group / billing_period` 归档
- SQLite 入库
- Excel 导出
- PySide6 界面：邮箱配置、采集按钮、日志区、发票表格

## 筛选范围与速度

- 当前配置文件默认 `search_criteria` 为 `ALL`，表示扫描收件箱全部邮件。
- 如果历史邮件很多，首次采集会比较慢，这是 IMAP 逐封检索带来的正常现象。
- 现在界面已支持实时进度显示，会显示“第 N / 总数”。
- 建议把搜索条件改成最近时间范围，例如：

```json
"search_criteria": "SINCE 1-Mar-2026"
```

- `SINCE` 的日期格式需为 IMAP 标准格式，例如 `1-Mar-2026`、`15-Feb-2026`。

## 主题解析规则

- 号码支持从中文主题中提取，例如 `代表号码为19120076109，账期为08月的电子发票`。
- 账期优先解析完整年月。
- 如果主题里只有月份，例如 `08月`，程序会按邮件接收日期推断年份，并统一保存为 `YYYY-MM`。

## 运行

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp config/config.example.json config/config.json
python main.py
```

## 打包 exe

建议使用 PyInstaller：

```bash
pyinstaller -F -w -n YeahInvoiceCollector main.py
```
