# InvoiceAuto - Automated Invoice Processing Tool

Automated invoice processing system that supports downloading invoices from email, automatic data extraction, file archiving, Excel report generation, and XERO Bill CSV export.

## Features

- 📧 **Email Auto-Download**: Automatically download invoices and BL attachments from QQ Mail
- 🤖 **AI Data Extraction**: Use DeepSeek API to automatically extract key invoice information
- 📁 **Smart File Classification**: Automatically identify and classify Invoice and BL files
- 📊 **Excel Report Generation**: Automatically generate Excel reports containing invoice information
- 💳 **XERO Bill Export**: Generate XERO-compatible CSV files for accounting import
- 📅 **Smart Due Date Calculation**: Configurable due date rules per supplier (SRTS: ETA+7, others: invoice due date or Invoice Date+30)
- 💱 **Multi-Currency Support**: Automatically detect and extract currency (USD, EUR, CNY, etc.)
- 📦 **Flexible Container Type Recognition**: Recognize various container type formats (40HQ, 40FT High Cube, 20GP, etc.)
- 🖥️ **Graphical Interface**: User-friendly GUI for easy operation
- 👥 **Client Information Verification**: Match and verify client information from Booking List
- 💰 **Automatic Price Lookup**: Automatically match freight prices from Price List

## Project Structure

```
InvoiceAuto/
├── main.py                 # Command-line main program
├── gui_app.py              # Graphical interface program
├── EmailHandler.py         # Email processing module
├── invoice_extractor.py    # Invoice data extraction module
├── PDFClassifier.py        # PDF file classification module
├── report_generator.py     # Report generation module (Internal Booking List & XERO Bill)
├── config_loader.py        # Configuration loading module
├── client_check.py         # Client information verification module
├── price_matcher.py        # Automatic price matching module
├── port_codes.json         # Port code mapping file
├── config.example.ini      # Configuration file template
├── config.ini              # Configuration file (create manually, not committed to Git)
├── requirements.txt        # Project dependencies
└── README.md               # Project documentation
```

## Installation

### 1. Requirements

- Python 3.7+
- Git (for version control)

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Configuration

1. Copy the configuration file template:
   ```bash
   copy config.example.ini config.ini
   ```

2. Edit the `config.ini` file and fill in your configuration information:
   - **Email Account**: QQ Mail address
   - **Email Authorization Code**: QQ Mail authorization code (not login password)
   - **API Key**: DeepSeek API key

⚠️ **Important**: The `config.ini` file contains sensitive information and will not be committed to Git. Please keep this file secure.

## Usage

### Command-Line Mode

```bash
python main.py
```

### Graphical Interface Mode

```bash
python gui_app.py
```

Or use the packaged executable:

```bash
dist/InvoiceAuto_V1.6-Official.exe
```

### GUI Features

The GUI application provides the following features:

- **Save Location Selection**: Choose where to save processed files
- **Booking List (Optional)**: Select a Booking List Excel file for client information verification
- **Price List (Optional)**: Select a Price List Excel file for automatic price lookup
- **Real-time Log Display**: View processing logs in real-time

## Workflow

1. **Initialize Directories**: Create folder structure organized by date
2. **Download Emails**: Download and process attachments from mailbox
3. **File Classification**: Automatically identify Invoice and BL files
4. **Data Extraction**: Use AI to extract key invoice information
5. **File Archiving**: Rename and move files to corresponding folders according to rules
6. **Generate Reports**: Generate Excel reports and running logs
7. **Client Verification** (Optional): Match client information from Booking List
8. **Price Lookup** (Optional): Automatically match freight prices from Price List

## Output Files

The program will generate the following files in the `Download/{date}/` directory:

- `Invoice附件/`: Processed invoice files
- `BL附件/`: Processed BL files
- `info.xlsx`: Excel report containing all invoice data
- `internal_booking_list_{date}.xlsx`: Internal booking list for tracking
- `XERO_Bill_{date}.csv`: XERO-compatible bill import file
- `当日运行清单.xlsx`: Running statistics

### Excel Report Columns

The `info.xlsx` file contains the following columns:

- NO, File Name, FILENO, File No, DATE
- Carrier, Vessel/Voyage
- Loading Port, Loading Port Code, Destination, Destination Code
- ETD, ETA, Receipt, OBL, HBL, MBL
- Item, Quantity, Unit Price, Container Type, Amount
- Booking No, Supplier Name, Due Date, Currency
- Client Name (added after client verification)
- Booking List Position (added after client verification)
- Standard Freight Price (added after price lookup)

### Due Date Calculation Rules

| Supplier | Rule |
|----------|------|
| SRTS | DueDate = ETA + 7 days |
| Others | Use invoice Due Date if available, otherwise Invoice Date + 30 days |

### Supported Container Types

The system recognizes various container type formats:

- **40HQ**: 40HQ, 40HC, 40'HQ, 40FT High Cube, 40FT HC, 45HQ, etc.
- **40GP**: 40GP, 40DC, 40FT, 40FT Standard, FEU, etc.
- **20GP**: 20GP, 20FT, 20FT Standard Container, 20DC, TEU, etc.

## Dependencies

- `requests`: HTTP requests
- `pdfplumber`: PDF text extraction
- `pandas`: Data processing
- `openpyxl`: Excel file operations
- `imap_tools`: Email processing
- `tkinter`: Graphical interface (built-in with Python)

## Notes

1. Ensure that IMAP service is enabled in your email and obtain an authorization code
2. API Key must be valid and have sufficient quota
3. File paths cannot contain special characters
4. It is recommended to regularly backup important data
5. The Booking List should contain columns with keywords: Client, Customer, Cnee, or Consignee
6. The Price List should contain columns: Carrier, POL Code, POD Code, Effective Date, Expiry Date, and price columns (20GP, 40GP, 40HQ)

## Version History

### V1.6 (2026-01-07)
- Added smart Due Date calculation (SRTS: ETA+7, others: invoice Due Date or Invoice Date+30)
- Added multi-currency support (USD, EUR, CNY, etc.)
- Added automatic quantity/price fallback when only total amount is available
- Enhanced container type recognition with flexible pattern matching
- Improved date format recognition (supports "30 Dec 2025", "Voucher Date", etc.)
- Added automatic module reload in GUI to ensure latest code is used

### V1.5
- Added XERO Bill CSV generation
- Added Internal Booking List generation
- Port code auto-matching

### V1.4
- Added client information verification
- Added automatic price lookup
- GUI improvements

## License

This project is for learning and personal use only.

## Contributing

Issues and Pull Requests are welcome!

---

# InvoiceAuto - 发票自动处理工具

自动化发票处理系统，支持从邮箱下载发票、自动提取数据、文件归档、Excel 报表生成和 XERO Bill CSV 导出。

## 功能特性

- 📧 **邮件自动下载**：从 QQ 邮箱自动下载发票和提单附件
- 🤖 **AI 数据提取**：使用 DeepSeek API 自动提取发票关键信息
- 📁 **智能文件分类**：自动识别并分类 Invoice 和 BL 文件
- 📊 **Excel 报表生成**：自动生成包含发票信息的 Excel 报表
- 💳 **XERO Bill 导出**：生成 XERO 兼容的 CSV 文件用于会计导入
- 📅 **智能到期日计算**：按供应商配置到期日规则（SRTS: ETA+7天，其他: 发票到期日或发票日期+30天）
- 💱 **多币种支持**：自动检测和提取币种（USD、EUR、CNY 等）
- 📦 **灵活柜型识别**：识别各种柜型格式（40HQ、40FT High Cube、20GP 等）
- 🖥️ **图形界面**：提供友好的 GUI 界面，方便操作
- 👥 **客户信息核对**：从 Booking List 匹配并验证客户信息
- 💰 **自动查价**：从 Price List 自动匹配运费价格

## 项目结构

```
InvoiceAuto/
├── main.py                 # 命令行主程序
├── gui_app.py              # 图形界面程序
├── EmailHandler.py         # 邮件处理模块
├── invoice_extractor.py    # 发票数据提取模块
├── PDFClassifier.py        # PDF 文件分类模块
├── report_generator.py     # 报表生成模块（Internal Booking List 和 XERO Bill）
├── config_loader.py        # 配置加载模块
├── client_check.py         # 客户信息核对模块
├── price_matcher.py        # 自动查价模块
├── port_codes.json         # 港口代码映射文件
├── config.example.ini      # 配置文件模板
├── config.ini              # 配置文件（需自行创建，不提交到 Git）
├── requirements.txt        # 项目依赖
└── README.md               # 项目说明文档
```

## 安装说明

### 1. 环境要求

- Python 3.7+
- Git（用于版本控制）

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置

1. 复制配置文件模板：
   ```bash
   copy config.example.ini config.ini
   ```

2. 编辑 `config.ini` 文件，填写你的配置信息：
   - **邮箱账号**：QQ 邮箱地址
   - **邮箱授权码**：QQ 邮箱授权码（不是登录密码）
   - **API Key**：DeepSeek API 密钥

⚠️ **重要**：`config.ini` 文件包含敏感信息，不会被提交到 Git。请妥善保管此文件。

## 使用方法

### 命令行模式

```bash
python main.py
```

### 图形界面模式

```bash
python gui_app.py
```

或使用打包后的可执行文件：

```bash
dist/InvoiceAuto_V1.6-Official.exe
```

### GUI 功能

图形界面应用程序提供以下功能：

- **保存位置选择**：选择处理后的文件保存位置
- **Booking List（可选）**：选择 Booking List Excel 文件用于客户信息核对
- **Price List（可选）**：选择 Price List Excel 文件用于自动查价
- **实时日志显示**：实时查看处理日志

## 工作流程

1. **初始化目录**：创建按日期组织的文件夹结构
2. **下载邮件**：从邮箱下载并处理附件
3. **文件分类**：自动识别 Invoice 和 BL 文件
4. **数据提取**：使用 AI 提取发票关键信息
5. **文件归档**：按规则重命名并移动到对应文件夹
6. **生成报表**：生成 Excel 报表和运行清单
7. **客户核对**（可选）：从 Booking List 匹配客户信息
8. **自动查价**（可选）：从 Price List 自动匹配运费价格

## 输出文件

程序会在 `Download/{日期}/` 目录下生成：

- `Invoice附件/`：处理后的发票文件
- `BL附件/`：处理后的提单文件
- `info.xlsx`：包含所有发票数据的 Excel 报表
- `internal_booking_list_{日期}.xlsx`：内部订舱清单
- `XERO_Bill_{日期}.csv`：XERO 兼容的账单导入文件
- `当日运行清单.xlsx`：运行统计信息

### Excel 报表列

`info.xlsx` 文件包含以下列：

- NO, File Name, FILENO, File No, DATE
- Carrier, Vessel/Voyage
- Loading Port, Loading Port Code, Destination, Destination Code
- ETD, ETA, Receipt, OBL, HBL, MBL
- Item, Quantity, Unit Price, Container Type, Amount
- Booking No, Supplier Name, Due Date, Currency
- Client Name（客户核对后添加）
- Booking List Position（客户核对后添加）
- Standard Freight Price（自动查价后添加）

### 到期日计算规则

| 供应商 | 规则 |
|--------|------|
| SRTS | 到期日 = ETA + 7 天 |
| 其他 | 优先使用发票上的到期日，如果没有则使用发票日期 + 30 天 |

### 支持的柜型格式

系统可识别各种柜型格式：

- **40HQ**：40HQ、40HC、40'HQ、40FT High Cube、40FT HC、45HQ 等
- **40GP**：40GP、40DC、40FT、40FT Standard、FEU 等
- **20GP**：20GP、20FT、20FT Standard Container、20DC、TEU 等

## 依赖库

- `requests`：HTTP 请求
- `pdfplumber`：PDF 文本提取
- `pandas`：数据处理
- `openpyxl`：Excel 文件操作
- `imap_tools`：邮件处理
- `tkinter`：图形界面（Python 内置）

## 注意事项

1. 确保邮箱已开启 IMAP 服务并获取授权码
2. API Key 需要有效且有足够的配额
3. 文件路径中不能包含特殊字符
4. 建议定期备份重要数据
5. Booking List 应包含关键词为 Client、Customer、Cnee 或 Consignee 的列
6. Price List 应包含列：Carrier、POL Code、POD Code、Effective Date、Expiry Date 以及价格列（20GP、40GP、40HQ）

## 版本历史

### V1.6 (2026-01-07)
- 新增智能到期日计算（SRTS: ETA+7天，其他: 发票到期日或发票日期+30天）
- 新增多币种支持（USD、EUR、CNY 等）
- 新增自动数量/单价兜底（当只有总价时自动设置数量=1，单价=总价）
- 增强柜型识别，支持灵活的模式匹配
- 改进日期格式识别（支持 "30 Dec 2025"、"Voucher Date" 等格式）
- GUI 启动时自动重载模块以确保使用最新代码

### V1.5
- 新增 XERO Bill CSV 生成
- 新增 Internal Booking List 生成
- 港口代码自动匹配

### V1.4
- 新增客户信息核对功能
- 新增自动查价功能
- GUI 界面优化

## 许可证

本项目仅供学习和个人使用。

## 贡献

欢迎提交 Issue 和 Pull Request！
