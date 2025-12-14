# InvoiceAuto - 发票自动处理工具

自动化发票处理系统，支持从邮箱下载发票、自动提取数据、文件归档和 Excel 报表生成。

## 功能特性

- 📧 **邮件自动下载**：从 QQ 邮箱自动下载发票和提单附件
- 🤖 **AI 数据提取**：使用 DeepSeek API 自动提取发票关键信息
- 📁 **智能文件分类**：自动识别并分类 Invoice 和 BL 文件
- 📊 **Excel 报表生成**：自动生成包含发票信息的 Excel 报表
- 🖥️ **图形界面**：提供友好的 GUI 界面，方便操作

## 项目结构

```
InvoiceAuto/
├── main.py                 # 命令行主程序
├── gui_app.py             # 图形界面程序
├── EmailHandler.py        # 邮件处理模块
├── invoice_extractor.py   # 发票数据提取模块
├── PDFClassifier.py       # PDF 文件分类模块
├── config_loader.py       # 配置加载模块
├── config.example.ini     # 配置文件模板
├── config.ini            # 配置文件（需自行创建，不提交到 Git）
├── requirements.txt       # 项目依赖
└── README.md             # 项目说明文档
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
dist/gui_app.exe
```

## 工作流程

1. **初始化目录**：创建按日期组织的文件夹结构
2. **下载邮件**：从邮箱下载并处理附件
3. **文件分类**：自动识别 Invoice 和 BL 文件
4. **数据提取**：使用 AI 提取发票关键信息
5. **文件归档**：按规则重命名并移动到对应文件夹
6. **生成报表**：生成 Excel 报表和运行清单

## 输出文件

程序会在 `Download/{日期}/` 目录下生成：

- `Invoice附件/`：处理后的发票文件
- `BL附件/`：处理后的提单文件
- `info.xlsx`：包含所有发票数据的 Excel 报表
- `当日运行清单.xlsx`：运行统计信息

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

## 许可证

本项目仅供学习和个人使用。

## 贡献

欢迎提交 Issue 和 Pull Request！

