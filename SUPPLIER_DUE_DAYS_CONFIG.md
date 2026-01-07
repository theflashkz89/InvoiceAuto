# 供应商 DueDate 配置说明

## 概述

XERO Bill 的 DueDate（到期日）根据供应商类型采用不同的计算规则。

## DueDate 计算规则

### 1. SRTS 供应商

**计算公式**: `DueDate = ETA + 7 天`

- 使用 **ETA**（预计到港日期），注意不是 ETD
- 固定增加 **7 天**
- 如果 ETA 为空，则回退使用 Invoice Date + 30 天

**配置位置**: `report_generator.py` 中的 `SRTS_DUE_DAYS_FROM_ETA` 常量

```python
SRTS_DUE_DAYS_FROM_ETA = 7  # SRTS 供应商使用 ETA + 7 天
```

### 2. 其他供应商

**计算优先级**:
1. **优先使用** invoice 中提取的 Due Date 字段
2. **如果 Due Date 为空**，则使用 Invoice Date + 30 天

**配置位置**: `report_generator.py` 中的 `XERO_DEFAULT_DUE_DAYS` 常量

```python
XERO_DEFAULT_DUE_DAYS = 30  # 默认付款期限（天数）
```

## 数据流说明

1. **数据提取阶段** (`invoice_extractor.py`):
   - AI 从 PDF 发票中提取 `DueDate` 字段（如果发票上有标注）
   - 提取的数据保存到 `info.xlsx` 的 `Due Date` 列

2. **报表生成阶段** (`report_generator.py`):
   - 读取 `info.xlsx`
   - 根据供应商名称判断计算规则
   - 生成 XERO Bill CSV 文件

## 计算逻辑示例

### 示例 1: SRTS 供应商
- ETA: 2026-01-15
- 计算: 2026-01-15 + 7 天
- DueDate: **2026/01/22**

### 示例 2: 其他供应商（发票有 Due Date）
- Invoice Due Date: 2026-02-15
- DueDate: **2026/02/15**（直接使用发票上的日期）

### 示例 3: 其他供应商（发票无 Due Date）
- Invoice Date: 2026-01-10
- Due Date: 空
- 计算: 2026-01-10 + 30 天
- DueDate: **2026/02/09**

## 相关文件

| 文件 | 说明 |
|------|------|
| `invoice_extractor.py` | 从 PDF 提取 DueDate 字段 |
| `main.py` | info.xlsx 表头包含 "Due Date" 列 |
| `report_generator.py` | DueDate 计算逻辑和常量配置 |

## 注意事项

1. 供应商名称匹配是**不区分大小写**的，只要名称中包含 "SRTS" 即视为 SRTS 供应商
2. 日期格式统一为 `YYYY/MM/DD`
3. 如果所有日期字段都为空，DueDate 将为空字符串
