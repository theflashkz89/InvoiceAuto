"""
报表生成模块
基于 info.xlsx 生成 internal_booking_list.xlsx 和 XERO_Bill.csv
"""

import os
import re
import pandas as pd
from datetime import datetime, timedelta

# ================= 配置常量 =================
# XERO Bill 相关配置
XERO_DEFAULT_DUE_DAYS = 30  # 默认付款期限（天数），用于 invoice 没有 DueDate 时
XERO_ACCOUNT_CODE = "310"  # XERO 账户代码
XERO_TAX_TYPE = "Tax on Purchases"  # XERO 税务类型
XERO_CURRENCY = "USD"  # XERO 币种
XERO_DATE_FORMAT = "%Y/%m/%d"  # XERO 日期格式

# 供应商 DueDate 特殊配置
# SRTS 供应商：DueDate = ETA + 7 天
SRTS_DUE_DAYS_FROM_ETA = 7

# DueDate 计算规则说明：
# 1. SRTS 供应商：DueDate = ETA + 7 天（使用 ETA，不是 ETD）
# 2. 其他供应商：优先使用 invoice 表中的 Due Date 字段
# 3. 如果 invoice 没有 Due Date，则使用 Invoice Date + 30 天
# ===========================================

# ================= 数据清洗工具函数 =================

def clean_price(value):
    """
    清洗价格字段，移除货币符号和千位分隔符
    
    参数:
        value: 价格值，可能是字符串（如 "$1,000.50"）、数字或 NaN
    
    返回:
        float: 清洗后的价格数值，空值返回 0.0
    
    示例:
        clean_price("$1,000.50") -> 1000.50
        clean_price("1000.50") -> 1000.50
        clean_price(1000.50) -> 1000.50
        clean_price(None) -> 0.0
    """
    if pd.isna(value) or value == '':
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    # 移除货币符号、空格、千位分隔符
    cleaned = re.sub(r'[^\d.]', '', str(value))
    return float(cleaned) if cleaned else 0.0


def safe_join(parts, separator='/'):
    """
    安全拼接字符串，跳过空值和 NaN
    
    参数:
        parts: 字符串列表或可迭代对象
        separator: 分隔符，默认为 '/'
    
    返回:
        str: 拼接后的字符串，空列表返回空字符串
    
    示例:
        safe_join(['FILE123', nan, 'HBL456']) -> 'FILE123/HBL456'
        safe_join(['A', '', 'B']) -> 'A/B'
        safe_join([nan, nan]) -> ''
    """
    valid_parts = [str(p).strip() for p in parts if pd.notna(p) and str(p).strip()]
    return separator.join(valid_parts)


def safe_str(value):
    """
    安全转换为字符串，处理 NaN 和空值
    
    参数:
        value: 要转换的值，可能是任何类型
    
    返回:
        str: 转换后的字符串，NaN 或空值返回空字符串
    
    示例:
        safe_str(123) -> '123'
        safe_str('text') -> 'text'
        safe_str(nan) -> ''
        safe_str(None) -> ''
    """
    if pd.isna(value) or value == '':
        return ''
    return str(value).strip()

# ===========================================

# ================= 柜型标准化函数 =================

def normalize_container_type(raw_type):
    """
    标准化柜型，将各种格式的柜型统一为标准格式
    
    采用灵活的模式匹配，支持各种常见的柜型表示方式。
    
    核心逻辑：
    1. 先识别尺寸（20 或 40 或 45）
    2. 如果是 40/45 尺，检查是否包含"高柜"关键词
    3. 有高柜标识则为 40HQ，否则为 40GP
    4. 20 尺统一为 20GP
    
    参数:
        raw_type: 原始柜型字符串，支持各种格式如：
            - 简写: "40HQ", "40'HC", "20GP", "40DC"
            - 全称: "40FT High Cube", "20FT Standard Container"
            - 带单位: "40 Feet", "20'"
            - 其他变体: "High Cube 40", "HC40", "1x40HQ"
    
    返回:
        str: 标准化后的柜型：
            - "40HQ" - 40/45英尺高柜
            - "40GP" - 40英尺普柜/干柜
            - "20GP" - 20英尺柜
            - "Unknown" - 无法识别的柜型
    """
    if pd.isna(raw_type) or not raw_type:
        return "Unknown"
    
    # 统一转大写，便于匹配
    text = str(raw_type).upper()
    
    # 高柜关键词（任意一个出现即视为高柜）
    high_cube_keywords = ['HQ', 'HC', 'HIGH', 'CUBE', 'HI-CUBE', 'HICUBE']
    
    # 检查是否包含高柜关键词
    is_high_cube = any(kw in text for kw in high_cube_keywords)
    
    # 使用正则提取尺寸数字（支持 20, 40, 45 等）
    size_match = re.search(r'(20|40|45)', text)
    
    if size_match:
        size = size_match.group(1)
        if size in ['40', '45']:
            return "40HQ" if is_high_cube else "40GP"
        elif size == '20':
            return "20GP"
    
    # 没有明确尺寸时，尝试通过关键词推断
    # 如果只写了 "High Cube" 没写尺寸，默认 40HQ
    if is_high_cube:
        return "40HQ"
    
    # TEU 通常指 20 尺
    if 'TEU' in text:
        return "20GP"
    
    # FEU 通常指 40 尺
    if 'FEU' in text:
        return "40GP"
    
    return "Unknown"

# ===========================================

# ================= 供应商名称映射函数 =================

def map_supplier_name(name):
    """
    映射供应商名称到 XERO ContactName
    
    特殊规则：
    - 如果供应商名称中包含 "SRTS"，则映射为 "SRTS Far East Ltd"
    - 其他供应商保持原名（去除首尾空格）
    
    参数:
        name: 原始供应商名称，可能是字符串或 NaN
    
    返回:
        str: 映射后的供应商名称，空值返回空字符串
    
    示例:
        map_supplier_name("SRTS") -> "SRTS Far East Ltd"
        map_supplier_name("SRTS Logistics") -> "SRTS Far East Ltd"
        map_supplier_name("ABC Shipping Co.") -> "ABC Shipping Co."
        map_supplier_name(None) -> ""
    """
    if pd.isna(name) or not name:
        return ""
    name_upper = str(name).strip().upper()
    if 'SRTS' in name_upper:
        return "SRTS Far East Ltd"
    return str(name).strip()

# ===========================================

# ================= 日期处理函数 =================

def format_date(date_value, format_str):
    """
    格式化日期为指定格式
    
    支持多种输入格式：
    - datetime 对象
    - pandas Timestamp
    - 字符串格式的日期（如 "2024/01/15", "2024-01-15"）
    
    参数:
        date_value: 日期值，可能是 datetime、Timestamp、字符串或 NaN
        format_str: 输出格式字符串，如 "%Y/%m/%d"
    
    返回:
        str: 格式化后的日期字符串，空值返回空字符串
    
    示例:
        format_date(datetime(2024, 1, 15), "%Y/%m/%d") -> "2024/01/15"
        format_date("2024-01-15", "%Y/%m/%d") -> "2024/01/15"
        format_date(None, "%Y/%m/%d") -> ""
    """
    if pd.isna(date_value) or date_value == '':
        return ''
    
    # 如果已经是 datetime 对象
    if isinstance(date_value, datetime):
        try:
            return date_value.strftime(format_str)
        except (ValueError, AttributeError):
            return ''
    
    # 如果是 pandas Timestamp
    if isinstance(date_value, pd.Timestamp):
        try:
            return date_value.strftime(format_str)
        except (ValueError, AttributeError):
            return ''
    
    # 如果是字符串，尝试解析
    if isinstance(date_value, str):
        date_str = date_value.strip()
        if not date_str:
            return ''
        
        # 尝试多种日期格式解析
        date_formats = [
            "%Y/%m/%d",
            "%Y-%m-%d",
            "%Y.%m.%d",
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%m/%d/%Y",
            "%m-%d-%Y",
        ]
        
        for fmt in date_formats:
            try:
                parsed_date = datetime.strptime(date_str, fmt)
                return parsed_date.strftime(format_str)
            except (ValueError, TypeError):
                continue
        
        # 如果所有格式都失败，返回空字符串
        return ''
    
    # 其他类型，尝试转换为字符串
    try:
        return str(date_value)
    except Exception:
        return ''


def calculate_due_date(invoice_date, days):
    """
    计算到期日（发票日期 + 指定天数）
    
    参数:
        invoice_date: 发票日期，可能是 datetime、Timestamp、字符串或 NaN
        days: 付款期限（天数），整数
    
    返回:
        str: 格式化后的到期日字符串（使用 XERO_DATE_FORMAT），空值返回空字符串
    
    示例:
        calculate_due_date("2024/01/15", 45) -> "2024/03/01"
        calculate_due_date(datetime(2024, 1, 15), 45) -> "2024/03/01"
        calculate_due_date(None, 45) -> ""
    """
    if pd.isna(invoice_date) or invoice_date == '':
        return ''
    
    # 先尝试解析日期
    date_obj = None
    
    # 如果已经是 datetime 对象
    if isinstance(invoice_date, datetime):
        date_obj = invoice_date
    # 如果是 pandas Timestamp
    elif isinstance(invoice_date, pd.Timestamp):
        date_obj = invoice_date.to_pydatetime()
    # 如果是字符串，尝试解析
    elif isinstance(invoice_date, str):
        date_str = invoice_date.strip()
        if not date_str:
            return ''
        
        # 尝试多种日期格式解析
        date_formats = [
            "%Y/%m/%d",
            "%Y-%m-%d",
            "%Y.%m.%d",
            "%d/%m/%Y",
            "%d-%m-%Y",
            "%m/%d/%Y",
            "%m-%d-%Y",
        ]
        
        for fmt in date_formats:
            try:
                date_obj = datetime.strptime(date_str, fmt)
                break
            except (ValueError, TypeError):
                continue
        
        if date_obj is None:
            return ''
    else:
        return ''
    
    # 计算到期日
    try:
        due_date = date_obj + timedelta(days=days)
        return due_date.strftime(XERO_DATE_FORMAT)
    except Exception:
        return ''

# ===========================================

# ================= Internal Booking List 生成器 =================

def generate_internal_booking_list(info_path, output_path):
    """
    生成 Internal Booking List Excel 文件
    
    参数:
        info_path: info.xlsx 文件路径
        output_path: 输出文件路径（internal_booking_list_YYYYMMDD.xlsx）
    
    返回:
        bool: 成功返回 True，失败返回 False
    """
    try:
        print(f"正在读取 info.xlsx: {info_path}")
        
        # 读取 info.xlsx
        if not os.path.exists(info_path):
            print(f"错误: info.xlsx 文件不存在: {info_path}")
            return False
        
        df_info = pd.read_excel(info_path, engine='openpyxl')
        
        if df_info.empty:
            print("警告: info.xlsx 文件为空")
            return False
        
        print(f"成功读取 {len(df_info)} 行数据")
        
        # 定义 Internal Booking List 表头
        headers = [
            'Type', 'From/to', 'todo', 'check rate', 'POL', 'POD', 'Carrier for SRTS', 
            'File no', 'MBL', 'HBLs', 'Booking number matching', 'ETD/atd', 'ETA', 
            'Customer', 'INV-Number', '# 20ft', '# 40ft', '# 40ft hq', 
            'price 20ft', 'price 40ft', 'price 40ft hq', 'ETS 20ft', 'ETS 40ft', 
            'ETS HQ', 'Other charge per container', 'comment on other charge', 'total',
            '# 20ft.1', '# 40ft.1', '# 40ft hq.1', 'price 20ft.1', 'price 40ft.1', 
            'price 40ft hq.1', 'ets 20ft', 'ets 40ft', 'ets hq', 
            'Other charge per container.1', 'Comment on other charge', 'Total', 
            'Difference', 'JC check'
        ]
        
        # 转换数据行
        rows = []
        for idx, row in df_info.iterrows():
            # 标准化柜型
            container_type = normalize_container_type(safe_str(row.get('Container Type', '')))
            
            # 获取数量和单价
            quantity = safe_str(row.get('Quantity', ''))
            unit_price = clean_price(row.get('Unit Price', 0))
            total_amount = clean_price(row.get('Amount', 0))
            
            # 特殊处理：如果没有数量和单价但有总价，则数量设为1，单价用总价
            if (not quantity or quantity == '0') and unit_price == 0 and total_amount > 0:
                quantity = '1'
                unit_price = total_amount
            
            # 根据柜型分列数量和价格
            qty_20ft = ''
            qty_40ft = ''
            qty_40ft_hq = ''
            price_20ft = ''
            price_40ft = ''
            price_40ft_hq = ''
            
            if container_type == '20GP':
                qty_20ft = quantity
                price_20ft = unit_price if unit_price > 0 else ''
            elif container_type == '40GP':
                qty_40ft = quantity
                price_40ft = unit_price if unit_price > 0 else ''
            elif container_type == '40HQ':
                qty_40ft_hq = quantity
                price_40ft_hq = unit_price if unit_price > 0 else ''
            
            # 计算总价
            try:
                qty_num = float(quantity) if quantity else 0
                total = unit_price * qty_num if qty_num > 0 and unit_price > 0 else ''
            except (ValueError, TypeError):
                total = ''
            
            # 构建数据行（使用容错字段获取，支持多种列名变体）
            data_row = {
                'Type': 'Bill',
                'From/to': safe_str(row.get('Supplier Name', '')),
                'todo': '',
                'check rate': 'Checked/correct',
                'POL': safe_str(row.get('Loading Port', '')),
                'POD': safe_str(row.get('Destination', '')),
                'Carrier for SRTS': safe_str(row.get('Carrier', '')),
                'File no': safe_str(row.get('File No', row.get('FILENO', ''))),  # 优先使用 File No，否则使用 FILENO
                'MBL': safe_str(row.get('OBL', '')),  # 修复：MBL 列应映射到 OBL 字段
                'HBLs': safe_str(row.get('HBL', '')),
                'Booking number matching': safe_str(row.get('Booking No', row.get('Booking No.', ''))),
                'ETD/atd': safe_str(row.get('ETD', '')),
                'ETA': safe_str(row.get('ETA', '')),
                'Customer': safe_str(row.get('Client Name', '')),  # 如果字段不存在则留空
                'INV-Number': '',  # 未来功能：与 XERO Sales Invoice 匹配
                '# 20ft': qty_20ft,
                '# 40ft': qty_40ft,
                '# 40ft hq': qty_40ft_hq,
                'price 20ft': price_20ft,
                'price 40ft': price_40ft,
                'price 40ft hq': price_40ft_hq,
                'ETS 20ft': '',
                'ETS 40ft': '',
                'ETS HQ': '',
                'Other charge per container': '',
                'comment on other charge': '',
                'total': total,
                '# 20ft.1': '',  # Selling 部分全部留空
                '# 40ft.1': '',
                '# 40ft hq.1': '',
                'price 20ft.1': '',
                'price 40ft.1': '',
                'price 40ft hq.1': '',
                'ets 20ft': '',
                'ets 40ft': '',
                'ets hq': '',
                'Other charge per container.1': '',
                'Comment on other charge': '',
                'Total': '',
                'Difference': '',
                'JC check': ''
            }
            rows.append(data_row)
        
        # 创建 DataFrame
        df_output = pd.DataFrame(rows, columns=headers)
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 保存到 Excel
        df_output.to_excel(output_path, index=False, engine='openpyxl')
        print(f"✓ 已生成 Internal Booking List: {output_path}")
        print(f"  共生成 {len(rows)} 行数据")
        
        return True
        
    except Exception as e:
        print(f"✗ 生成 Internal Booking List 失败: {e}")
        import traceback
        traceback.print_exc()
        return False

# ===========================================

# ================= XERO Bill CSV 生成器 =================

def generate_xero_bill(info_path, output_path, due_days=None):
    """
    生成 XERO Bill CSV 文件
    
    DueDate 计算规则：
    1. SRTS 供应商：DueDate = ETA + 7 天
    2. 其他供应商：优先使用 info.xlsx 中的 Due Date 字段
    3. 如果 Due Date 为空，则使用 Invoice Date + 30 天
    
    参数:
        info_path: info.xlsx 文件路径
        output_path: 输出文件路径（XERO_Bill_YYYYMMDD.csv）
        due_days: 付款期限（天数），默认使用 XERO_DEFAULT_DUE_DAYS（仅用于兜底）
    
    返回:
        bool: 成功返回 True，失败返回 False
    """
    try:
        if due_days is None:
            due_days = XERO_DEFAULT_DUE_DAYS
        
        print(f"正在读取 info.xlsx: {info_path}")
        
        # 读取 info.xlsx
        if not os.path.exists(info_path):
            print(f"错误: info.xlsx 文件不存在: {info_path}")
            return False
        
        df_info = pd.read_excel(info_path, engine='openpyxl')
        
        if df_info.empty:
            print("警告: info.xlsx 文件为空")
            return False
        
        print(f"成功读取 {len(df_info)} 行数据")
        
        # 定义 XERO CSV 表头
        headers = [
            '*ContactName', 'EmailAddress', 'POAddressLine1', 'POAddressLine2', 
            'POAddressLine3', 'POAddressLine4', 'POCity', 'PORegion', 
            'POPostalCode', 'POCountry', '*InvoiceNumber', '*InvoiceDate', 
            '*DueDate', 'Total', 'InventoryItemCode', 'Description', 
            '*Quantity', '*UnitAmount', '*AccountCode', '*TaxType', 
            'TaxAmount', 'TrackingName1', 'TrackingOption1', 'TrackingName2', 
            'TrackingOption2', 'Currency'
        ]
        
        # 转换数据行
        rows = []
        for idx, row in df_info.iterrows():
            # 供应商名称映射
            supplier_name = map_supplier_name(row.get('Supplier Name', ''))
            
            # 构建 InvoiceNumber: {File No}/{OBL}/{HBL}（跳过空值）
            # 使用 File No 列（OriginalFileNo 文件编号）
            fileno = safe_str(row.get('File No', ''))
            obl = safe_str(row.get('OBL', ''))
            hbl = safe_str(row.get('HBL', ''))
            invoice_number = safe_join([fileno, obl, hbl], '/')
            
            # 格式化日期（容错处理：支持多种日期列名）
            date_value = row.get('DATE', row.get('Date', row.get('Invoice Date', '')))
            invoice_date = format_date(date_value, XERO_DATE_FORMAT)
            
            # DueDate 计算逻辑：
            # 1. SRTS 供应商：DueDate = ETA + 7 天
            # 2. 其他供应商：优先使用 info.xlsx 中的 Due Date 字段
            # 3. 如果 Due Date 为空，则使用 Invoice Date + 30 天
            
            supplier_name_raw = safe_str(row.get('Supplier Name', '')).upper()
            
            # 判断是否为 SRTS 供应商
            if 'SRTS' in supplier_name_raw:
                # SRTS 供应商：使用 ETA + 7 天
                eta_value = row.get('ETA', '')
                due_date = calculate_due_date(eta_value, SRTS_DUE_DAYS_FROM_ETA)
                if not due_date:
                    # 如果 ETA 为空，尝试使用 Invoice Date + 30 作为兜底
                    due_date = calculate_due_date(date_value, XERO_DEFAULT_DUE_DAYS)
            else:
                # 其他供应商：优先使用 invoice 提取的 Due Date
                invoice_due_date = row.get('Due Date', '')
                
                if pd.notna(invoice_due_date) and safe_str(invoice_due_date):
                    # 使用 invoice 中的 Due Date
                    due_date = format_date(invoice_due_date, XERO_DATE_FORMAT)
                else:
                    # 如果没有 Due Date，使用 Invoice Date + 30 天
                    due_date = calculate_due_date(date_value, XERO_DEFAULT_DUE_DAYS)
            
            # 清洗单价和数量
            unit_amount = clean_price(row.get('Unit Price', 0))
            quantity = safe_str(row.get('Quantity', ''))
            total_amount = clean_price(row.get('Amount', 0))
            
            # 特殊处理：如果没有单价但有总价，则数量设为1，单价用总价
            # 这样 XERO 系统才能正确识别这条记录
            if unit_amount == 0 and total_amount > 0:
                unit_amount = total_amount
                quantity = '1'
            
            # 构建数据行
            data_row = {
                '*ContactName': supplier_name,
                'EmailAddress': '',
                'POAddressLine1': '',
                'POAddressLine2': '',
                'POAddressLine3': '',
                'POAddressLine4': '',
                'POCity': '',
                'PORegion': '',
                'POPostalCode': '',
                'POCountry': '',
                '*InvoiceNumber': invoice_number,
                '*InvoiceDate': invoice_date,
                '*DueDate': due_date,
                'Total': '',  # XERO 会自动计算
                'InventoryItemCode': '',
                'Description': safe_str(row.get('Item', row.get('Description', row.get('Fee Name', '')))),  # 容错：优先 Item，否则 Description 或 Fee Name
                '*Quantity': quantity if quantity else '',
                '*UnitAmount': unit_amount if unit_amount > 0 else '',
                '*AccountCode': XERO_ACCOUNT_CODE,
                '*TaxType': XERO_TAX_TYPE,
                'TaxAmount': '0',
                'TrackingName1': '',
                'TrackingOption1': '',
                'TrackingName2': '',
                'TrackingOption2': '',
                'Currency': safe_str(row.get('Currency', '')) or XERO_CURRENCY  # 优先使用发票币种，否则默认 USD
            }
            rows.append(data_row)
        
        # 创建 DataFrame
        df_output = pd.DataFrame(rows, columns=headers)
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 使用 utf-8-sig 编码保存 CSV（带 BOM，防止乱码）
        df_output.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"✓ 已生成 XERO Bill CSV: {output_path}")
        print(f"  共生成 {len(rows)} 行数据")
        print(f"  付款期限: {due_days} 天")
        
        return True
        
    except Exception as e:
        print(f"✗ 生成 XERO Bill CSV 失败: {e}")
        import traceback
        traceback.print_exc()
        return False

# ===========================================

# ================= 统一入口函数 =================

def generate_all_reports(info_path):
    """
    一键生成所有报表（Internal Booking List 和 XERO Bill）
    
    参数:
        info_path: info.xlsx 文件路径
    
    返回:
        dict: 生成结果字典，包含：
            - success (bool): 是否成功
            - output_dir (str): 输出目录路径
            - files (list): 生成的文件路径列表
            - error (str): 错误信息（如有）
    
    示例:
        result = generate_all_reports('path/to/info.xlsx')
        if result['success']:
            print(f"生成成功，文件保存在: {result['output_dir']}")
            for file in result['files']:
                print(f"  - {file}")
        else:
            print(f"生成失败: {result['error']}")
    """
    result = {
        'success': False,
        'output_dir': '',
        'files': [],
        'error': ''
    }
    
    try:
        # 检查输入文件是否存在
        if not os.path.exists(info_path):
            result['error'] = f"info.xlsx 文件不存在: {info_path}"
            return result
        
        # 获取输出目录（与 info.xlsx 同目录）
        output_dir = os.path.dirname(os.path.abspath(info_path))
        result['output_dir'] = output_dir
        
        # 生成日期后缀（YYYYMMDD 格式）
        date_suffix = datetime.now().strftime("%Y%m%d")
        
        # 生成输出文件路径
        internal_booking_list_path = os.path.join(
            output_dir, 
            f"internal_booking_list_{date_suffix}.xlsx"
        )
        xero_bill_path = os.path.join(
            output_dir,
            f"XERO_Bill_{date_suffix}.csv"
        )
        
        print("=" * 60)
        print("开始生成报表...")
        print("=" * 60)
        print(f"输入文件: {info_path}")
        print(f"输出目录: {output_dir}")
        print()
        
        # 生成 Internal Booking List
        print("【步骤 1】生成 Internal Booking List...")
        success_ibl = generate_internal_booking_list(info_path, internal_booking_list_path)
        
        if success_ibl:
            result['files'].append(internal_booking_list_path)
            print("✓ Internal Booking List 生成成功\n")
        else:
            result['error'] = "Internal Booking List 生成失败"
            print("✗ Internal Booking List 生成失败\n")
            return result
        
        # 生成 XERO Bill CSV
        print("【步骤 2】生成 XERO Bill CSV...")
        success_xero = generate_xero_bill(info_path, xero_bill_path, XERO_DEFAULT_DUE_DAYS)
        
        if success_xero:
            result['files'].append(xero_bill_path)
            print("✓ XERO Bill CSV 生成成功\n")
        else:
            result['error'] = "XERO Bill CSV 生成失败"
            print("✗ XERO Bill CSV 生成失败\n")
            return result
        
        # 所有报表生成成功
        result['success'] = True
        print("=" * 60)
        print("所有报表生成完成！")
        print("=" * 60)
        print(f"输出目录: {output_dir}")
        print(f"生成文件:")
        for file in result['files']:
            print(f"  - {os.path.basename(file)}")
        print()
        
        return result
        
    except Exception as e:
        result['error'] = f"生成报表时发生异常: {str(e)}"
        print(f"✗ {result['error']}")
        import traceback
        traceback.print_exc()
        return result

# ===========================================
