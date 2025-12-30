import pandas as pd
import os
from datetime import datetime, date


class FreightMatcher:
    """
    运费匹配器类，用于加载和处理价格列表数据
    """
    
    def __init__(self):
        """
        初始化 FreightMatcher 实例
        """
        self.price_list = None
    
    def load_price_list(self, excel_path):
        """
        读取 Price List Excel 文件，合并所有 Sheet 的数据并进行数据清洗
        
        参数:
            excel_path: Excel 文件路径
        
        返回:
            pandas.DataFrame: 处理后的价格列表数据
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"文件不存在: {excel_path}")
        
        print(f"正在读取 Price List 文件: {excel_path}")
        
        # 读取所有 Sheet 的数据
        excel_file = pd.ExcelFile(excel_path)
        all_sheets = []
        
        for sheet_name in excel_file.sheet_names:
            print(f"  正在读取 Sheet: {sheet_name}")
            
            # 读取时，强制将第1列（Month列）作为字符串处理
            # 使用 dtype 参数指定第1列为字符串类型
            df = pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                dtype={0: str}  # 第1列（索引0）强制为字符串
            )
            
            # 如果 DataFrame 不为空，添加到列表中
            if not df.empty:
                all_sheets.append(df)
        
        if not all_sheets:
            raise ValueError("Excel 文件中没有有效的数据")
        
        # 合并所有 Sheet 的数据
        self.price_list = pd.concat(all_sheets, ignore_index=True)
        print(f"合并完成，共 {len(self.price_list)} 行数据")
        
        # 数据标准化处理
        self._clean_column_names()  # 先清洗表头
        self._clean_data()  # 再处理数据内容
        
        return self.price_list
    
    def _clean_column_names(self):
        """
        表头清洗：去除列名中的空格，特别是 20GP, 40GP, 40HQ 等列名
        
        处理规则：
        1. 所有列名去除首尾空格
        2. 如果列名包含 20GP, 40GP, 40HQ，则去除所有空格（包括中间的空格）
        """
        if self.price_list is None or self.price_list.empty:
            return
        
        print("正在清洗表头...")
        
        # 创建列名映射
        new_columns = {}
        for col in self.price_list.columns:
            original_col = col
            cleaned_col = str(col).strip()  # 去除首尾空格
            
            # 特别处理 20GP, 40GP, 40HQ 等列名
            # 如果列名包含这些关键词，确保去除所有空格（防止 Excel 中不小心敲了空格）
            col_upper = cleaned_col.upper()
            if '20GP' in col_upper or '40GP' in col_upper or '40HQ' in col_upper:
                # 去除所有空格（包括中间的空格），确保列名紧凑
                cleaned_col = cleaned_col.replace(' ', '').replace('　', '')  # 普通空格和全角空格
            
            # 如果列名有变化，记录映射关系
            if cleaned_col != original_col:
                new_columns[original_col] = cleaned_col
                print(f"  ✓ 列名清洗: '{original_col}' -> '{cleaned_col}'")
        
        # 重命名列
        if new_columns:
            self.price_list.rename(columns=new_columns, inplace=True)
            print(f"  ✓ 共清洗了 {len(new_columns)} 个列名")
        else:
            print("  ✓ 所有列名无需清洗")
        
    def _clean_data(self):
        """
        数据标准化处理：
        1. 日期列：Effective Date 和 Expiry Date 转换为 datetime，失败时填充大范围日期
        2. 文本列：Carrier, POL Code, POD Code 转大写并去空格
        3. 确保 Month 列是字符串格式
        """
        if self.price_list is None or self.price_list.empty:
            return
        
        print("正在进行数据标准化处理...")
        
        # 1. 处理日期列：Effective Date 和 Expiry Date
        date_columns = []
        for col in self.price_list.columns:
            col_str = str(col).lower()
            if 'effective date' in col_str or 'effective_date' in col_str or '生效日期' in col_str:
                date_columns.append((col, 'Effective Date'))
            elif 'expiry date' in col_str or 'expiry_date' in col_str or '到期日期' in col_str or 'expire' in col_str:
                date_columns.append((col, 'Expiry Date'))
        
        # 设置默认日期范围（过去到未来）
        min_date = pd.Timestamp('1900-01-01')
        max_date = pd.Timestamp('2100-12-31')
        
        for col, col_type in date_columns:
            if col in self.price_list.columns:
                try:
                    # 尝试转换为 datetime
                    self.price_list[col] = pd.to_datetime(self.price_list[col], errors='coerce')
                    
                    # 对于空值，根据列类型填充默认值
                    if col_type == 'Effective Date':
                        # Effective Date 为空时，填充为过去的最早日期
                        self.price_list[col] = self.price_list[col].fillna(min_date)
                    elif col_type == 'Expiry Date':
                        # Expiry Date 为空时，填充为未来的最晚日期
                        self.price_list[col] = self.price_list[col].fillna(max_date)
                    
                    print(f"  ✓ {col_type} 列 ({col}) 已转换为 datetime 格式")
                except Exception as e:
                    print(f"  ⚠ 警告：{col_type} 列 ({col}) 转换失败: {e}")
                    # 转换失败时填充默认值
                    if col_type == 'Effective Date':
                        self.price_list[col] = min_date
                    else:
                        self.price_list[col] = max_date
        
        # 2. 处理文本列：Carrier, POL Code, POD Code
        # 2.1 处理 Carrier 列
        carrier_cols = [col for col in self.price_list.columns 
                       if 'carrier' in str(col).lower() or '船公司' in str(col)]
        if carrier_cols:
            for col in carrier_cols:
                if col in self.price_list.columns:
                    # 转换为字符串，转大写，去除首尾空格
                    self.price_list[col] = self.price_list[col].astype(str).str.strip().str.upper()
                    print(f"  ✓ Carrier 列 ({col}) 已标准化")
        
        # 2.2 处理 POL Code 列
        pol_cols = [col for col in self.price_list.columns 
                   if 'pol' in str(col).lower() and 'code' in str(col).lower()]
        if pol_cols:
            for col in pol_cols:
                if col in self.price_list.columns:
                    # 转换为字符串，转大写，去除首尾空格
                    self.price_list[col] = self.price_list[col].astype(str).str.strip().str.upper()
                    print(f"  ✓ POL Code 列 ({col}) 已标准化")
        
        # 2.3 处理 POD Code 列
        pod_cols = [col for col in self.price_list.columns 
                   if 'pod' in str(col).lower() and 'code' in str(col).lower()]
        
        # 如果没有找到，尝试使用第8列（H列，索引为7）
        if not pod_cols and len(self.price_list.columns) > 7:
            pod_cols = [self.price_list.columns[7]]  # H列
        
        if pod_cols:
            for col in pod_cols:
                if col in self.price_list.columns:
                    # 转换为字符串，转大写，去除首尾空格（保留中间的空格，用于逗号分隔的多个港口代码）
                    self.price_list[col] = self.price_list[col].astype(str).str.strip().str.upper()
                    print(f"  ✓ POD Code 列 ({col}) 已标准化")
        
        # 3. 确保 Month 列（第1列）是字符串格式
        first_col = self.price_list.columns[0]
        self.price_list[first_col] = self.price_list[first_col].astype(str)
        print(f"  ✓ Month 列 ({first_col}) 已转换为字符串格式")
        
        print("数据标准化处理完成！")
    
    def _normalize_container_type(self, raw_type):
        """
        标准化集装箱类型
        
        规则：
        1. 输入转大写
        2. 若包含 "40" 且包含 ("HQ" 或 "HIGH" 或 "CUBE") -> 返回 "40HQ"
        3. 若包含 "20" -> 返回 "20GP"
        4. 若包含 "40" 且不含 HQ 特征 -> 返回 "40GP"
        5. 其他情况返回 "Unknown"
        
        参数:
            raw_type: 原始集装箱类型字符串
        
        返回:
            str: 标准化后的集装箱类型 ("40HQ", "20GP", "40GP", "Unknown")
        """
        if not raw_type or pd.isna(raw_type):
            return "Unknown"
        
        # 转大写
        normalized = str(raw_type).upper()
        
        # 规则1: 若包含 "40" 且包含 ("HQ" 或 "HIGH" 或 "CUBE") -> 返回 "40HQ"
        if "40" in normalized:
            if "HQ" in normalized or "HIGH" in normalized or "CUBE" in normalized:
                return "40HQ"
            # 规则4: 若包含 "40" 且不含 HQ 特征 -> 返回 "40GP"
            return "40GP"
        
        # 规则3: 若包含 "20" -> 返回 "20GP"
        if "20" in normalized:
            return "20GP"
        
        # 规则5: 其他情况返回 "Unknown"
        return "Unknown"
    
    def _standardize_carrier_name(self, carrier_name):
        """
        标准化船公司名称，将可能出现的船公司全称映射为标准的简写代码（SCAC Code）
        
        规则（关键词包含匹配）：
        - 若包含 "YANG MING" 或 "YANGMING" -> 返回 "YML"
        - 若包含 "HYUNDAI" 或 "HMM" -> 返回 "HMM"
        - 若包含 "EVERGREEN" -> 返回 "EMC"
        - 若包含 "MAERSK" 或 "MSK" -> 返回 "MSK"
        - 若包含 "COSCO" -> 返回 "COSCO"
        - 若包含 "ONE" -> 返回 "ONE"
        - 若包含 "CMA" -> 返回 "CMA"
        - 若包含 "MSC" -> 返回 "MSC"
        - 若包含 "OOCL" -> 返回 "OOCL"
        - 如果都不匹配，返回原始的大写字符串（strip后）
        
        参数:
            carrier_name: 原始船公司名称字符串
        
        返回:
            str: 标准化后的船公司代码
        """
        if not carrier_name or pd.isna(carrier_name):
            return ""
        
        # 转大写并去除首尾空格
        normalized = str(carrier_name).strip().upper()
        
        # 关键词映射规则（按顺序检查）
        if "YANG MING" in normalized or "YANGMING" in normalized:
            return "YML"
        elif "HYUNDAI" in normalized or "HMM" in normalized:
            return "HMM"
        elif "EVERGREEN" in normalized:
            return "EMC"
        elif "MAERSK" in normalized or "MSK" in normalized:
            return "MSK"
        elif "COSCO" in normalized:
            return "COSCO"
        elif "ONE" in normalized:
            return "ONE"
        elif "CMA" in normalized:
            return "CMA"
        elif "MSC" in normalized:
            return "MSC"
        elif "OOCL" in normalized:
            return "OOCL"
        else:
            # 如果都不匹配，返回原始的大写字符串（strip后）
            return normalized
    
    def _convert_etd_to_month(self, etd_value):
        """
        将 ETD 转换为 YYYYMM 格式的字符串
        
        参数:
            etd_value: ETD 值，可能是字符串 "2025/11/15" 或 datetime 对象
        
        返回:
            str: YYYYMM 格式的字符串，例如 "202511"
        """
        if pd.isna(etd_value) or not etd_value:
            return None
        
        try:
            # 如果是 datetime 对象
            if isinstance(etd_value, (datetime, pd.Timestamp)):
                return etd_value.strftime("%Y%m")
            
            # 如果是字符串
            etd_str = str(etd_value).strip()
            
            # 尝试解析常见的日期格式
            # 格式1: "2025/11/15"
            if '/' in etd_str:
                parts = etd_str.split('/')
                if len(parts) >= 2:
                    year = parts[0].strip()
                    month = parts[1].strip().zfill(2)
                    return f"{year}{month}"
            
            # 格式2: "2025-11-15"
            if '-' in etd_str:
                parts = etd_str.split('-')
                if len(parts) >= 2:
                    year = parts[0].strip()
                    month = parts[1].strip().zfill(2)
                    return f"{year}{month}"
            
            # 尝试使用 pandas 解析
            parsed_date = pd.to_datetime(etd_str, errors='coerce')
            if not pd.isna(parsed_date):
                return parsed_date.strftime("%Y%m")
            
            return None
        except Exception as e:
            print(f"警告: ETD 转换失败: {etd_value}, 错误: {e}")
            return None
    
    def _find_price_column(self, container_type):
        """
        根据标准化后的柜型查找 Price List 中对应的价格列
        
        支持模糊匹配逻辑：
        1. 优先匹配完全相同的列（忽略大小写和空格）
        2. 如果找不到，尝试别名映射
        3. 如果还找不到，尝试模糊匹配（包含关系）
        
        参数:
            container_type: 标准化后的柜型 ("20GP", "40GP", "40HQ")
        
        返回:
            str: 价格列名，如果找不到则返回 None
        """
        if not container_type or container_type == "Unknown":
            return None
        
        # 定义别名映射
        alias_map = {
            "40HQ": ["40HC", "40 HC", "40High", "40 HIGH", "40HC", "40 HC"],
            "20GP": ["20 GP", "20FT", "20 FT", "20GP", "20 GP"]
        }
        
        # 获取目标柜型的大写形式（去除空格）
        target_upper = container_type.upper().replace(" ", "")
        
        # 第一步：优先匹配完全相同的列（忽略大小写和空格）
        for col in self.price_list.columns:
            col_normalized = str(col).upper().replace(" ", "").replace("　", "")
            if col_normalized == target_upper:
                return col
        
        # 第二步：如果找不到，尝试别名映射
        if container_type in alias_map:
            aliases = alias_map[container_type]
            for alias in aliases:
                alias_normalized = alias.upper().replace(" ", "").replace("　", "")
                for col in self.price_list.columns:
                    col_normalized = str(col).upper().replace(" ", "").replace("　", "")
                    if col_normalized == alias_normalized:
                        return col
                    # 也尝试包含匹配
                    if alias_normalized in col_normalized or col_normalized in alias_normalized:
                        return col
        
        # 第三步：如果还找不到，尝试模糊匹配（包含关系）
        for col in self.price_list.columns:
            col_normalized = str(col).upper().replace(" ", "").replace("　", "")
            if target_upper in col_normalized or col_normalized in target_upper:
                return col
        
        # 第四步：如果还找不到，尝试使用别名进行模糊匹配
        if container_type in alias_map:
            aliases = alias_map[container_type]
            for alias in aliases:
                alias_normalized = alias.upper().replace(" ", "").replace("　", "")
                for col in self.price_list.columns:
                    col_normalized = str(col).upper().replace(" ", "").replace("　", "")
                    if alias_normalized in col_normalized or col_normalized in alias_normalized:
                        return col
        
        return None
    
    def run_matching(self, info_excel_path):
        """
        执行价格匹配逻辑
        
        处理流程：
        1. 读取 info.xlsx，确保 ETD 列转换为 datetime（只保留日期部分）
        2. 遍历每一行，提取 ETD, Carrier, Loading Port Code, Destination Code, Container Type
        3. 在 Price List 中查找匹配行（Carrier、POL、POD、日期有效期）
        4. 根据标准化后的柜型获取价格并回填到 Standard Freight Price 列
        5. 覆盖保存 info.xlsx
        
        参数:
            info_excel_path: info.xlsx 文件路径
        
        返回:
            pandas.DataFrame: 更新后的 info.xlsx 数据
        """
        if self.price_list is None or self.price_list.empty:
            raise ValueError("请先调用 load_price_list() 加载价格列表")
        
        if not os.path.exists(info_excel_path):
            raise FileNotFoundError(f"文件不存在: {info_excel_path}")
        
        print(f"\n开始执行价格匹配...")
        print(f"读取 info.xlsx: {info_excel_path}")
        
        # 读取 info.xlsx
        df_info = pd.read_excel(info_excel_path, engine='openpyxl')
        print(f"  ✓ 成功读取 info.xlsx，共 {len(df_info)} 行数据")
        
        # 检查必要的列是否存在
        required_cols = ['ETD', 'Carrier', 'Loading Port Code', 'Destination Code', 'Container Type']
        missing_cols = [col for col in required_cols if col not in df_info.columns]
        if missing_cols:
            raise ValueError(f"info.xlsx 缺少必要的列: {', '.join(missing_cols)}")
        
        # 1. 将 ETD 列转换为 datetime（只保留日期部分，去除时间）
        print("正在转换 ETD 列为 datetime 格式...")
        df_info['ETD'] = pd.to_datetime(df_info['ETD'], errors='coerce')
        # 只保留日期部分，去除时间（使用 normalize() 保持为 datetime 类型）
        df_info['ETD'] = df_info['ETD'].dt.normalize()
        print(f"  ✓ ETD 列已转换为日期格式（去除时间部分）")
        
        # 初始化价格列
        df_info['Standard Freight Price'] = "N/A"
        
        # 打印所有列名用于调试
        print(f"  Price List 所有列名: {list(self.price_list.columns)}")
        
        # 查找 Price List 中的必要列
        carrier_col = None
        pol_col = None
        pod_col = None
        effective_date_col = None
        expiry_date_col = None
        
        for col in self.price_list.columns:
            col_str = str(col)
            col_lower = col_str.lower().strip()
            
            # Carrier 列匹配（更灵活）
            if not carrier_col:
                if ('carrier' in col_lower or 
                    '船公司' in col_str or 
                    'shipping line' in col_lower or
                    'line' in col_lower and 'carrier' not in col_lower or
                    col_lower in ['carrier', 'carrier name', 'carrier_name', '船公司名称']):
                    carrier_col = col
                    print(f"  ✓ 找到 Carrier 列: {col}")
            
            # POL Code 列匹配（更灵活）
            if not pol_col:
                if (('pol' in col_lower and 'code' in col_lower) or
                    ('pol' in col_lower and 'port' in col_lower) or
                    ('origin' in col_lower and 'code' in col_lower) or
                    ('origin' in col_lower and 'port' in col_lower and 'code' in col_lower) or
                    col_lower in ['pol code', 'pol_code', 'pol', 'origin port code', 'origin_port_code']):
                    pol_col = col
                    print(f"  ✓ 找到 POL Code 列: {col}")
            
            # POD Code 列匹配（更灵活）
            if not pod_col:
                if (('pod' in col_lower and 'code' in col_lower) or
                    ('pod' in col_lower and 'port' in col_lower) or
                    ('destination' in col_lower and 'code' in col_lower) or
                    ('discharge' in col_lower and 'code' in col_lower) or
                    col_lower in ['pod code', 'pod_code', 'pod', 'destination port code', 'destination_port_code']):
                    pod_col = col
                    print(f"  ✓ 找到 POD Code 列: {col}")
            
            # Effective Date 列匹配
            if not effective_date_col:
                if ('effective date' in col_lower or 
                    'effective_date' in col_lower or 
                    '生效日期' in col_str or
                    'effective' in col_lower and 'date' in col_lower):
                    effective_date_col = col
                    print(f"  ✓ 找到 Effective Date 列: {col}")
            
            # Expiry Date 列匹配
            if not expiry_date_col:
                if ('expiry date' in col_lower or 
                    'expiry_date' in col_lower or 
                    '到期日期' in col_str or 
                    'expire' in col_lower and 'date' in col_lower or
                    'valid until' in col_lower or
                    'valid_until' in col_lower):
                    expiry_date_col = col
                    print(f"  ✓ 找到 Expiry Date 列: {col}")
        
        # 如果还没找到，尝试按常见列位置查找（备用方案）
        # 通常 Carrier 在第2列（索引1），POL Code 在第3列（索引2），POD Code 在第8列（索引7）
        if not carrier_col and len(self.price_list.columns) > 1:
            # 尝试第2列（索引1）
            potential_carrier = self.price_list.columns[1]
            print(f"  ⚠ Carrier 列未找到，尝试使用第2列: {potential_carrier}")
            carrier_col = potential_carrier
        
        if not pol_col and len(self.price_list.columns) > 2:
            # 尝试第3列（索引2）
            potential_pol = self.price_list.columns[2]
            print(f"  ⚠ POL Code 列未找到，尝试使用第3列: {potential_pol}")
            pol_col = potential_pol
        
        # 如果没有找到 POD Code 列，尝试使用第8列（H列，索引为7）
        if not pod_col and len(self.price_list.columns) > 7:
            potential_pod = self.price_list.columns[7]
            print(f"  ⚠ POD Code 列未找到，尝试使用第8列: {potential_pod}")
            pod_col = potential_pod
        
        # 如果日期列未找到，尝试查找包含 "date" 的列
        if not effective_date_col:
            for col in self.price_list.columns:
                if 'date' in str(col).lower() and 'expir' not in str(col).lower():
                    effective_date_col = col
                    print(f"  ⚠ Effective Date 列未找到，尝试使用: {col}")
                    break
        
        if not expiry_date_col:
            for col in self.price_list.columns:
                if 'date' in str(col).lower() and col != effective_date_col:
                    expiry_date_col = col
                    print(f"  ⚠ Expiry Date 列未找到，尝试使用: {col}")
                    break
        
        if not carrier_col or not pol_col or not pod_col:
            error_msg = f"Price List 中缺少必要的列:\n"
            error_msg += f"  - Carrier: {carrier_col or '未找到'}\n"
            error_msg += f"  - POL Code: {pol_col or '未找到'}\n"
            error_msg += f"  - POD Code: {pod_col or '未找到'}\n"
            error_msg += f"\n请检查 Price List 文件，确保包含以下列：\n"
            error_msg += f"  - Carrier (或 船公司)\n"
            error_msg += f"  - POL Code (或 Origin Port Code)\n"
            error_msg += f"  - POD Code (或 Destination Port Code)"
            raise ValueError(error_msg)
        
        if not effective_date_col or not expiry_date_col:
            error_msg = f"Price List 中缺少日期列:\n"
            error_msg += f"  - Effective Date: {effective_date_col or '未找到'}\n"
            error_msg += f"  - Expiry Date: {expiry_date_col or '未找到'}\n"
            error_msg += f"\n请检查 Price List 文件，确保包含日期列。"
            raise ValueError(error_msg)
        
        print(f"  ✓ 最终匹配列: Carrier={carrier_col}, POL Code={pol_col}, POD Code={pod_col}")
        print(f"  ✓ 最终日期列: Effective Date={effective_date_col}, Expiry Date={expiry_date_col}")
        
        # 确保 Price List 的日期列是 datetime 格式
        if self.price_list[effective_date_col].dtype != 'datetime64[ns]':
            self.price_list[effective_date_col] = pd.to_datetime(self.price_list[effective_date_col], errors='coerce')
        if self.price_list[expiry_date_col].dtype != 'datetime64[ns]':
            self.price_list[expiry_date_col] = pd.to_datetime(self.price_list[expiry_date_col], errors='coerce')
        
        # 将日期标准化为只有日期部分（去除时间），但仍保持为 datetime 类型以便比较
        self.price_list[effective_date_col] = pd.to_datetime(self.price_list[effective_date_col]).dt.normalize()
        self.price_list[expiry_date_col] = pd.to_datetime(self.price_list[expiry_date_col]).dt.normalize()
        
        # 统计信息
        matched_count = 0
        unmatched_count = 0
        
        # 遍历 info.xlsx 的每一行
        for idx, row in df_info.iterrows():
            # 获取当前行的数据
            etd = row['ETD']
            carrier = row['Carrier']
            loading_port_code = row['Loading Port Code']
            destination_code = row['Destination Code']
            container_type = row['Container Type']
            
            # 跳过空值行
            if pd.isna(etd) or pd.isna(carrier) or pd.isna(loading_port_code) or pd.isna(destination_code):
                df_info.at[idx, 'Standard Freight Price'] = "N/A"
                unmatched_count += 1
                continue
            
            # 标准化 Carrier 和港口代码（转大写并去除首尾空格）
            carrier_normalized = self._standardize_carrier_name(carrier)
            loading_port_code_normalized = str(loading_port_code).strip().upper()
            destination_code_normalized = str(destination_code).strip().upper()
            
            # 标准化集装箱类型
            normalized_container = self._normalize_container_type(container_type)
            
            if normalized_container == "Unknown":
                df_info.at[idx, 'Standard Freight Price'] = "N/A"
                unmatched_count += 1
                continue
            
            # 在 Price List 中查找匹配的行
            # 匹配条件：
            # 1. Carrier 匹配
            # 2. POL Code 匹配
            # 3. POD Code 包含 Destination Code（使用包含逻辑）
            # 4. 日期有效期匹配：Effective Date <= ETD <= Expiry Date
            
            # 创建匹配条件
            # 对费率表的 Carrier 也应用同样的标准化函数，确保双方一致
            carrier_match = (self.price_list[carrier_col].apply(self._standardize_carrier_name) == carrier_normalized)
            pol_match = (self.price_list[pol_col].astype(str).str.strip().str.upper() == loading_port_code_normalized)
            pod_match = (self.price_list[pod_col].astype(str).str.contains(destination_code_normalized, case=False, na=False))
            
            # 日期有效期匹配：ETD 已经是 datetime 类型（在前面已转换）
            # 确保它是 datetime 类型（去除时间部分）
            try:
                if pd.isna(etd):
                    df_info.at[idx, 'Standard Freight Price'] = "N/A"
                    unmatched_count += 1
                    continue
                
                # ETD 应该已经是 pd.Timestamp 类型，确保是 normalize 的（只保留日期部分）
                if isinstance(etd, pd.Timestamp):
                    etd_datetime = etd.normalize()
                else:
                    # 如果还不是 datetime，转换它
                    etd_datetime = pd.to_datetime(etd, errors='coerce').normalize()
                    if pd.isna(etd_datetime):
                        df_info.at[idx, 'Standard Freight Price'] = "N/A"
                        unmatched_count += 1
                        continue
            except Exception as e:
                df_info.at[idx, 'Standard Freight Price'] = "N/A"
                unmatched_count += 1
                continue
            
            # 日期匹配：Effective Date <= ETD <= Expiry Date
            date_match = (
                (self.price_list[effective_date_col] <= etd_datetime) & 
                (self.price_list[expiry_date_col] >= etd_datetime)
            )
            
            # 组合所有匹配条件
            matches = self.price_list[carrier_match & pol_match & pod_match & date_match]
            
            if len(matches) > 0:
                # 找到匹配，取第一条（如果有重叠时间段，取第一条）
                match_row = matches.iloc[0]
                
                # 根据标准化后的柜型查找价格列
                price_col = self._find_price_column(normalized_container)
                
                if price_col and price_col in match_row.index:
                    price_value = match_row[price_col]
                    if not pd.isna(price_value):
                        df_info.at[idx, 'Standard Freight Price'] = price_value
                        matched_count += 1
                        print(f"  ✓ 行 {idx + 2}: 匹配成功，价格={price_value} ({normalized_container})")
                    else:
                        df_info.at[idx, 'Standard Freight Price'] = "N/A"
                        unmatched_count += 1
                        print(f"  ✗ 行 {idx + 2}: 匹配成功但价格为空 ({normalized_container})")
                else:
                    df_info.at[idx, 'Standard Freight Price'] = "N/A"
                    unmatched_count += 1
                    print(f"  ✗ 行 {idx + 2}: 匹配成功但找不到价格列 ({normalized_container})")
            else:
                df_info.at[idx, 'Standard Freight Price'] = "N/A"
                unmatched_count += 1
                etd_display = etd_datetime.strftime("%Y-%m-%d") if not pd.isna(etd_datetime) else str(etd)
                print(f"  ✗ 行 {idx + 2}: 未找到匹配 (ETD={etd_display}, Carrier={carrier_normalized}, POL={loading_port_code_normalized}, POD={destination_code_normalized})")
        
        # 保存更新后的文件
        print(f"\n匹配完成: 成功 {matched_count} 条，失败 {unmatched_count} 条")
        print(f"保存更新后的 info.xlsx...")
        
        # 覆盖保存到原文件
        df_info.to_excel(info_excel_path, index=False, engine='openpyxl')
        print(f"  ✓ 已保存到: {info_excel_path}")
        
        return df_info
    
    def get_price_list(self):
        """
        获取当前加载的价格列表
        
        返回:
            pandas.DataFrame: 价格列表数据，如果未加载则返回 None
        """
        return self.price_list


if __name__ == "__main__":
    # 测试代码
    matcher = FreightMatcher()
    # 示例用法：
    # matcher.load_price_list("path/to/price_list.xlsx")
    # print(matcher.get_price_list().head())


