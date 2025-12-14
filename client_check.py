"""
client_check.py
用于处理 Excel 数据的后期匹配和校验模块
"""

import os
import pandas as pd
from openpyxl import load_workbook


def run_check(info_excel_path, booking_list_path):
    """
    主函数：执行客户匹配和校验
    
    Args:
        info_excel_path: info.xlsx 的路径
        booking_list_path: Booking List Excel 文件的路径
    """
    print("=" * 60)
    print("开始执行客户匹配和校验...")
    print("=" * 60)
    
    # ==================== 步骤 1: Pre-check ====================
    print(f"\n【步骤 1】预检查...")
    if not booking_list_path or not os.path.exists(booking_list_path):
        print(f"⚠ 警告: Booking List 文件路径为空或文件不存在: {booking_list_path}")
        print("跳过客户匹配流程")
        return
    
    if not os.path.exists(info_excel_path):
        print(f"✗ 错误: info.xlsx 文件不存在: {info_excel_path}")
        return
    
    print(f"✓ Booking List 文件存在: {booking_list_path}")
    print(f"✓ info.xlsx 文件存在: {info_excel_path}")
    
    # ==================== 步骤 2: Load Source Data (Booking List) ====================
    print(f"\n【步骤 2】加载 Booking List 数据（全行扫描模式）...")
    
    try:
        # 读取所有 Sheet
        excel_file = pd.ExcelFile(booking_list_path)
        sheet_names = excel_file.sheet_names
        print(f"  发现 {len(sheet_names)} 个 Sheet: {', '.join(sheet_names)}")
        
        # 汇总所有数据，记录来源信息
        source_data = []
        for sheet_name in sheet_names:
            print(f"  正在读取 Sheet: {sheet_name}")
            df_sheet = pd.read_excel(booking_list_path, sheet_name=sheet_name)
            
            # 步骤 2.1: 确定 Client 列索引
            # 扫描表头，寻找包含关键词的列
            client_keywords = ['Client', 'Customer', 'Cnee', 'Consignee']
            client_col_index = None
            client_col_name = None
            
            for col_idx, col_name in enumerate(df_sheet.columns):
                col_name_str = str(col_name).strip().upper()
                for keyword in client_keywords:
                    if keyword.upper() in col_name_str:
                        client_col_index = col_idx
                        client_col_name = col_name
                        print(f"    ✓ 找到 Client 列: 第 {col_idx + 1} 列 '{col_name}'")
                        break
                if client_col_index is not None:
                    break
            
            # 如果找不到 Client 列，跳过该 Sheet
            if client_col_index is None:
                print(f"    ⚠ 警告: Sheet '{sheet_name}' 中未找到 Client 列（关键词: {', '.join(client_keywords)}），跳过该 Sheet")
                continue
            
            # 步骤 2.2: 全行扫描 - 记录每一行的所有数据
            for idx, row in df_sheet.iterrows():
                # Excel 行号 = pandas 索引 + 2 (因为 Excel 第一行是表头，从第2行开始是数据)
                excel_row_index = idx + 2
                
                # 将当前行所有单元格的值转换为字符串列表（去除空格，转大写）
                row_values = []
                for cell_value in row:
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip().upper()
                        if cell_str:  # 只添加非空值
                            row_values.append(cell_str)
                
                # 提取 Client 值（从确定的列索引获取）
                client_value = ''
                if client_col_index < len(row):
                    client_raw = row.iloc[client_col_index]
                    if pd.notna(client_raw):
                        client_value = str(client_raw).strip()
                
                source_data.append({
                    'Sheet Name': sheet_name,
                    'Row Index': excel_row_index,
                    'Row Values': row_values,  # 全行数据（大写，去空格）
                    'Client': client_value,
                    'Client Col Index': client_col_index,
                    'Client Col Name': client_col_name
                })
            
            print(f"    ✓ Sheet '{sheet_name}' 加载完成，共 {len([s for s in source_data if s['Sheet Name'] == sheet_name])} 行数据")
        
        print(f"✓ 共加载 {len(source_data)} 条源数据记录")
        
    except Exception as e:
        print(f"✗ 读取 Booking List 失败: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # ==================== 步骤 3: Load Target Data ====================
    print(f"\n【步骤 3】加载 info.xlsx 数据...")
    
    try:
        df_target = pd.read_excel(info_excel_path)
        print(f"✓ 成功加载 info.xlsx，共 {len(df_target)} 行数据")
        
        # 确保必要的列存在
        required_columns = ['OBL', 'HBL', 'Booking No']
        missing_columns = [col for col in required_columns if col not in df_target.columns]
        if missing_columns:
            print(f"✗ 错误: info.xlsx 缺少必要的列: {', '.join(missing_columns)}")
            return
        
    except Exception as e:
        print(f"✗ 读取 info.xlsx 失败: {e}")
        return
    
    # ==================== 步骤 4: Iterate & Match ====================
    print(f"\n【步骤 4】开始逐行匹配...")
    
    # 初始化结果列
    client_names = []
    booking_list_positions = []
    notes = []
    
    match_statistics = {
        'total_rows': len(df_target),
        'no_match': 0,
        'single_match': 0,
        'multiple_match': 0
    }
    
    for idx, row in df_target.iterrows():
        # 获取匹配用的值（转大写，去空格，用于匹配）
        obl_raw = str(row.get('OBL', '')).strip() if pd.notna(row.get('OBL')) else ''
        hbl_raw = str(row.get('HBL', '')).strip() if pd.notna(row.get('HBL')) else ''
        booking_no_raw = str(row.get('Booking No', '')).strip() if pd.notna(row.get('Booking No')) else ''
        
        # 转换为大写用于匹配
        obl = obl_raw.upper() if obl_raw else ''
        hbl = hbl_raw.upper() if hbl_raw else ''
        booking_no = booking_no_raw.upper() if booking_no_raw else ''
        
        # 空值熔断：如果全部为空，判定为未找到
        if not obl and not hbl and not booking_no:
            print(f"  行 {idx + 2}: OBL, HBL, Booking No 全部为空，判定为未找到")
            client_names.append("no client mapping")
            booking_list_positions.append("")
            notes.append("")
            match_statistics['no_match'] += 1
            continue
        
        # 在源数据中查找匹配（全行扫描）
        matches = []
        for source_item in source_data:
            row_values = source_item['Row Values']  # 已经是大写、去空格的列表
            
            # 匹配规则：检查目标值是否存在于该行的任何单元格中
            # 只要任一条件满足即可（使用 or 逻辑）
            matched = (
                (booking_no and booking_no in row_values) or
                (obl and obl in row_values) or
                (hbl and hbl in row_values)
            )
            
            if matched:
                matches.append(source_item)
        
        # ==================== 步骤 5: Determine Result ====================
        if len(matches) == 0:
            # Case A: 0 Matches
            print(f"  行 {idx + 2}: 未找到匹配 (OBL={obl_raw or '(空)'}, HBL={hbl_raw or '(空)'}, Booking No={booking_no_raw or '(空)'})")
            client_names.append("no client mapping")
            booking_list_positions.append("")
            notes.append("")
            match_statistics['no_match'] += 1
            
        elif len(matches) == 1:
            # Case B: 1 Match
            match = matches[0]
            client_name = match['Client']
            position = f"{match['Sheet Name']}-row{match['Row Index']}"
            print(f"  行 {idx + 2}: 找到 1 个匹配 -> Client: {client_name}, Position: {position}")
            client_names.append(client_name)
            booking_list_positions.append(position)
            notes.append("")
            match_statistics['single_match'] += 1
            
        else:
            # Case C: >1 Matches
            first_match = matches[0]
            client_name = first_match['Client']
            position = f"{first_match['Sheet Name']}-row{first_match['Row Index']}"
            note = f"Warning: Multiple matches found ({len(matches)})"
            print(f"  行 {idx + 2}: 找到 {len(matches)} 个匹配 -> 使用第一个: Client: {client_name}, Position: {position}")
            client_names.append(client_name)
            booking_list_positions.append(position)
            notes.append(note)
            match_statistics['multiple_match'] += 1
    
    # 打印匹配统计
    print(f"\n【匹配统计】")
    print(f"  总行数: {match_statistics['total_rows']}")
    print(f"  未匹配: {match_statistics['no_match']}")
    print(f"  单匹配: {match_statistics['single_match']}")
    print(f"  多匹配: {match_statistics['multiple_match']}")
    
    # ==================== 步骤 6: Save ====================
    print(f"\n【步骤 6】保存结果到 info.xlsx...")
    
    try:
        # 将结果列添加到 DataFrame
        df_target['Client Name'] = client_names
        df_target['Booking List Position'] = booking_list_positions
        df_target['Note'] = notes
        
        # 使用 openpyxl 引擎保存，确保格式正确
        df_target.to_excel(info_excel_path, index=False, engine='openpyxl')
        print(f"✓ 成功保存结果到: {info_excel_path}")
        print(f"  - Client Name 列 (V列) 已更新")
        print(f"  - Booking List Position 列 (W列) 已更新")
        print(f"  - Note 列 (X列) 已更新")
        
    except Exception as e:
        print(f"✗ 保存失败: {e}")
        return
    
    print("\n" + "=" * 60)
    print("客户匹配和校验完成！")
    print("=" * 60)


if __name__ == "__main__":
    # 测试代码
    print("client_check.py 模块已加载")
    print("使用方法: run_check(info_excel_path, booking_list_path)")

