"""
主程序入口文件
负责协调邮件下载、发票提取、文件归档和Excel报表生成
"""

import os
import shutil
import pandas as pd
from datetime import datetime
import time
import invoice_extractor
import EmailHandler
import config_loader


# ================= 配置区域 =================
# 从配置文件加载邮箱信息
try:
    MAIL_USER, MAIL_PASS = config_loader.get_email_config()
except (FileNotFoundError, ValueError) as e:
    print(f"❌ 配置加载失败: {e}")
    print("请确保已创建 config.ini 文件并填写正确的配置信息。")
    exit(1)

# 根目录（当前工作目录）
BASE_DIR = os.getcwd()
# ===========================================


def sanitize_filename(filename):
    """
    清理文件名中的非法字符（Windows文件名不能包含：< > : " / \ | ? *）
    
    参数:
        filename (str): 原始文件名
        
    返回:
        str: 清理后的文件名
    """
    # Windows文件名非法字符
    illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    return filename


def main():
    """
    主函数：执行完整的邮件处理流程
    """
    print("=" * 60)
    print("开始执行发票自动处理程序")
    print("=" * 60)
    
    # 用于统计
    all_excel_data = []  # 存储所有Excel行数据
    processed_email_count = 0  # 处理的邮件总数
    success_extract_count = 0  # 成功提取的发票数量
    start_time = datetime.now()  # 记录开始时间
    
    try:
        # ==================== 步骤 1：初始化目录 ====================
        print("\n【步骤 1】初始化目录结构...")
        
        # 获取当前日期字符串（格式：YYYYMMDD）
        current_date = datetime.now().strftime("%Y%m%d")
        print(f"当前日期: {current_date}")
        
        # 构建基础路径：{BASE_DIR}/Download/{日期}/
        base_path = os.path.join(BASE_DIR, "Download", current_date)
        print(f"基础路径: {base_path}")
        
        # 创建基础目录（如果不存在）
        if not os.path.exists(base_path):
            os.makedirs(base_path)
            print(f"✓ 已创建基础目录: {base_path}")
        
        # 在该路径下创建三个文件夹
        temp_dir = os.path.join(base_path, "Temp")  # 临时文件夹（存放EmailHandler下载的原始文件）
        invoice_dir = os.path.join(base_path, "Invoice附件")  # 发票附件文件夹
        bl_dir = os.path.join(base_path, "BL附件")  # 提单附件文件夹
        
        # 创建各个子目录
        for dir_path in [temp_dir, invoice_dir, bl_dir]:
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
                print(f"✓ 已创建目录: {dir_path}")
            else:
                print(f"✓ 目录已存在: {dir_path}")
        
        print("目录初始化完成！\n")
        
        # ==================== 步骤 2：执行下载 ====================
        print("【步骤 2】从邮箱下载并处理附件...")
        email_list = EmailHandler.download_and_process_attachments(MAIL_USER, MAIL_PASS, temp_dir)
        
        if not email_list:
            print("⚠ 警告：没有获取到任何邮件，程序结束。")
            return
        
        print(f"✓ 成功获取 {len(email_list)} 封有效邮件\n")
        
        # ==================== 步骤 3：处理每一封邮件 ====================
        print("【步骤 3】处理邮件和附件...")
        
        for email_idx, email_info in enumerate(email_list, 1):
            processed_email_count += 1
            print(f"\n--- 处理第 {email_idx}/{len(email_list)} 封邮件 ---")
            print(f"邮件标题: {email_info['subject']}")
            print(f"Booking No: {email_info.get('booking_no', '未提取到')}")
            
            # 在当前邮件的附件中，寻找 type 为 "INVOICE" 或 "UNKNOWN" 的文件（容错机制：UNKNOWN 也尝试作为发票处理）
            # 以及 type 为 "BL" 或 "UNKNOWN" 的文件（容错机制：UNKNOWN 也尝试作为提单处理）
            # 注意：如果 UNKNOWN 文件被当作发票处理并移动后，在处理 BL 时文件已不在原位置，移动操作会失败但不影响程序
            invoice_files = [att for att in email_info['attachments'] if att['type'] == 'INVOICE' or att['type'] == 'UNKNOWN']
            bl_files = [att for att in email_info['attachments'] if att['type'] == 'BL' or att['type'] == 'UNKNOWN']
            
            # 跳过逻辑：如果这封邮件里没有 Invoice 文件，打印日志并跳过
            if not invoice_files:
                print("  ⚠ 跳过：该邮件中没有 Invoice 文件")
                continue
            
            print(f"  发现 {len(invoice_files)} 个 Invoice 文件，{len(bl_files)} 个 BL 文件")
            
            # 处理每个 Invoice 文件（通常一封邮件只有一个Invoice）
            for invoice_att in invoice_files:
                invoice_path = invoice_att['path']
                invoice_filename = os.path.basename(invoice_path)
                print(f"\n  处理 Invoice: {invoice_filename}")
                
                # 提取数据
                print("  正在调用 AI 提取发票数据...")
                extracted_data = invoice_extractor.extract_invoice_data(invoice_path)
                
                # 如果返回数据为空，跳过
                if not extracted_data:
                    print("  ⚠ 跳过：AI 提取失败或返回空数据")
                    continue
                
                # 获取提取结果的第一条数据（通常是列表的第一个元素）
                first_data = extracted_data[0]
                
                # 从中提取发票号和 HBL 字段（用于文件重命名）
                # 优先使用 InvoiceNo，如果为空则使用 OriginalFileNo 作为兜底
                invoice_no = first_data.get('InvoiceNo', '').strip() if first_data.get('InvoiceNo') else ''
                if not invoice_no:
                    invoice_no = first_data.get('OriginalFileNo', '').strip() if first_data.get('OriginalFileNo') else ''
                hbl = first_data.get('HBL', '').strip() if first_data.get('HBL') else ''
                
                print(f"  提取到 InvoiceNo: {invoice_no}")
                print(f"  提取到 HBL: {hbl}")
                
                # 数据汇总：生成 Excel 行数据
                booking_no = email_info.get('booking_no', '')
                excel_row = invoice_extractor.prepare_excel_row(first_data, invoice_filename, booking_no)
                all_excel_data.append(excel_row)
                success_extract_count += 1
                print(f"  ✓ 已添加到 Excel 数据列表")
                
                # ========== 重命名与归档（核心部分）==========
                print("\n  【文件移动和重命名】")
                
                # 处理 Invoice：移动到 Invoice附件 文件夹并重命名
                if invoice_no:
                    # 构建目标文件名：invoice {invoice_no}.pdf
                    target_invoice_name = f"invoice {invoice_no}.pdf"
                    # 清理文件名中的非法字符
                    target_invoice_name = sanitize_filename(target_invoice_name)
                    target_invoice_path = os.path.join(invoice_dir, target_invoice_name)
                    
                    # 如果目标文件已存在，添加序号避免覆盖
                    counter = 1
                    original_target = target_invoice_path
                    while os.path.exists(target_invoice_path):
                        name, ext = os.path.splitext(original_target)
                        target_invoice_path = f"{name}_{counter}{ext}"
                        counter += 1
                    
                    # 使用 shutil.move 将文件从 Temp 移到 Invoice附件 文件夹并重命名
                    try:
                        shutil.move(invoice_path, target_invoice_path)
                        print(f"  ✓ Invoice 已移动并重命名: {os.path.basename(target_invoice_path)}")
                    except Exception as e:
                        print(f"  ✗ Invoice 移动失败: {e}")
                else:
                    # 如果 invoice_no 为空，保留原文件名但移动到目标文件夹
                    target_invoice_name = invoice_filename
                    target_invoice_path = os.path.join(invoice_dir, target_invoice_name)
                    
                    # 处理重名情况
                    counter = 1
                    original_target = target_invoice_path
                    while os.path.exists(target_invoice_path):
                        name, ext = os.path.splitext(original_target)
                        target_invoice_path = f"{name}_{counter}{ext}"
                        counter += 1
                    
                    try:
                        shutil.move(invoice_path, target_invoice_path)
                        print(f"  ✓ Invoice 已移动（未重命名，invoice_no为空）: {os.path.basename(target_invoice_path)}")
                    except Exception as e:
                        print(f"  ✗ Invoice 移动失败: {e}")
                
                # 处理 BL（如果存在）：移动到 BL附件 文件夹并重命名
                if bl_files:
                    for bl_att in bl_files:
                        bl_path = bl_att['path']
                        bl_filename = os.path.basename(bl_path)
                        
                        # 检查文件是否还存在（可能已经被当作发票处理并移动了）
                        if not os.path.exists(bl_path):
                            print(f"  ⚠ BL 文件已不存在（可能已被处理）: {bl_filename}")
                            continue
                        
                        if hbl:
                            # 构建目标文件名：BL {HBL}.pdf
                            target_bl_name = f"BL {hbl}.pdf"
                            # 清理文件名中的非法字符
                            target_bl_name = sanitize_filename(target_bl_name)
                            target_bl_path = os.path.join(bl_dir, target_bl_name)
                            
                            # 如果目标文件已存在，添加序号避免覆盖
                            counter = 1
                            original_target = target_bl_path
                            while os.path.exists(target_bl_path):
                                name, ext = os.path.splitext(original_target)
                                target_bl_path = f"{name}_{counter}{ext}"
                                counter += 1
                            
                            # 使用 shutil.move 将文件从 Temp 移到 BL附件 文件夹并重命名
                            try:
                                shutil.move(bl_path, target_bl_path)
                                print(f"  ✓ BL 已移动并重命名: {os.path.basename(target_bl_path)}")
                            except Exception as e:
                                print(f"  ✗ BL 移动失败: {e}")
                        else:
                            # 如果 HBL 为空，保留原名或标记未知
                            target_bl_name = f"BL_未知_{bl_filename}"
                            target_bl_path = os.path.join(bl_dir, target_bl_name)
                            
                            # 处理重名情况
                            counter = 1
                            original_target = target_bl_path
                            while os.path.exists(target_bl_path):
                                name, ext = os.path.splitext(original_target)
                                target_bl_path = f"{name}_{counter}{ext}"
                                counter += 1
                            
                            try:
                                shutil.move(bl_path, target_bl_path)
                                print(f"  ✓ BL 已移动（HBL为空，标记为未知）: {os.path.basename(target_bl_path)}")
                            except Exception as e:
                                print(f"  ✗ BL 移动失败: {e}")
        
        print(f"\n邮件处理完成！共处理 {processed_email_count} 封邮件，成功提取 {success_extract_count} 条发票数据\n")
        
        # ==================== 步骤 4：生成 Excel 报表 ====================
        print("【步骤 4】生成 Excel 报表...")
        
        # 生成 info.xlsx
        if all_excel_data:
            # 定义表头（严格按照要求的顺序）
            headers = [
                "NO", "File Name", "FILENO", "File No", "DATE", "Carrier", "Vessel/Voyage",
                "Loading Port", "Loading Port Code", "Destination", "Destination Code",
                "ETD", "ETA", "Receipt", "OBL", "HBL", "MBL",
                "Item", "Quantity", "Unit Price", "Container Type", "Amount", "Booking No"
            ]
            
            # 使用 pandas 创建 DataFrame
            df = pd.DataFrame(all_excel_data, columns=headers)
            
            # 保存到 Excel 文件
            info_excel_path = os.path.join(base_path, "info.xlsx")
            df.to_excel(info_excel_path, index=False, engine='openpyxl')
            print(f"✓ 已生成 info.xlsx: {info_excel_path}")
        else:
            print("⚠ 警告：没有数据可写入 info.xlsx")
        
        # 生成 当日运行清单.xlsx
        end_time = datetime.now()
        run_duration = (end_time - start_time).total_seconds()
        status = "成功" if success_extract_count > 0 else "无数据"
        
        summary_data = {
            "运行时间": [start_time.strftime("%Y-%m-%d %H:%M:%S")],
            "处理邮件总数": [processed_email_count],
            "成功提取数": [success_extract_count],
            "状态": [status]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_excel_path = os.path.join(base_path, "当日运行清单.xlsx")
        summary_df.to_excel(summary_excel_path, index=False, engine='openpyxl')
        print(f"✓ 已生成 当日运行清单.xlsx: {summary_excel_path}")
        
        print("Excel 报表生成完成！\n")
        
        # ==================== 步骤 5：清理环境 ====================
        print("【步骤 5】清理临时文件...")
        
        # 尝试删除空的 Temp 文件夹
        try:
            if os.path.exists(temp_dir):
                # 检查文件夹是否为空
                if not os.listdir(temp_dir):
                    os.rmdir(temp_dir)
                    print(f"✓ 已删除空的 Temp 文件夹: {temp_dir}")
                else:
                    print(f"⚠ Temp 文件夹不为空，保留文件夹（可能还有未处理的文件）")
        except Exception as e:
            print(f"⚠ 删除 Temp 文件夹时出错: {e}")
        
        print("\n" + "=" * 60)
        print("程序执行完成！")
        print("=" * 60)
        print(f"处理邮件数: {processed_email_count}")
        print(f"成功提取数: {success_extract_count}")
        print(f"输出目录: {base_path}")
        
    except Exception as e:
        print(f"\n❌ 程序执行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

