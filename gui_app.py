"""
自动化单据处理工具 - 图形界面入口
包含敏感信息的硬编码配置，不显示在界面上
"""

import os
import sys
import shutil
import pandas as pd
from datetime import datetime
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import traceback

# 导入项目模块
import invoice_extractor
import EmailHandler

# ================= 硬编码配置区域（不显示在界面）=================
MAIL_USER = "freightforceone@qq.com"
MAIL_PASS = "wehhokderyvncbbf"
API_KEY = "sk-cb441e489cd84dc8906e37733ed9181e"
# ================================================================


class TextRedirector:
    """重定向 print 输出到 ScrolledText 组件"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
    
    def write(self, text):
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)  # 自动滚动到底部
        self.text_widget.update_idletasks()  # 强制刷新界面
    
    def flush(self):
        pass


def sanitize_filename(filename):
    """
    清理文件名中的非法字符（Windows文件名不能包含：< > : " / \ | ? *）
    
    参数:
        filename (str): 原始文件名
        
    返回:
        str: 清理后的文件名
    """
    illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    return filename


def run_main_process(base_dir, log_output):
    """
    主处理函数：执行完整的邮件处理流程（在独立线程中运行）
    
    参数:
        base_dir (str): 基础目录路径
        log_output: TextRedirector 对象，用于输出日志
    """
    # 强制覆盖 invoice_extractor 模块的 API_KEY
    invoice_extractor.API_KEY = API_KEY
    
    # 重定向 print 输出到 GUI 文本框
    old_stdout = sys.stdout
    sys.stdout = log_output
    
    try:
        print("=" * 60)
        print("开始执行发票自动处理程序")
        print("=" * 60)
        
        # 用于统计
        all_excel_data = []
        processed_email_count = 0
        success_extract_count = 0
        start_time = datetime.now()
        
        # ==================== 步骤 1：初始化目录 ====================
        print("\n【步骤 1】初始化目录结构...")
        
        current_date = datetime.now().strftime("%Y%m%d")
        print(f"当前日期: {current_date}")
        
        base_path = os.path.join(base_dir, "Download", current_date)
        print(f"基础路径: {base_path}")
        
        if not os.path.exists(base_path):
            os.makedirs(base_path)
            print(f"✓ 已创建基础目录: {base_path}")
        
        temp_dir = os.path.join(base_path, "Temp")
        invoice_dir = os.path.join(base_path, "Invoice附件")
        bl_dir = os.path.join(base_path, "BL附件")
        
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
            
            invoice_files = [att for att in email_info['attachments'] if att['type'] == 'INVOICE' or att['type'] == 'UNKNOWN']
            bl_files = [att for att in email_info['attachments'] if att['type'] == 'BL' or att['type'] == 'UNKNOWN']
            
            if not invoice_files:
                print("  ⚠ 跳过：该邮件中没有 Invoice 文件")
                continue
            
            print(f"  发现 {len(invoice_files)} 个 Invoice 文件，{len(bl_files)} 个 BL 文件")
            
            for invoice_att in invoice_files:
                invoice_path = invoice_att['path']
                invoice_filename = os.path.basename(invoice_path)
                print(f"\n  处理 Invoice: {invoice_filename}")
                
                print("  正在调用 AI 提取发票数据...")
                extracted_data = invoice_extractor.extract_invoice_data(invoice_path)
                
                if not extracted_data:
                    print("  ⚠ 跳过：AI 提取失败或返回空数据")
                    continue
                
                first_data = extracted_data[0]
                
                invoice_no = first_data.get('InvoiceNo', '').strip() if first_data.get('InvoiceNo') else ''
                if not invoice_no:
                    invoice_no = first_data.get('OriginalFileNo', '').strip() if first_data.get('OriginalFileNo') else ''
                hbl = first_data.get('HBL', '').strip() if first_data.get('HBL') else ''
                
                print(f"  提取到 InvoiceNo: {invoice_no}")
                print(f"  提取到 HBL: {hbl}")
                
                booking_no = email_info.get('booking_no', '')
                excel_row = invoice_extractor.prepare_excel_row(first_data, invoice_filename, booking_no)
                all_excel_data.append(excel_row)
                success_extract_count += 1
                print(f"  ✓ 已添加到 Excel 数据列表")
                
                # ========== 重命名与归档 ==========
                print("\n  【文件移动和重命名】")
                
                # 处理 Invoice
                if invoice_no:
                    target_invoice_name = f"invoice {invoice_no}.pdf"
                    target_invoice_name = sanitize_filename(target_invoice_name)
                    target_invoice_path = os.path.join(invoice_dir, target_invoice_name)
                    
                    counter = 1
                    original_target = target_invoice_path
                    while os.path.exists(target_invoice_path):
                        name, ext = os.path.splitext(original_target)
                        target_invoice_path = f"{name}_{counter}{ext}"
                        counter += 1
                    
                    try:
                        shutil.move(invoice_path, target_invoice_path)
                        print(f"  ✓ Invoice 已移动并重命名: {os.path.basename(target_invoice_path)}")
                    except Exception as e:
                        print(f"  ✗ Invoice 移动失败: {e}")
                else:
                    target_invoice_name = invoice_filename
                    target_invoice_path = os.path.join(invoice_dir, target_invoice_name)
                    
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
                
                # 处理 BL
                if bl_files:
                    for bl_att in bl_files:
                        bl_path = bl_att['path']
                        bl_filename = os.path.basename(bl_path)
                        
                        if not os.path.exists(bl_path):
                            print(f"  ⚠ BL 文件已不存在（可能已被处理）: {bl_filename}")
                            continue
                        
                        if hbl:
                            target_bl_name = f"BL {hbl}.pdf"
                            target_bl_name = sanitize_filename(target_bl_name)
                            target_bl_path = os.path.join(bl_dir, target_bl_name)
                            
                            counter = 1
                            original_target = target_bl_path
                            while os.path.exists(target_bl_path):
                                name, ext = os.path.splitext(original_target)
                                target_bl_path = f"{name}_{counter}{ext}"
                                counter += 1
                            
                            try:
                                shutil.move(bl_path, target_bl_path)
                                print(f"  ✓ BL 已移动并重命名: {os.path.basename(target_bl_path)}")
                            except Exception as e:
                                print(f"  ✗ BL 移动失败: {e}")
                        else:
                            target_bl_name = f"BL_未知_{bl_filename}"
                            target_bl_path = os.path.join(bl_dir, target_bl_name)
                            
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
        
        if all_excel_data:
            headers = [
                "NO", "File Name", "FILENO", "DATE", "Carrier",
                "Loading Port", "Loading Port Code", "Destination", "Destination Code",
                "ETD", "ETA", "Receipt", "OBL", "HBL", "MBL",
                "Item", "Quantity", "Unit Price", "Container Type", "Amount", "Booking No"
            ]
            
            df = pd.DataFrame(all_excel_data, columns=headers)
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
        
        try:
            if os.path.exists(temp_dir):
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
        traceback.print_exc()
    finally:
        # 恢复标准输出
        sys.stdout = old_stdout


class InvoiceAutoGUI:
    """主 GUI 应用程序类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("自动化单据处理工具")
        self.root.geometry("800x600")
        
        # 保存文件夹路径
        self.save_dir = os.getcwd()  # 默认使用当前运行目录
        
        # 是否正在运行
        self.is_running = False
        
        # 创建界面
        self.create_widgets()
    
    def create_widgets(self):
        """创建 GUI 组件"""
        
        # 顶部：路径选择区域
        path_frame = tk.Frame(self.root, pady=10)
        path_frame.pack(fill=tk.X, padx=10)
        
        tk.Label(path_frame, text="当前保存位置：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        
        self.path_label = tk.Label(path_frame, text=self.save_dir, font=("Microsoft YaHei", 9), 
                                   fg="blue", anchor="w")
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_btn = tk.Button(path_frame, text="浏览文件夹", command=self.browse_folder,
                              font=("Microsoft YaHei", 10), bg="#4CAF50", fg="white", 
                              relief=tk.RAISED, cursor="hand2")
        browse_btn.pack(side=tk.RIGHT)
        
        # 分隔线
        tk.Frame(self.root, height=2, bg="gray").pack(fill=tk.X, padx=10, pady=5)
        
        # 中部：开始处理按钮
        button_frame = tk.Frame(self.root, pady=20)
        button_frame.pack()
        
        self.start_btn = tk.Button(button_frame, text="开始处理", command=self.start_processing,
                                   font=("Microsoft YaHei", 16, "bold"), bg="#2196F3", fg="white",
                                   relief=tk.RAISED, cursor="hand2", padx=30, pady=10)
        self.start_btn.pack()
        
        # 分隔线
        tk.Frame(self.root, height=2, bg="gray").pack(fill=tk.X, padx=10, pady=5)
        
        # 底部：日志显示区域
        log_frame = tk.Frame(self.root)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(log_frame, text="运行日志：", font=("Microsoft YaHei", 10, "bold")).pack(anchor="w")
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, 
                                                   font=("Consolas", 9),
                                                   bg="#f5f5f5", fg="#000000")
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def browse_folder(self):
        """浏览并选择保存文件夹"""
        if self.is_running:
            return
        
        folder = filedialog.askdirectory(initialdir=self.save_dir, title="选择保存文件夹")
        if folder:
            self.save_dir = folder
            self.path_label.config(text=self.save_dir)
            print(f"已选择保存位置: {self.save_dir}")
    
    def start_processing(self):
        """开始处理（在独立线程中运行）"""
        if self.is_running:
            return
        
        # 检查是否选择了文件夹
        if not self.save_dir:
            messagebox.showwarning("警告", "请先选择保存文件夹！")
            return
        
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        # 禁用按钮
        self.start_btn.config(state=tk.DISABLED, text="处理中...")
        self.is_running = True
        
        # 创建日志重定向对象
        log_redirector = TextRedirector(self.log_text)
        
        # 在独立线程中运行主处理函数
        thread = threading.Thread(target=self.run_in_thread, args=(log_redirector,), daemon=True)
        thread.start()
    
    def run_in_thread(self, log_redirector):
        """在线程中运行主处理函数"""
        try:
            run_main_process(self.save_dir, log_redirector)
        finally:
            # 恢复按钮状态（需要在主线程中执行）
            self.root.after(0, self.reset_button)
    
    def reset_button(self):
        """重置按钮状态（在主线程中调用）"""
        self.start_btn.config(state=tk.NORMAL, text="开始处理")
        self.is_running = False


def main():
    """主函数：启动 GUI 应用"""
    root = tk.Tk()
    app = InvoiceAutoGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

