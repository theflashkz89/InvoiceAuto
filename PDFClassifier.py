"""
PDF 文件分类器模块
用于识别 PDF 文件的类型（发票、提单等）
"""

import pdfplumber
import os


def classify_pdf_content(file_path):
    """
    根据 PDF 文件内容识别文件类型
    
    参数:
        file_path (str): PDF 文件的路径
        
    返回:
        str: 文件类型标识
            - "INVOICE": 发票文件（包含 "INVOICE"、"DEBIT NOTE"、"TAX RECEIPT"、"PAYMENT REQUEST"、"CREDIT NOTE" 等）
            - "BL": 提单文件（包含 "BILL OF LADING"、"WAYBILL"、"TELEX RELEASE"、"CARGO RECEIPT" 等，或文件名包含 HBL/MBL）
            - "IGNORE": 垃圾文件（包含 "BANK DETAILS"）
            - "UNKNOWN": 未知类型
            
    判定逻辑:
        1. 优先检查文件名（HBL/MBL 且不含 INVOICE -> BL）
        2. 冲突仲裁：同时包含发票和提单关键词时，通过金额特征词（TOTAL/AMOUNT 等）判断
        3. 有金额特征词 -> 判定为发票；无金额特征词 -> 判定为提单
            
    异常:
        如果文件无法打开或读取，会返回 "UNKNOWN" 并打印错误信息
    """
    pdf_file = None
    try:
        # 打开 PDF 文件
        pdf_file = pdfplumber.open(file_path)
        
        # 检查是否有页面
        if len(pdf_file.pages) == 0:
            return "UNKNOWN"
        
        # 读取第一页的文本内容
        first_page = pdf_file.pages[0]
        text_content = first_page.extract_text()
        
        # 如果无法提取文本，返回 UNKNOWN
        if text_content is None:
            return "UNKNOWN"
        
        # 【调试信息】打印前100个字符（去除换行符）
        debug_text = text_content.replace('\n', ' ').replace('\r', ' ').strip()[:100]
        print(f"[DEBUG文本]: {debug_text}")
        
        # 将文本转换为大写，便于不区分大小写的匹配
        text_upper = text_content.upper()
        filename_upper = os.path.basename(file_path).upper()
        
        # 判断是否为垃圾文件：包含 "BANK DETAILS"（优先判断，防止误判为其他类型）
        if "BANK DETAILS" in text_upper:
            return "IGNORE"
        
        # 1. 定义关键词库
        # 发票标题词
        invoice_keywords = ["INVOICE", "DEBIT NOTE", "TAX RECEIPT", "PAYMENT REQUEST", "CREDIT NOTE"]
        # 提单标题词 (注意：不要用 HBL 这种短词作为正文关键词，容易误判，要用全称)
        bl_keywords = ["BILL OF LADING", "WAYBILL", "TELEX RELEASE", "CARGO RECEIPT"]
        # 算钱特征词 (发票通常会有这些，提单通常没有)
        money_keywords = ["TOTAL", "AMOUNT DUE", "GRAND TOTAL", "BALANCE", "SUBTOTAL"]
        
        # 2. 状态检测
        is_invoice = any(kw in text_upper for kw in invoice_keywords)
        is_bl = any(kw in text_upper for kw in bl_keywords)
        has_money = any(kw in text_upper for kw in money_keywords)
        
        # 3. 优先基于文件名的强力辅助判决 (文件名往往最准)
        if "HBL" in filename_upper and "INVOICE" not in filename_upper:
            return "BL"
        if "MBL" in filename_upper and "INVOICE" not in filename_upper:
            return "BL"
        
        # 4. 冲突仲裁逻辑
        if is_invoice and is_bl:
            # 既像发票又像提单 (最常见情况：发票里写了 Bill of Lading No)
            # 新增：检查 "BILL OF LADING" 是否作为文档标题出现（在前 500 字符内）
            # 如果是，说明这是一个真正的提单文件，而不是发票中引用了提单号
            first_500_chars = text_upper[:500]
            if "BILL OF LADING" in first_500_chars:
                print(f"[DEBUG分类]: 检测到 BILL OF LADING 在文件开头，判定为 BL")
                return "BL"  # 标题是 Bill of Lading -> 认为是提单
            
            if has_money:
                return "INVOICE"  # 有"TOTAL/AMOUNT" -> 认为是发票
            else:
                return "BL"       # 没谈钱 -> 认为是提单附件
        
        if is_invoice:
            return "INVOICE"
            
        if is_bl:
            return "BL"
        
        # 如果都不匹配，返回未知类型
        return "UNKNOWN"
        
    except Exception as e:
        # 处理文件打开或读取异常
        print(f"错误：无法读取 PDF 文件 {file_path}: {str(e)}")
        return "UNKNOWN"
        
    finally:
        # 确保文件流被关闭
        if pdf_file is not None:
            try:
                pdf_file.close()
            except Exception as e:
                print(f"警告：关闭 PDF 文件时出错 {file_path}: {str(e)}")

