"""
PDF 文件分类器模块
用于识别 PDF 文件的类型（发票、提单等）
"""

import pdfplumber


def classify_pdf_content(file_path):
    """
    根据 PDF 文件内容识别文件类型
    
    参数:
        file_path (str): PDF 文件的路径
        
    返回:
        str: 文件类型标识
            - "INVOICE": 发票文件（包含 "INVOICE"）
            - "BL": 提单文件（包含 "BILL OF LADING" 或 "WAYBILL" 或 "PORT OF"）
            - "IGNORE": 垃圾文件（包含 "BANK DETAILS"）
            - "UNKNOWN": 未知类型
            
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
        
        # 判断是否为垃圾文件：包含 "BANK DETAILS"（优先判断，防止误判为其他类型）
        if "BANK DETAILS" in text_upper:
            return "IGNORE"
        
        # 判断是否为发票：只要包含 "INVOICE" 就判定为发票（放宽规则，去掉 AMOUNT 要求）
        if "INVOICE" in text_upper:
            return "INVOICE"
        
        # 判断是否为提单：包含 "BILL OF LADING" 或 "WAYBILL" 或 "PORT OF"（增加关键词）
        if "BILL OF LADING" in text_upper or "WAYBILL" in text_upper or "PORT OF" in text_upper:
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

