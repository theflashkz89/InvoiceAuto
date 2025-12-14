"""
邮件处理模块
用于从 QQ 邮箱下载 PDF 附件并自动分类处理
"""

import os
import re
from imap_tools import MailBox, AND
from PDFClassifier import classify_pdf_content


def download_and_process_attachments(username, password, save_root_dir):
    """
    从 QQ 邮箱下载未读邮件的 PDF 附件，并根据内容分类处理
    
    参数:
        username (str): QQ 邮箱账号
        password (str): QQ 邮箱授权码（不是登录密码）
        save_root_dir (str): 保存附件的根目录路径
        
    返回:
        list: 邮件列表，每个元素是一个字典，代表一封邮件，包含：
            - "subject": 邮件标题
            - "body": 邮件正文
            - "booking_no": 提取到的订舱号（从邮件正文中提取）
            - "attachments": 这封邮件下的有效附件列表，每个附件包含：
                - "type": 文件类型（"INVOICE"、"BL" 或 "UNKNOWN"）
                - "path": 文件保存路径
            注意：如果一封邮件里没有有效附件（都被分类为 IGNORE 或没附件），
            则不会把这封邮件加入返回列表。
            
    异常:
        如果连接邮箱失败，会打印错误信息并返回空列表
    """
    # 确保保存目录存在
    if not os.path.exists(save_root_dir):
        os.makedirs(save_root_dir)
        print(f"已创建保存目录: {save_root_dir}")
    
    # 存储邮件列表（以邮件为单位分组）
    email_list = []
    
    try:
        # 连接到 QQ 邮箱 IMAP 服务器
        print(f"正在连接到 QQ 邮箱: {username}")
        with MailBox('imap.qq.com').login(username, password) as mailbox:
            print("连接成功！")
            
            # 获取所有未读邮件
            print("正在获取未读邮件...")
            unseen_emails = mailbox.fetch(AND(seen=False))
            
            email_count = 0
            attachment_count = 0
            
            # 遍历每封未读邮件
            for email in unseen_emails:
                email_count += 1
                email_subject = email.subject
                email_body = email.text or email.html or ""  # 获取邮件正文（优先文本，其次HTML）
                print(f"\n处理邮件 {email_count}: {email_subject}")
                
                # 从邮件标题中提取 Order No
                order_no = ""
                order_pattern = r"ORDER NO\s*([A-Za-z0-9]+)"
                order_match = re.search(order_pattern, email_subject, re.IGNORECASE)
                if order_match:
                    order_no = order_match.group(1)
                    print(f"  提取到 Order No: {order_no}")
                
                # 从邮件正文中提取 Booking No
                booking_no = ""
                # 匹配 "ORDER nbr : xxx" 或 "Booking No : xxx" 或 "Order No : xxx"
                # (?:...) 是非捕获组，兼容 nbr/No/Ref，忽略大小写
                booking_pattern = r"(?:ORDER|Booking)\s*(?:nbr|No|Ref|#)?\s*[:\.]?\s*([A-Za-z0-9\-\/]+)"
                booking_match = re.search(booking_pattern, email_body, re.IGNORECASE)
                if booking_match:
                    booking_no = booking_match.group(1)
                    print(f"  提取到 Booking No: {booking_no}")
                
                # 存储当前邮件的有效附件
                valid_attachments = []
                
                # 遍历每封邮件的每个附件
                for attachment in email.attachments:
                    attachment_filename = attachment.filename
                    
                    # 跳过包含 "bank detail" 或 "bank_detail" 的文件（忽略大小写）
                    if "bank detail" in attachment_filename.lower() or "bank_detail" in attachment_filename.lower():
                        print(f"  ⏭ 已跳过 Bank Detail 文件: {attachment_filename}")
                        continue
                    
                    # 只处理 PDF 文件（忽略大小写）
                    if attachment_filename.lower().endswith('.pdf'):
                        attachment_count += 1
                        print(f"  发现 PDF 附件: {attachment_filename}")
                        
                        # 构建保存路径
                        file_path = os.path.join(save_root_dir, attachment_filename)
                        
                        # 如果文件已存在，添加序号避免覆盖
                        counter = 1
                        original_path = file_path
                        while os.path.exists(file_path):
                            name, ext = os.path.splitext(original_path)
                            file_path = f"{name}_{counter}{ext}"
                            counter += 1
                        
                        try:
                            # 下载附件
                            print(f"  正在下载到: {file_path}")
                            with open(file_path, 'wb') as f:
                                f.write(attachment.payload)
                            
                            # 立即调用分类器识别文件类型
                            print(f"  正在分类文件...")
                            file_type = classify_pdf_content(file_path)
                            print(f"  分类结果: {file_type}")
                            
                            # 根据分类结果处理文件
                            if file_type == "IGNORE":
                                # 删除垃圾文件
                                os.remove(file_path)
                                print(f"  ✓ 已删除垃圾文件: {attachment_filename}")
                            else:
                                # 保留文件，添加到当前邮件的有效附件列表
                                attachment_info = {
                                    "type": file_type,
                                    "path": file_path
                                }
                                valid_attachments.append(attachment_info)
                                print(f"  ✓ 已保留文件: {attachment_filename} (类型: {file_type})")
                                
                        except Exception as e:
                            print(f"  ✗ 处理附件时出错 {attachment_filename}: {str(e)}")
                            # 如果下载失败，尝试删除可能已创建的文件
                            if os.path.exists(file_path):
                                try:
                                    os.remove(file_path)
                                except:
                                    pass
                
                # 如果当前邮件有有效附件，则添加到邮件列表
                if valid_attachments:
                    email_info = {
                        "subject": email_subject,
                        "body": email_body,
                        "booking_no": booking_no,
                        "attachments": valid_attachments
                    }
                    email_list.append(email_info)
                    print(f"  ✓ 邮件已添加到结果列表（包含 {len(valid_attachments)} 个有效附件）")
                else:
                    print(f"  - 邮件无有效附件，已跳过")
            
            print(f"\n处理完成！共处理 {email_count} 封邮件，{attachment_count} 个 PDF 附件")
            print(f"有效邮件数量: {len(email_list)}")
            
    except Exception as e:
        print(f"错误：连接或处理邮箱时出错: {str(e)}")
        print("提示：请确保：")
        print("  1. QQ 邮箱已开启 IMAP 服务")
        print("  2. 使用的是授权码（不是登录密码）")
        print("  3. 网络连接正常")
    
    return email_list

