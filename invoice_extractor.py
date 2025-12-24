import json
import requests
import pdfplumber
import re
import time  # 如需使用 sleep，请使用 time.sleep()
import os
import sys
import config_loader

# ================= 配置区域 =================
# 从配置文件加载 API Key
try:
    API_KEY = config_loader.get_api_key()
except (FileNotFoundError, ValueError) as e:
    print(f"[错误] API Key 配置加载失败: {e}")
    print("请确保已创建 config.ini 文件并填写正确的 API Key。")
    API_KEY = ""  # 设置为空字符串，后续调用会失败并提示
# ===========================================

# ================= 港口代码缓存 =================
# 全局变量：缓存加载的港口代码字典，避免重复加载
_PORT_CODES_CACHE = None
# ===========================================

# ================= 港口代码加载函数 =================
def load_port_codes():
    """
    功能：从 port_codes.json 文件加载港口代码字典（带缓存）
    
    返回：
        dict: 港口名称（大写）-> 5位代码 的字典
        如果文件不存在或读取失败，返回空字典并记录错误日志
    """
    global _PORT_CODES_CACHE
    
    # 如果已经加载过，直接返回缓存
    if _PORT_CODES_CACHE is not None:
        return _PORT_CODES_CACHE
    
    try:
        # 判断是否为打包后的 EXE 环境
        if getattr(sys, 'frozen', False):
            # EXE 环境：使用 EXE 文件所在的目录
            base_path = os.path.dirname(sys.executable)
        else:
            # 普通 Python 脚本运行：使用脚本文件所在的目录
            base_path = os.path.dirname(__file__)
        
        json_file = os.path.join(base_path, 'port_codes.json')
        
        if not os.path.exists(json_file):
            print(f"[警告] 港口代码文件不存在: {json_file}")
            print("请确保 port_codes.json 文件存在于项目根目录。")
            _PORT_CODES_CACHE = {}
            return {}
        
        with open(json_file, 'r', encoding='utf-8') as f:
            port_dict = json.load(f)
        
        print(f"[成功] 成功加载港口代码字典，共 {len(port_dict)} 个港口。")
        _PORT_CODES_CACHE = port_dict
        return port_dict
        
    except json.JSONDecodeError as e:
        print(f"[错误] port_codes.json 文件格式错误: {e}")
        _PORT_CODES_CACHE = {}
        return {}
    except Exception as e:
        print(f"[错误] 加载港口代码文件失败: {e}")
        _PORT_CODES_CACHE = {}
        return {}
# ===========================================

def get_port_code(port_name):
    """
    功能：根据港口名称查找对应的 UN/LOCODE 代码（严格匹配模式）
    
    参数：
        port_name: 港口名称（例如 "Shanghai" 或 "SHANGHAI"）
    
    返回：
        对应的 5 位代码，如果找不到则返回空字符串
    """
    # 1. 如果 port_name 为空，返回空字符串
    if not port_name:
        return ""
    
    # 2. 加载港口代码字典
    port_dict = load_port_codes()
    
    # 3. 如果字典为空（文件加载失败），返回空字符串
    if not port_dict:
        return ""
    
    # 4. 标准化输入：去除首尾空格并转换为大写
    normalized_name = str(port_name).strip().upper()
    
    # 5. 严格匹配：直接查表
    code = port_dict.get(normalized_name, "")
    
    if code:
        print(f"DEBUG: 港口匹配成功: [{port_name}] -> [{code}]")
    else:
        print(f"DEBUG: 港口匹配失败: [{port_name}] (标准化后: [{normalized_name}])")
    
    return code

def extract_invoice_data(pdf_path):
    """
    功能：调用 DeepSeek 提取 PDF 中的发票数据
    """
    # 1. 判空检查
    if not pdf_path:
        print("错误：传入的PDF路径是空的")
        return []

    # 2. 读取PDF文字
    print(f"正在读取PDF文件：{pdf_path}")
    full_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
    except Exception as e:
        print(f"读取PDF失败: {e}")
        return []

    if not full_text:
        print("警告：无法提取文本，可能是扫描图片PDF")
        return []

    # 3. 构造提示词
    prompt = """
    你是一个物流单据提取专家。请分析用户的 Invoice 文本。
    该 Invoice 可能包含多行费用明细。
    
    请输出一个 JSON 列表 (List of Objects)。
    如果是单行费用，列表中只有一个对象；如果是多行费用，列表中有多个对象。
    每个对象必须包含【表头信息】和【当前行的费用明细】。

    请严格提取以下字段（Key必须完全一致，找不到填 null）：
    
    【表头通用信息】(每行都要带上):
    - InvoiceNo: (提取 "INVOICE NO" 标题后的编号，通常以 'S' 开头，例如 S2511SED...)
    - OriginalFileNo: (精确提取 "FILE NO." 后面的完整编号，例如：SRSE202508-00631，注意提取完整内容包含所有字符和连字符)
    - DATE: (提取 "DATE" 后的日期，格式 YYYY/MM/DD)
    - Carrier: (提取 "Carrier" 后的内容)
    - loadingport: (提取 "Loading port" 后的港口)
    - Destination: (提取 "Destination" 或 "Discharge port" 后的港口)
    - Vessel_Voyage: (提取 "Vessel" 或 "Vessel / Voyage" 后的完整内容，包含船名和航次，例如："MSC JESSENIA R V. HN531A"，请确保提取完整的字符串)
    - ETD: (提取 "ETD" 后的日期，格式 YYYY/MM/DD)
    - ETADate: (提取 "ETA Date" 后的日期，格式 YYYY/MM/DD)
    - OBL: (提取 "OBL" 后的单号)
    - HBL: (提取 "HBL" 后的单号)
    - Receipt: (提取 "Receipt" 后的地点)

    【费用明细信息】(根据费用行提取):
    - OCEANFREIGHT: (费用项目名称，通常在 Description 列)
    - XUSD: (提取 "X USD" 前面的数量，纯数字)
    - USD: (提取该行的总金额)
    - Unit_Price: (提取单价，如 "2042.000/40' HQ" 中的 "2042.000")
    - Container_Type: (提取柜型，如 "2042.000/40' HQ" 中的 "40' HQ")

    注意：
    1. 即使只有一行数据，也必须返回列表格式 `[{...}]`。
    2. 不要使用 Markdown 格式，直接返回 JSON 字符串。
    3. InvoiceNo 是必填字段，请优先提取 "INVOICE NO" 后的编号。
    """

    print("正在调用 DeepSeek 进行智能提取...")
    
    url = "https://api.deepseek.com/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }
    
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "你是一个精通物流单据的数据提取助手，只输出 JSON。"},
            {"role": "user", "content": prompt + "\n\n【单据内容】:\n" + full_text}
        ],
        "temperature": 0.1
    }

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=60)
        
        if response.status_code == 200:
            res_json = response.json()
            content = res_json['choices'][0]['message']['content']
            
            # 清洗数据
            content = content.replace("```json", "").replace("```", "").strip()
            
            # 解析 JSON
            result_list = json.loads(content)
            
            # 兼容性处理
            if isinstance(result_list, dict):
                result_list = [result_list]
                
            print(f"提取成功！共找到 {len(result_list)} 条费用记录。")
            return result_list
        else:
            print(f"API调用失败: {response.text}")
            return []
            
    except Exception as e:
        print(f"发生代码错误: {e}")
        return []

# =================================================================
# 👇 这里是新增的函数：专门用于把数据组装成 Excel 的一行 (适配 Sheet1)
# =================================================================
def prepare_excel_row(invoice_data, file_path, booking_no):
    """
    功能：适配 info.xlsx - Sheet1 的表头格式
    特点：Carrier不重复，港口后预留Code列
    """
    if not invoice_data:
        invoice_data = {}
    
    def clean(text):
        if not text: return ""
        return str(text).strip()

    # 获取港口代码
    loading_port = clean(invoice_data.get("loadingport"))
    destination = clean(invoice_data.get("Destination"))
    loading_port_code = get_port_code(loading_port)
    dest_port_code = get_port_code(destination)

    # 组装列表 (严格对应 main.py 中 headers 列表的顺序，确保一一对应)
    # FILENO 列优先使用 InvoiceNo，如果为空则使用 OriginalFileNo
    fileno_value = clean(invoice_data.get("InvoiceNo")) or clean(invoice_data.get("OriginalFileNo"))
    
    # 从 invoice_data 中获取 OriginalFileNo 和 Vessel_Voyage 的值
    original_file_no = clean(invoice_data.get("OriginalFileNo"))
    vessel_voyage = clean(invoice_data.get("Vessel_Voyage"))
    
    row = [
        "2",                                    # 1. NO (固定值)
        file_path,                              # 2. File Name
        fileno_value,                           # 3. FILENO (优先 InvoiceNo，否则 OriginalFileNo)
        original_file_no,                       # 4. File No (从 OriginalFileNo 字段获取)
        clean(invoice_data.get("DATE")),        # 5. DATE
        
        # --- 船公司信息 ---
        clean(invoice_data.get("Carrier")),     # 6. Carrier
        vessel_voyage,                          # 7. Vessel/Voyage (从 Vessel_Voyage 字段获取)
        
        # --- 装货港 (名称 + 代码) ---
        clean(invoice_data.get("loadingport")), # 8. Loading Port
        loading_port_code,                      # 9. Loading Port Code (自动匹配)
        
        # --- 目的港 (名称 + 代码) ---
        clean(invoice_data.get("Destination")), # 10. Destination
        dest_port_code,                         # 11. Destination Code (自动匹配)
        
        # --- 时间与单号 ---
        clean(invoice_data.get("ETD")),         # 12. ETD
        clean(invoice_data.get("ETADate")),     # 13. ETA
        clean(invoice_data.get("Receipt")),     # 14. Receipt
        clean(invoice_data.get("OBL")),         # 15. OBL
        clean(invoice_data.get("HBL")),         # 16. HBL
        clean(invoice_data.get("MBL")),         # 17. MBL (预留位置)
        
        # --- 费用明细 ---
        clean(invoice_data.get("OCEANFREIGHT")),# 18. Item
        clean(invoice_data.get("XUSD")),        # 19. Quantity
        clean(invoice_data.get("Unit_Price")),  # 20. Unit Price
        clean(invoice_data.get("Container_Type")),# 21. Container Type
        clean(invoice_data.get("USD")),         # 22. Amount
        
        # --- 邮件提取 ---
        booking_no                              # 23. Booking No (从邮件中提取)
    ]
    
    return row

def main(args):
    # 测试代码
    pass


