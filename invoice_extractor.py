import json
import requests
import pdfplumber
import re
import time  # 如需使用 sleep，请使用 time.sleep()

# ================= 配置区域 =================
# 🔴 你的 API Key (已保留你刚才提供的)
API_KEY = "sk-cb441e489cd84dc8906e37733ed9181e" 
# ===========================================

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
    - OriginalFileNo: (提取 "FILE NO." 后的编号)
    - DATE: (提取 "DATE" 后的日期，格式 YYYY/MM/DD)
    - Carrier: (提取 "Carrier" 后的内容)
    - loadingport: (提取 "Loading port" 后的港口)
    - Destination: (提取 "Destination" 或 "Discharge port" 后的港口)
    - Vessel: (提取 "Vessel" 后的船名)
    - ETD: (提取 "ETD" 后的日期)
    - ETADate: (提取 "ETA Date" 后的日期)
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

    # 组装列表 (严格对应 Sheet1 的 A, B, C... 列顺序)
    # FILENO 列优先使用 InvoiceNo，如果为空则使用 OriginalFileNo
    fileno_value = clean(invoice_data.get("InvoiceNo")) or clean(invoice_data.get("OriginalFileNo"))
    
    row = [
        "2",                                    # A列: NO (固定)
        file_path,                              # B列: File Name
        fileno_value,                           # C列: FILENO (优先 InvoiceNo，否则 OriginalFileNo)
        clean(invoice_data.get("DATE")),        # D列: DATE
        
        # --- 船公司 (只填一次) ---
        clean(invoice_data.get("Carrier")),     # E列: Carrier
        
        # --- 装货港 (名称 + 代码) ---
        clean(invoice_data.get("loadingport")), # F列: Loading Port
        "",                                     # G列: Loading Port Code (留空，后期人工填)
        
        # --- 目的港 (名称 + 代码) ---
        clean(invoice_data.get("Destination")), # H列: Destination
        "",                                     # I列: Destination Code (预留空位)
        
        # --- 时间与单号 ---
        clean(invoice_data.get("ETD")),         # J列: ETD
        clean(invoice_data.get("ETADate")),     # K列: ETA
        clean(invoice_data.get("Receipt")),     # L列: Receipt
        clean(invoice_data.get("OBL")),         # M列: OBL
        clean(invoice_data.get("HBL")),         # N列: HBL
        clean(invoice_data.get("MBL")),         # O列: MBL (预留位置)
        
        # --- 费用明细 ---
        clean(invoice_data.get("OCEANFREIGHT")),# P列: Item/Description
        clean(invoice_data.get("XUSD")),        # Q列: Quantity
        clean(invoice_data.get("Unit_Price")),  # R列: Unit Price
        clean(invoice_data.get("Container_Type")),# S列: Container Type
        clean(invoice_data.get("USD")),         # T列: Amount
        
        # --- 邮件提取 ---
        booking_no                              # U列: Booking No (自动填入)
    ]
    
    return row

def main(args):
    # 测试代码
    pass


