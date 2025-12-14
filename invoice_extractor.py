import json
import requests
import pdfplumber
import re
import time  # 如需使用 sleep，请使用 time.sleep()
import difflib
import config_loader

# ================= 配置区域 =================
# 从配置文件加载 API Key
try:
    API_KEY = config_loader.get_api_key()
except (FileNotFoundError, ValueError) as e:
    print(f"❌ API Key 配置加载失败: {e}")
    print("请确保已创建 config.ini 文件并填写正确的 API Key。")
    API_KEY = ""  # 设置为空字符串，后续调用会失败并提示
# ===========================================

# ================= 港口代码字典 =================
PORT_CODES = {
    # 中国大陆主要港口
    'SHANGHAI': 'CNSHG',
    'NINGBO': 'CNNGB',
    'SHENZHEN': 'CNSZX',
    'GUANGZHOU': 'CNCAN',
    'QINGDAO': 'CNTAO',
    'TIANJIN': 'CNTXG',
    'DALIAN': 'CNDLC',
    'XIAMEN': 'CNXMN',
    'YANTIAN': 'CNYTN',
    'SHEKOU': 'CNSHK',
    'CHIWAN': 'CNCWN',
    'FOSHAN': 'CNFOS',
    'ZHONGSHAN': 'CNZSN',
    'ZHUHAI': 'CNZUH',
    'BEIJING': 'CNBJO',
    'NANJING': 'CNNJG',
    'WUHAN': 'CNWUH',
    'CHONGQING': 'CNCKG',
    'FOS': 'CNFOS',
    'ANQING': 'CNAQG',
    'CHANGSHA': 'CNCSX',
    'CHANGGU': 'CNCGU',
    'FANGCHENG': 'CNFAN',
    'FUZHOU': 'CNFOC',
    'GONGZHULING': 'CNGON',
    'HUADU': 'CNHUA',
    'JIANGMEN': 'CNJMN',
    'JINGZHOU': 'CNJGZ',
    'JIUJIANG': 'CNJIU',
    'LIANYUNGANG': 'CNLYG',
    'KUNSHAN': 'CNKHN',
    'NANSHA': 'CNNSA',
    'NANTONG': 'CNNTG',
    'QUANZHOU': 'CNQZH',
    'SHANWEI': 'CNSWA',
    'SUDONG': 'CNSUD',
    'TAICANG': 'CNTAG',
    'TAIZHOU': 'CNTZO',
    'WEIHAI': 'CNWEI',
    'WENZHOU': 'CNWNZ',
    'WUHU': 'CNWHI',
    'XIGANG': 'CNXIG',
    'YANGZHOU': 'CNYZH',
    'YICHANG': 'CNYIC',
    'YUYUAN': 'CNYUY',
    'ZHANGJIAGANG': 'CNZJG',
    'ZHAOQING': 'CNZHA',
    'ZHEJIANG': 'CNZHE',
    'TANGSHAN': 'CNTSG',
    'LUZHOU': 'CNLUZ',
    'RIZHAO': 'CNROQ',
    'SANJIAO': 'CNSJQ',
    'XILING': 'CNXIL',
    'YANTING': 'CNYTG',
    'ZHUQING': 'CNZQG',
    'DONGCUOBA': 'CNDCB',
    'LANZHOU': 'CNLAN',
    
    # 香港、澳门、台湾
    'HONGKONG': 'HKHKG',
    'MACAU': 'MOMFM',
    'MACAO': 'MOMFM',
    'TAIPEI': 'TWTPE',
    'KAOHSIUNG': 'TWKHH',
    'KEELUNG': 'TWKEL',
    'TAICHUNG': 'TWTXG',
    'TAOYUAN': 'TWTYN',
    
    # 美国主要港口
    'LOSANGELES': 'USLAX',
    'LONGBEACH': 'USLGB',
    'NEWYORK': 'USNYC',
    'NEWARK': 'USEWR',
    'SAVANNAH': 'USSAV',
    'CHARLESTON': 'USCHS',
    'HOUSTON': 'USHOU',
    'MIAMI': 'USMIA',
    'SEATTLE': 'USSEA',
    'TACOMA': 'USTIW',
    'OAKLAND': 'USOAK',
    'NORFOLK': 'USORF',
    'BALTIMORE': 'USBWI',
    'BOSTON': 'USBOS',
    'PHILADELPHIA': 'USPHL',
    'CHICAGO': 'USCHI',
    'DETROIT': 'USDET',
    'PORTLAND': 'USPDX',
    
    # 欧洲主要港口
    'ROTTERDAM': 'NLRTM',
    'AMSTERDAM': 'NLAMS',
    'ANTWERP': 'BEANR',
    'HAMBURG': 'DEHAM',
    'BREMEN': 'DEBRE',
    'FELIXSTOWE': 'GBFXT',
    'LONDON': 'GBLON',
    'SOUTHAMPTON': 'GBSOU',
    'LEHAVRE': 'FRLEH',
    'MARSEILLE': 'FRMRS',
    'FOSSURMER': 'FRFOS',
    'BARCELONA': 'ESBCN',
    'VALENCIA': 'ESVLC',
    'ALGECIRAS': 'ESALG',
    'GENOA': 'ITGOA',
    'LASPEZIA': 'ITSPE',
    'NAPLES': 'ITNAP',
    'GIOIATAURO': 'ITGIT',
    'PIRAEUS': 'GRPIR',
    'THESSALONIKI': 'GRSKG',
    'GOTHENBURG': 'SEGOT',
    'STOCKHOLM': 'SESTO',
    'GDANSK': 'PLGDN',
    'GDYNIA': 'PLGDY',
    'DUNKIRK': 'FRDKK',
    'WILHELMSHAVEN': 'DEWVN',
    'ZEEBRUGGE': 'BEZEE',
    'BREMERHAVEN': 'DEBRV',
    'GATEWAY': 'GBLGP',
    'IMMINGHAM': 'GBIMM',
    'BELFAST': 'GBBEL',
    'COPENHAGEN': 'DKCPH',
    'AARHUS': 'DKAAR',
    'OSLO': 'NOOSL',
    'DUBLIN': 'IEDUB',
    'CORK': 'IECORK',
    'LISBON': 'PTLIS',
    'OPORTO': 'PTOPO',
    
    # 东南亚主要港口
    'SINGAPORE': 'SGSIN',
    'KELANG': 'MYPKG',
    'PENANG': 'MYPEN',
    'PASIRGUDANG': 'MYPGU',
    'BANGKOK': 'THBKK',
    'LAEMCHABANG': 'THLCH',
    'LAEMKRABANG': 'THLKR',
    'SONGKHLA': 'THSGZ',
    'HOCHIMINH': 'VNSGN',
    'HAIPHONG': 'VNHPH',
    'DANANG': 'VNDAD',
    'QUYNHON': 'VNUIH',
    'VUNGTAU': 'VNVUT',
    'CAMRANH': 'VNCMT',
    'MANILA': 'PHMNL',
    'CEBU': 'PHCEB',
    'CAGAYAN': 'PHCGY',
    'GENERALSANTOS': 'PHGES',
    'JAKARTA': 'IDJKT',
    'PANJANG': 'IDPNJ',
    'SURABAYA': 'IDSUB',
    'BELAWAN': 'IDBLW',
    'SEMARANG': 'IDSRG',
    'BATAM': 'IDBTH',
    'YANGON': 'MMRGN',
    'PHNOMPENH': 'KHPNH',
    'SIHANOUKVILLE': 'KHSCH',
    
    # 日韩主要港口
    'TOKYO': 'JPTYO',
    'YOKOHAMA': 'JPYOK',
    'OSAKA': 'JPOSA',
    'KOBE': 'JPUKB',
    'NAGOYA': 'JPNGO',
    'BUSAN': 'KRPUS',
    'INCHEON': 'KRINC',
    'ULSAN': 'KRUSN',
    'GWANGYANG': 'KRKAN',
    
    # 中东主要港口
    'DUBAI': 'AEDXB',
    'JEBELALI': 'AEJEA',
    'ABUDHABI': 'AEAUH',
    'DAMMAM': 'SADMM',
    'JEDDAH': 'SAJED',
    'RIYADH': 'SARUH',
    'KUWAIT': 'KWKWI',
    'DOHA': 'QADOH',
    'BAHRAIN': 'BHBAH',
    'MUSCAT': 'OMMCT',
    'BANDARABBAS': 'IRBND',
    'ASHDOD': 'ILASH',
    'HAIFA': 'ILHFA',
    
    # 南亚主要港口
    'MUMBAI': 'INBOM',
    'NEWDELHI': 'INNDE',
    'CHENNAI': 'INMAA',
    'KOLKATA': 'INCCU',
    'COCHIN': 'INCOK',
    'VISAKHAPATNAM': 'INVTZ',
    'KARACHI': 'PKKHI',
    'LAHORE': 'PKLHE',
    'COLOMBO': 'LKCMB',
    'CHITTAGONG': 'BDCGP',
    'DHAKA': 'BDDAC',
    
    # 澳洲主要港口
    'SYDNEY': 'AUSYD',
    'MELBOURNE': 'AUMEL',
    'BRISBANE': 'AUBNE',
    'FREMANTLE': 'AUFRE',
    'ADELAIDE': 'AUADL',
    'AUCKLAND': 'NZAKL',
    'WELLINGTON': 'NZWLG',
    'LYTTELTON': 'NZLYT',
    
    # 南美主要港口
    'SANTOS': 'BRSSZ',
    'RIODEJANEIRO': 'BRRIO',
    'BUENOSAIRES': 'ARBUE',
    'VALPARAISO': 'CLVAP',
    'CALLAO': 'PECLL',
    'CARTAGENA': 'COCTG',
    'MANZANILLO': 'MXZLO',
    'VERACRUZ': 'MXVER',
    
    # 非洲主要港口
    'DURBAN': 'ZADUR',
    'CAPETOWN': 'ZACPT',
    'ELIZABETH': 'ZAPLZ',
    'CASABLANCA': 'MACAS',
    'ALEXANDRIA': 'EGALY',
    'SAID': 'EGPSD',
    'LAGOS': 'NGLOS',
    'MOMBASA': 'KEMBA',
    'DARESSALAAM': 'TZDAR',
    
    # 其他重要港口
    'VANCOUVER': 'CAVAN',
    'TORONTO': 'CATOR',
    'MONTREAL': 'CAMTR',
    'HALIFAX': 'CAHAL',
    'VLADIVOSTOK': 'RUVVO',
    'STPETERSBURG': 'RULED',
    'MURMANSK': 'RUMMK',
}
# ===========================================

def get_port_code(port_name):
    """
    功能：根据港口名称查找对应的 UN/LOCODE 代码（关键词扫描模式）
    
    参数：
        port_name: 港口名称（可能是 "Shanghai" 或 "Shanghai, China"）
    
    返回：
        对应的 5 位代码，如果找不到则返回空字符串
    """
    print(f"DEBUG: 正在查找港口: [{port_name}]")
    
    # 1. 如果 port_name 为空，返回空字符串
    if not port_name:
        return ""
    
    # 2. 将输入的 port_name 转换为全大写字符串
    input_str = port_name.upper()
    
    # 3. 核心逻辑：遍历 PORT_CODES 字典的所有 Key
    #    先将字典的 Key 按长度从长到短排序，优先匹配长关键词
    #    防止误判（比如防止 "AN" 匹配到 "TIANJIN"）
    sorted_keys = sorted(PORT_CODES.keys(), key=len, reverse=True)
    
    # 4. 检查：如果 Key 存在于 input_str 中（子字符串匹配）
    for key in sorted_keys:
        if key in input_str:
            print(f"DEBUG: >> 包含匹配成功! 关键词=[{key}] -> 代码=[{PORT_CODES[key]}]")
            return PORT_CODES[key]
    
    # 5. 如果循环结束还没匹配到，使用 difflib.get_close_matches 进行模糊匹配
    #    cutoff 设为 0.7，匹配最接近的 Key
    close_matches = difflib.get_close_matches(input_str, PORT_CODES.keys(), n=1, cutoff=0.7)
    if close_matches:
        print(f"DEBUG: >> 模糊匹配成功! 输入=[{input_str}] 接近=[{close_matches[0]}]")
        return PORT_CODES[close_matches[0]]
    
    # 6. 如果都没找到，返回空字符串
    print(f"DEBUG: >> 查找失败，字典中无此港口: [{port_name}]")
    return ""

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
        loading_port_code,                      # G列: Loading Port Code (自动匹配)
        
        # --- 目的港 (名称 + 代码) ---
        clean(invoice_data.get("Destination")), # H列: Destination
        dest_port_code,                         # I列: Destination Code (自动匹配)
        
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


