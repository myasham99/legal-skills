import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime
from io import BytesIO

# ==========================================
# 1. 核心配置与常量库 (Config & Constants)
# ==========================================

# 核心路径配置
D_ROOT = Path("D:/【用户名】")
D_LITIGATION = D_ROOT / "【诉讼】"
D_DEFENDANT = D_LITIGATION / "【被诉案件】"
D_JUDGMENT = D_LITIGATION / "【判决裁定】"
D_PRESERVATION = D_LITIGATION / "【财产保全】"
D_EXECUTION = D_LITIGATION / "【执行履行】"
D_PLAINTIFF = D_LITIGATION / "起诉案件"

Z_ROOT = Path("Z:/5 案件管理/0案件收文2026年/02待指派律师的案件")
Z_BRANCH = "中山"  # Z盘指派路径中的分公司标识，所有指派案件统一使用

# OCR配置（用于扫描件自动识别）
TESSERACT_PATH = r"D:\Tesseract\tesseract.exe"
OCR_LANG = "chi_sim+eng"

# 排除的己方关键词
EXCLUDE_PARTIES = ["××××公司", "公司", "珠海分公司", "××分公司", "广州分公司", "广东总承包", "基础设施分公司", "城市管网"]

# 材料简称库
MATERIAL_PREFIXES = {
    # 特定长名优先（先匹配，避免被通用短键抢占）
    "诉讼支付金额明细表": "诉讼明细", "支付金额明细": "诉讼明细", "诉讼明细": "诉讼明细",
    # 诉讼核心
    "起诉状": "起诉材料", "起诉材料": "起诉材料", "答辩状": "答辩状", "答辩意见": "答辩意见",
    "上诉状": "上诉状",
    "反诉状": "反诉状", "证据": "证据材料", "质证": "质证意见", "数据确认": "数据确认",
    "二审判决书": "二审判决", "一审判决书": "一审判决", "判决书": "一审判决",
    "调解书": "调解书", "裁定书": "裁定", "保全结果通知书": "保全通知",
    "冻结证明": "冻结证明", "解封申请": "解封申请", "保全申请": "保全申请",
    "解除财产保全": "解封申请", "解封裁定": "解封裁定", "执行异议": "执行异议", "执行异议申请书": "执行异议",
    "送达地址确认书": "送达地址确认书", "律师函": "律师函",
    # 程序性
    "传票": "开庭传票", "开庭传票": "开庭传票", "授权委托书": "授权委托书",
    "授权材料": "授权材料", "管辖异议": "管辖异议", "延期开庭": "延期开庭",
    # 执行与履行
    "债务计算": "债务计算", "执行款": "执行款", "上诉费": "上诉费", "诉讼费": "诉讼费",
    "执行申请": "执行申请", "执行裁定": "执行裁定", "终本裁定": "终本裁定",
    # 案件管理与其他
    "案件策划": "案件策划", "刑事侦查": "刑事侦查", "和解协议": "和解协议",
    "调解协议": "和解协议", "鉴定报告": "鉴定报告", "庭审笔录": "庭审笔录", "代理词": "代理词"
}

# 完整项目缩写映射表
PROJECT_MAPPING = [
    {"code": "1220138904", "full": "珠海万科海上城市项目三期土建施工总承包工程", "short": "海三", "profit_center": "L400001890"},
    {"code": "1220138904", "full": "珠海万科海上城市项目三期土建施工总承包工程", "short": "星光海岸", "profit_center": "L400001890"},
    {"code": "1220186711", "full": "声博士声学产业基地及负分贝实验室项目基坑支护及土方开挖工程", "short": "声博士", "profit_center": ""},
    {"code": "1220171470", "full": "新青年城市之家项目", "short": "新青年", "profit_center": "L400100331"},
    {"code": "1220188797", "full": "中山市石岐街道方基冲新村片区老旧小区改造工程项目", "short": "方基冲新村", "profit_center": ""},
    {"code": "1220199661", "full": "鹤山国机南联摩托车工业有限公司年产50万辆摩托车建设项目", "short": "升仕摩托50", "profit_center": ""},
    {"code": "1220208099", "full": "广东升仕工业集团（鹤山）有限公司年产30万辆摩托车新建项目", "short": "升仕摩托30", "profit_center": ""},
    {"code": "1220198713", "full": "集宁区正域现代高科技农业产业示范园项目地块二、三部分EPC项目", "short": "正域农业", "profit_center": ""},
    {"code": "1220185435", "full": "新能源汽车超级充电基础设施建设项目", "short": "新能源超充", "profit_center": ""},
    {"code": "1220197189", "full": "龙华区福城街道田背工业区城市更新项目02、03地块总承包(EPC)工程", "short": "澜汇云庭", "profit_center": ""},
    {"code": "1220168297", "full": "诚辉健康科技创新园项目", "short": "诚辉", "profit_center": ""},
    {"code": "1210057115", "full": "江门里仁中学", "short": "里仁中学", "profit_center": "L400001274"},
    {"code": "1210057215", "full": "江门农林双朗小学三期", "short": "双朗小学", "profit_center": "L400001270"},
    {"code": "1220131045", "full": "珠海索卡体育文旅综合产业基地一期工程", "short": "索卡", "profit_center": "L400001706"},
    {"code": "1220188815", "full": "OPPO智能制造中心地块一南区、地块三北区淤泥固化及基坑支护工程", "short": "OPPO-C", "profit_center": ""},
    {"code": "1220185042", "full": "华大时空中心华大股份A区管理办公室装修工程", "short": "华大时空", "profit_center": ""},
    {"code": "1220145252", "full": "深圳市龙华区福城街道田背工业区城市更新项目三标段05地块总承包工程", "short": "澜汇云庭", "profit_center": ""},
    {"code": "1210056734", "full": "深圳华大基因中心", "short": "华大基因", "profit_center": ""},
    {"code": "1220164465", "full": "鄱阳科技园项目总承包工程（B标段）", "short": "鄱阳B", "profit_center": ""},
    {"code": "1220168962", "full": "鼎湖科技园项目之总承包工程(项目部5)", "short": "鼎湖5", "profit_center": ""},
    {"code": "1210057239", "full": "江门紫茶中学项目", "short": "紫茶中学", "profit_center": "L400001272"},
    {"code": "1210057301", "full": "珠海保利碧桂园海棠花园", "short": "海棠花园", "profit_center": "L400000497"},
    {"code": "1210057336", "full": "江门紫茶学校北校区续建小学项目", "short": "紫茶小学", "profit_center": "L400001271"},
    {"code": "1220121905", "full": "东莞市万江区保利天际花园项目二标段土建及水电安装工程项目", "short": "万江二标", "profit_center": "L400001618"},
    {"code": "1220129615", "full": "OPPO长安研发中心项目6号及7号楼总承包工程", "short": "OPPO研发", "profit_center": "L400001677"},
    {"code": "1220130177", "full": "台山市保利公馆项目四期工程土建暨水电安装工程项目", "short": "保利和公馆", "profit_center": "L400001691"},
    {"code": "1220136271", "full": "江门市保利国际广场服务型公寓C、D栋、4#商业、二期地下室", "short": "保利国际广场二期", "profit_center": "L400001799"},
    {"code": "1220137306", "full": "江门御海阳光花园项目二期项目总承包工程", "short": "御海阳光", "profit_center": "L400001830"},
    {"code": "1220154852", "full": "明阳智慧能源集团总部基地项目（二期、三期工程）施工总承包", "short": "明阳二三", "profit_center": "L400002314"},
    {"code": "1220155757", "full": "中山市工业技术研究中心项目施工", "short": "中山工业园", "profit_center": "L400002334"},
    {"code": "1220161245", "full": "明阳智慧能源集团总部基地项目（四期工程）", "short": "明阳四期", "profit_center": "L400002491"},
    {"code": "1220164507", "full": "鄱阳科技园项目总承包工程（C标段）", "short": "鄱阳C", "profit_center": "L400002579"},
    {"code": "1220168774", "full": "鼎湖科技园项目之总承包工程A", "short": "鼎湖A", "profit_center": "L400002667"},
    {"code": "1220176826", "full": "珠海市维琪科技有限公司新建维琪健康产业园项目", "short": "维琪", "profit_center": ""},
    {"code": "1210057186", "full": "东莞保利岭山林语花园项目", "short": "岭山林语", "profit_center": "L400000493"},
    {"code": "1210057196", "full": "惠州实地东部现代城花园三、四期", "short": "实地三四", "profit_center": "L400100176"},
    {"code": "1210057224", "full": "惠州实地东部现代城花园五期", "short": "实地五期", "profit_center": "L400100207"},
    {"code": "1210057243", "full": "珠海三一南方总部大厦", "short": "三一南方", "profit_center": "L400000495"},
    {"code": "1210057272", "full": "珠海保利茉莉花园", "short": "茉莉花园", "profit_center": "L400100190"},
    {"code": "1210057312", "full": "珠海华发绿洋湾花园主体建安工程（2标段）-项目", "short": "华发绿洋湾", "profit_center": "L400100175"},
    {"code": "1210057364", "full": "珠海灏怡财富中心主体工程", "short": "灏怡", "profit_center": "L400000463"},
    {"code": "1210057379", "full": "珠海恒大海泉湾花园（二标段）", "short": "恒大海泉湾", "profit_center": "L400100128"},
    {"code": "1210057399", "full": "珠海保利百合公馆", "short": "保利百合公馆", "profit_center": "L400100130"},
    {"code": "1210057456", "full": "江门里仁小学", "short": "里仁小学", "profit_center": "L400001275"},
    {"code": "1210057459", "full": "珠海华发依山郡花园一期工程（二标段）", "short": "华发依山郡", "profit_center": "L400100184"},
    {"code": "1210057467", "full": "江门保利中心一期", "short": "保利中心一期", "profit_center": "L400000462"},
    {"code": "1210057480", "full": "江门丰泰小学", "short": "丰泰小学", "profit_center": "L400001273"},
    {"code": "1210057489", "full": "江门保利××项目", "short": "××项目", "profit_center": "L400000507"},
    {"code": "1210057493", "full": "珠海灏怡财富中心", "short": "灏怡", "profit_center": "L400000465"},
    {"code": "1210057671", "full": "惠州实地东部现代城花园四季水乡", "short": "四季水乡", "profit_center": "L400100189"},
    {"code": "1210057693", "full": "珠海碧桂园澳邻帕克文化商业中心", "short": "澳邻帕克", "profit_center": "L400000500"},
    {"code": "1220010778", "full": "江门宝茵翠亭", "short": "宝茵翠亭", "profit_center": "L400100023"},
    {"code": "1220010866", "full": "东莞广东欧珀增资扩产厂房五、六、七、宿舍五、地下室", "short": "OPPO增资", "profit_center": "L400000492"},
    {"code": "1220016081", "full": "三一平沙新城（一期）施工总承包工程", "short": "三一平沙", "profit_center": "L400000491"},
    {"code": "1220016081", "full": "三一平沙新城（一期）施工总承包工程", "short": "三一蓝海", "profit_center": "L400000491"},
    {"code": "1220022866", "full": "东部现代城小学、地下室", "short": "实地小学", "profit_center": "L400100022"},
    {"code": "1220075511", "full": "融创云水观棠花园一期总承包工程", "short": "云水观棠", "profit_center": "L400000486"},
    {"code": "1220082086", "full": "金玥湾项目施工总承包工程", "short": "金玥湾", "profit_center": "L400001387"},
    {"code": "1220088733", "full": "台山市保利公馆三期土建暨水电安装工程", "short": "保利和公馆", "profit_center": "L400100228"},
    {"code": "1220088860", "full": "江门市保利天汇项目土建暨水电安装工程", "short": "保利天汇", "profit_center": "L400100230"},
    {"code": "1220089366", "full": "华为南方工厂项目2号生产厂房、21号空压站、冷冻站土建零星工程", "short": "华为二期", "profit_center": "L400001421"},
    {"code": "1220091358", "full": "中山翠亨新区生物医药智创中心PC采购施工总承包", "short": "中山翠亨", "profit_center": "L400001439"},
    {"code": "1220093270", "full": "惠环南10号小区项目（一期、二期）", "short": "惠环南", "profit_center": "L400001446"},
    {"code": "1220093271", "full": "滨海大都会花园", "short": "滨海大都会", "profit_center": "L400001445"},
    {"code": "1220095400", "full": "东莞市万江区保利天际花园项目一标段土建及水电安装工程项目", "short": "万江一标", "profit_center": "L400001463"},
    {"code": "1220096498", "full": "东莞南方工厂二期（E2D）项目", "short": "华为二期", "profit_center": "L400001469"},
    {"code": "1220101015", "full": "vivo深圳总部-基坑支护工程", "short": "vivo基坑", "profit_center": "L400001493"},
    {"code": "1220101616", "full": "江门宝茵翠亭住宅小区建设工程总承包项目", "short": "宝茵翠亭", "profit_center": "L400001499"},
    {"code": "1220104648", "full": "万科金色城央花园", "short": "金色城央", "profit_center": "L400001514"},
    {"code": "1220123205", "full": "汇华水岸花园施工", "short": "汇华水岸", "profit_center": "L400001631"},
    {"code": "1220129616", "full": "OPPO智能制造中心C地块-数据中心总承包工程项目", "short": "OPPO数据", "profit_center": "L400001674"},
    {"code": "1220134890", "full": "粤福电子加速技术应用项目1-4幢厂房，宿舍楼，垃圾房", "short": "粤福电子", "profit_center": "L400001749"},
    {"code": "1220138472", "full": "中基君豪股份有限公司总部大楼项目", "short": "中基君豪", "profit_center": "L400100260"},
    {"code": "1220140006", "full": "谢岗镇依城大观花园项目二期土建及水电工程", "short": "依城大观", "profit_center": "L400001917"},
    {"code": "1220149224", "full": "东莞市长安镇莞港跨境货物运输接驳点工程", "short": "接驳站", "profit_center": "L400002160"},
    {"code": "1220157210", "full": "广州（清远）健康驿站项目施工总承包（一）", "short": "清远驿站", "profit_center": "L400002385"},
    {"code": "1220159643", "full": "滨海湾宿舍楼淤泥固化工程", "short": "滨海湾宿舍", "profit_center": "L400002446"},
    {"code": "1220160260", "full": "OPPO滨海湾高级人才房桩基础", "short": "OPPO滨海湾", "profit_center": "L400002466"},
]

# 法院代码映射（参考 references/court-codes.md；案号格式：粤+地级市2位+基层法院2位）
# 键为法院全称，值为 4 位代码（不含"粤"前缀）
COURT_CODES = {
    "广州市天河区人民法院": "0106",
    "广州市花都区人民法院": "0114",
    "广州市番禺区人民法院": "0113",
    "广州市增城区人民法院": "0118",
    "广州市从化区人民法院": "0117",
    "广州市黄埔区人民法院": "0112",
    "广州市白云区人民法院": "0111",
    "深圳市福田区人民法院": "0304",
    "深圳市罗湖区人民法院": "0306",
    "深圳市南山区人民法院": "0305",
    "深圳市宝安区人民法院": "0306",
    "深圳市盐田区人民法院": "0308",
    "深圳市龙岗区人民法院": "0307",
    "深圳市龙华区人民法院": "0309",
    "深圳市坪山区人民法院": "0310",
    "深圳市光明区人民法院": "0311",
    "东莞市第一人民法院": "1971",
    "东莞市第二人民法院": "1972",
    "东莞市第三人民法院": "1973",
    "惠州市惠城区人民法院": "1302",
    "惠州市惠阳区人民法院": "1303",
    "江门市蓬江区人民法院": "0703",
    "江门市新会区人民法院": "0705",
    "江门市江海区人民法院": "0704",
    "江门市台山市人民法院": "0781",
    "中山市第一人民法院": "2071",
    "中山市第二人民法院": "2072",
    "珠海市香洲区人民法院": "0402",
    "珠海市金湾区人民法院": "0404",
    "珠海市斗门区人民法院": "0403",
    "佛山市禅城区人民法院": "0604",
    "佛山市南海区人民法院": "0605",
    "佛山市顺德区人民法院": "0606",
    "佛山市高明区人民法院": "0608",
    "佛山市三水区人民法院": "0607",
    "肇庆市端州区人民法院": "1202",
    "横琴粤澳深度合作区人民法院": "0491",
}

# 项目/地点 → 法院代码的常见推定（仅供参考，最终以法院受理通知书为准）
# 注：鄱阳项目存在花都(0114)与龙华(0309)两种情况，需结合具体案卷确认
PROJECT_COURT_HINTS = {
    "万江": "1971",        # 东莞一院
    "台山": "0781",        # 江门台山
    "保利和公馆": "0781",
    "鼎湖": "0114",        # 花都
    "鄱阳": "0114",        # 多数花都；少数龙华(0309)需确认
    "龙华": "0309",
}


def court_suffix(court_code: str, year: int = None) -> str:
    """
    生成案号尾缀 (YYYY)粤XXXX民初，不含具体案号数字（序列号/号）。
    当已知管辖法院、但法院尚未分配具体案号时使用，例如 (20XX)粤XXXX民初。
    """
    if not court_code:
        return ""
    y = year if year else datetime.now().year
    return f"（{y}）粤{court_code}民初"


def suggest_court_code(text: str = "", project: str = "", counterparty: str = "") -> str:
    """
    根据项目/地点关键词、相对方、正文尽力推定管辖法院代码（4位）。
    返回空串表示无法推定，需人工确认。仅为辅助提示，最终以法院受理通知书案号为准。
    """
    hay = f"{project} {counterparty} {text}"
    for kw, code in PROJECT_COURT_HINTS.items():
        if kw in hay:
            return code
    for court_name, code in COURT_CODES.items():
        if court_name in text:
            return code
    return ""


# 多项目统称规则
MULTI_PROJECT_RULES = {
    frozenset(["明阳二三", "明阳四期"]): "明阳能源",
    frozenset(["万江一标", "万江二标"]): "万江保利",
    frozenset(["里仁中学", "里仁小学"]): "里仁中小学",
    frozenset(["紫茶中学", "紫茶小学"]): "紫茶中小学",
    frozenset(["实地三四", "实地五期", "四季水乡", "实地小学"]): "实地三四",
    frozenset(["三一平沙", "三一蓝海"]): "三一平沙",
    frozenset(["海三", "星光海岸"]): "星光海岸"
}

# ==========================================
# 2. 文本提取与处理模块 (Text Extractor)
# ==========================================

def get_ocr_text(file_path: Path) -> str:
    """
    获取文件OCR文本。优先级：
    1. 同名.txt文件
    2. .txt文件直接读取
    3. PDF有文字层 → pypdf提取
    4. PDF纯扫描件 → fitz渲染+Tesseract自动OCR
    5. 以上均失败 → 手动粘贴
    """
    txt_path = file_path.with_suffix('.txt')
    if txt_path.exists():
        return txt_path.read_text(encoding='utf-8')
    if file_path.suffix == '.txt':
        return file_path.read_text(encoding='utf-8')

    if file_path.suffix.lower() == '.pdf':
        # 尝试pypdf提取文字层
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(file_path))
            texts = []
            for i, page in enumerate(reader.pages):
                t = page.extract_text()
                if t and t.strip():
                    texts.append(f"===== 第{i+1}页 =====\n{t}")
            if texts:
                return "\n\n".join(texts)
        except Exception:
            pass
        # 尝试fitz提取文字层
        try:
            import fitz
            doc = fitz.open(str(file_path))
            texts = []
            for i in range(doc.page_count):
                t = doc[i].get_text().strip()
                if t:
                    texts.append(f"===== 第{i+1}页 =====\n{t}")
            doc.close()
            if texts:
                return "\n\n".join(texts)
        except Exception:
            pass
        # 自动OCR降级
        print(f"[OCR] 纯扫描件，正在自动OCR识别（{file_path.name}）...")
        try:
            return _ocr_pdf_pages(file_path)
        except Exception as e:
            print(f"[OCR] 自动OCR失败: {e}")

    return input(f"[OCR占位] 请粘贴 {file_path.name} 的文本内容: ")

def _ocr_pdf_pages(pdf_path: Path, pages: list = None) -> str:
    """
    用fitz渲染 + Tesseract OCR识别PDF页面（ocr.py同款方案）。
    pages: 要识别的页码列表（1-based），None则全页。
    返回合并文本，格式: "===== 第N页 =====\n文本"
    """
    import fitz
    doc = fitz.open(str(pdf_path))
    total = doc.page_count
    if pages is None:
        pages = list(range(1, total + 1))

    mat = fitz.Matrix(300/72, 300/72)
    all_text = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for pg in pages:
            idx = pg - 1
            if idx < 0 or idx >= total:
                continue
            pix = doc[idx].get_pixmap(matrix=mat)
            img_path = os.path.join(tmpdir, f"ocr_p{pg:04d}.png")
            pix.save(img_path)
            outbase = os.path.join(tmpdir, f"ocr_p{pg:04d}")
            result = subprocess.run(
                [TESSERACT_PATH, img_path, outbase, "-l", OCR_LANG],
                capture_output=True, text=True, timeout=120
            )
            txt_path = outbase + ".txt"
            if os.path.exists(txt_path):
                with open(txt_path, "r", encoding="utf-8") as f:
                    all_text.append(f"===== 第{pg}页 =====\n{f.read()}")
    doc.close()
    return "\n\n".join(all_text)


def deduplicate_pages(pdf_path: Path, threshold: int = 8) -> list:
    """
    用感知哈希(pHash)检测PDF中视觉重复的页面。
    参数:
        pdf_path: PDF文件路径
        threshold: Hamming距离阈值，小于此值视为重复（默认8）
    返回:
        [(保留页, [重复页...]), ...]  每组第一个为保留页，其余为可剔除的重复页
        空列表表示无重复
    """
    try:
        import fitz
        from PIL import Image
        import imagehash
    except ImportError as e:
        print(f"[去重] 缺少依赖: {e}，跳过去重")
        return []

    doc = fitz.open(str(pdf_path))
    total = doc.page_count
    if total <= 1:
        doc.close()
        return []

    print(f"[去重] 正在检测 {total} 页的重复情况...")
    mat = fitz.Matrix(150/72, 150/72)  # 150 DPI足够做哈希
    hashes = {}
    for i in range(total):
        pix = doc[i].get_pixmap(matrix=mat)
        img = Image.open(BytesIO(pix.tobytes("png")))
        hashes[i + 1] = imagehash.phash(img)
    doc.close()

    # 查找重复组
    duplicates = []
    matched = set()
    for i in range(1, total + 1):
        if i in matched:
            continue
        group = [i]
        for j in range(i + 1, total + 1):
            if j in matched:
                continue
            if (hashes[i] - hashes[j]) < threshold:
                group.append(j)
                matched.add(j)
        if len(group) > 1:
            duplicates.append((group[0], group[1:]))
            matched.add(i)

    if duplicates:
        print(f"[去重] 发现 {len(duplicates)} 组重复页面：")
        for keep, dups in duplicates:
            print(f"        保留第{keep}页，剔除第{','.join(map(str, dups))}页")
    else:
        print("[去重] 未发现重复页面")
    return duplicates


def clean_and_format_case_number(case_num: str) -> str:
    if not case_num: return ""
    clean = case_num.replace(" ", "").replace("\u3000", "")
    return clean.replace("(", "（").replace(")", "）")

def extract_all_case_numbers(text: str) -> list:
    raw_matches = re.findall(r'[（(]\s*\d{4}\s*[）)][\s\S]{0,30}?号', text)
    if not raw_matches:
        raw_matches = re.findall(r'\d{4}[\s\S]{0,30}?号', text)
    cleaned = set()
    for m in raw_matches:
        cleaned.add(clean_and_format_case_number(m))
    return list(cleaned)

def extract_first_instance_case_number(text: str) -> str:
    match = re.search(r'[（(]\s*\d{4}\s*[）)][\s\S]{0,20}?民初[\s\S]{0,10}?号', text)
    if match: return clean_and_format_case_number(match.group(0))
    return ""

def simplify_counterparty(name: str) -> str:
    if not name: return ""
    suffixes = ["（深圳）", "(深圳)", "（广州）", "(广州)", "（东莞）", "(东莞)",
                "有限公司", "有限责任公司", "股份有限公司", "集团"]
    clean_name = name
    for s in suffixes:
        clean_name = clean_name.replace(s, "")
    return clean_name.strip()

def extract_counterparty(text: str) -> str:
    patterns = [r'原告[：:]\s*([^\n,，]+)', r'上诉人[：:]\s*([^\n,，]+)', r'申请人[：:]\s*([^\n,，]+)',
                r'异议人[：:]\s*([^\n,，]+)', r'被申请人[：:]\s*([^\n,，]+)',
                r'被执行人[：:]\s*([^\n,，]+)', r'申请执行人[：:]\s*([^\n,，]+)']
    for p in patterns:
        matches = re.findall(p, text)
        for m in matches:
            m = m.strip()
            if not any(ex in m for ex in EXCLUDE_PARTIES) and len(m) > 1:
                return simplify_counterparty(m)
    return ""

def extract_hearing_date(text: str) -> str:
    match = re.search(r'(\d{1,2})月(\d{1,2})日[\s\S]{0,10}?开庭', text)
    if match: return f"{int(match.group(1)):02d}.{int(match.group(2)):02d}"
    return ""

def extract_project_short_name(text: str) -> str:
    found_shorts = set()
    codes = re.findall(r'\b(12[12]\d{8})\b', text)
    for code in codes:
        for p in PROJECT_MAPPING:
            if p['code'] == code:
                found_shorts.add(p['short'])
    for p in PROJECT_MAPPING:
        if p['full'] in text or p['short'] in text:
            found_shorts.add(p['short'])
    if "灏怡" in found_shorts:
        found_shorts = {"灏怡"}
    if len(found_shorts) > 1:
        for rule_set, unified_name in MULTI_PROJECT_RULES.items():
            if rule_set.issubset(found_shorts):
                return unified_name
        jiangmen_schools = {"里仁中小学", "紫茶中小学", "里仁中学", "里仁小学", "紫茶中学", "紫茶小学", "丰泰小学", "双朗小学"}
        if len(found_shorts.intersection(jiangmen_schools)) >= 2:
            return "江门学校"
    if len(found_shorts) == 1:
        return found_shorts.pop()
    elif len(found_shorts) > 1:
        print(f"[提示] 检测到多个项目: {', '.join(found_shorts)}，请手动确认统称。")
        return ""
    return ""


# ==========================================
# 3. ★★★ 搜索匹配引擎 (完全重写) ★★★
# ==========================================

def _clean_for_search(s: str) -> str:
    """清洗字符串用于搜索比对：去掉所有括号、空格、OCR噪声等特殊符号"""
    result = s.replace("（", "").replace("）", "").replace("(", "").replace(")", "") \
             .replace(" ", "").replace("\u3000", "").replace("【", "").replace("】", "") \
             .replace("《", "").replace("》", "").replace("〈", "").replace("〉", "")
    # 剔除Tesseract OCR对中文的常见误读字符，避免案号匹配失败
    # 例如: OCR把"粤"误读为"#" → "20XX#XXXX执XXXXX号" vs 文件夹"20XX粤XXXX执XXXXX号"
    OCR_NOISE = r"#$%^`~@&*+=|:;'"  # raw string literal, no escape warnings
    for ch in OCR_NOISE:
        result = result.replace(ch, "")
    return result

def _build_search_keywords(case_nums: list, counterparty: str, project: str, text: str = "") -> list:
    """
    构建多维度搜索关键词列表。
    返回格式: [(keyword_clean, keyword_label), ...]
    
    维度：
    1. 完整案号（清洗后）
    2. 案号核心片段（年份+编号，去掉法院代字）
    3. 相对方全称和简称
    4. 项目名
    5. 源文件OCR中的额外案号
    """
    keywords = []

    # --- 维度1：完整案号 ---
    for cn in case_nums:
        if cn:
            cleaned = _clean_for_search(cn)
            if cleaned:
                keywords.append((cleaned, f"案号:{cn}"))

    # --- 维度2：案号核心片段（只取数字+编号部分） ---
    # 例如 "（20XX）Xxxx执XXXX号" → "2025" + "14119"
    for cn in case_nums:
        if cn:
            # 提取年份
            year_match = re.search(r'(\d{4})', cn)
            if year_match:
                year = year_match.group(1)
                # 提取编号（最后一个数字串，在"号"前面）
                num_match = re.findall(r'(\d{3,})', cn)
                if num_match:
                    core_id = num_match[-1]  # 取最后一个数字段作为核心编号
                    if len(core_id) >= 3:
                        keywords.append((core_id, f"编号:{core_id}"))

    # --- 维度3：相对方 ---
    if counterparty:
        cleaned_cp = _clean_for_search(counterparty)
        if cleaned_cp:
            keywords.append((cleaned_cp, f"相对方:{counterparty}"))
            # 如果相对方较长，也加入前4个字作为模糊搜索
            if len(cleaned_cp) > 4:
                keywords.append((cleaned_cp[:4], f"相对方简称:{cleaned_cp[:4]}"))

    # --- 维度4：项目名 ---
    if project and project != "未知项目":
        cleaned_proj = _clean_for_search(project)
        if cleaned_proj:
            keywords.append((cleaned_proj, f"项目:{project}"))

    # --- 维度5：从OCR文本中额外提取案号（补充维度1） ---
    if text:
        extra_nums = extract_all_case_numbers(text)
        for en in extra_nums:
            cleaned = _clean_for_search(en)
            if cleaned and cleaned not in [k[0] for k in keywords]:
                keywords.append((cleaned, f"OCR案号:{en}"))

    # 去重
    seen = set()
    unique = []
    for kw, label in keywords:
        if kw not in seen and len(kw) >= 2:
            seen.add(kw)
            unique.append((kw, label))

    return unique


def _score_match(item_name_clean: str, keywords: list) -> int:
    """
    给一个文件夹/文件名打分，看它和搜索关键词的匹配程度。
    返回匹配分数（越高越好），0表示完全不匹配。
    
    打分规则：
    - 案号完全匹配：+100分
    - 案号编号片段匹配：+50分
    - 相对方完全匹配：+80分
    - 相对方部分匹配（前4字）：+30分
    - 项目名匹配：+40分
    """
    score = 0
    for kw, label in keywords:
        if kw not in item_name_clean:
            continue
        if "案号" in label and "OCR" not in label:
            score += 100
        elif "编号" in label:
            score += 50
        elif "相对方" in label and "简称" not in label:
            score += 80
        elif "相对方简称" in label:
            score += 30
        elif "项目" in label:
            score += 40
        elif "OCR案号" in label:
            score += 60
    return score


def search_case_folders(root_dirs: list, keywords: list, dirs_only: bool = False) -> list:
    """
    ★ 核心搜索函数 ★
    在指定目录下递归搜索所有文件夹（和文件），用多维度关键词打分匹配。
    
    返回: [(score, Path), ...] 按分数从高到低排序
    """
    if not keywords:
        return []

    results = []
    seen_paths = set()

    for root_dir in root_dirs:
        root_path = Path(root_dir)
        if not root_path.exists():
            print(f"[警告] 搜索目录不存在: {root_path}")
            continue

        try:
            for item in root_path.rglob("*"):
                # 跳过自己（正在处理的文件）
                real_item = item.resolve()
                if str(real_item) in seen_paths:
                    continue
                seen_paths.add(str(real_item))

                if not item.exists():
                    continue
                if dirs_only and not item.is_dir():
                    continue

                # 清洗文件名用于比对
                clean_name = _clean_for_search(item.name)

                # 打分
                score = _score_match(clean_name, keywords)

                if score > 0:
                    results.append((score, item))
        except PermissionError:
            print(f"[警告] 无权限访问: {root_path}")
        except Exception as e:
            print(f"[警告] 搜索 {root_path} 时出错: {e}")

    # 按分数从高到低排序
    results.sort(key=lambda x: x[0], reverse=True)
    return results


def extract_info_from_folder_name(folder_name: str) -> dict:
    """从已有的文件夹名中反向提取相对方、项目、案号"""
    info = {"counterparty": "", "project": "", "case_num": ""}

    # 提取相对方：【】内的内容
    cp_match = re.search(r'【(.+?)】', folder_name)
    if cp_match:
        info["counterparty"] = cp_match.group(1)

    # 提取案号
    case_match = re.search(r'[（(]\d{4}[）)][\s\S]*?号', folder_name)
    if case_match:
        info["case_num"] = case_match.group(0)

    # 提取项目：】和案号之间的内容
    if info["counterparty"] and info["case_num"]:
        bracket_end = folder_name.find("】") + 1
        case_start = folder_name.find(info["case_num"])
        if 0 < bracket_end < case_start:
            info["project"] = folder_name[bracket_end:case_start]
    elif info["counterparty"]:
        # 没有案号时，取】后面的所有内容作为项目（可能包含其他信息）
        bracket_end = folder_name.find("】") + 1
        if 0 < bracket_end < len(folder_name):
            remaining = folder_name[bracket_end:]
            # 去掉日期前缀（如"恢复开庭"保留，"0806"去掉）
            remaining = re.sub(r'^\d{4}', '', remaining)
            if remaining:
                info["project"] = remaining

    return info


# ==========================================
# 4. 命名引擎与归档路由 (Naming & Filing)
# ==========================================

def generate_new_filename(prefix: str, counterparty: str, project: str, case_num: str, is_second: bool, text: str, court_code: str = None) -> str:
    final_case_num = case_num
    if is_second:
        first_num = extract_first_instance_case_number(text)
        if first_num:
            final_case_num = first_num

    proj_str = project if project else "未知项目"
    new_name = f"{prefix}【{counterparty}】{proj_str}{final_case_num}"
    # 已知管辖法院、但法院尚未分配具体案号时，在末尾补充 (年份)粤XXXX民初 尾缀
    if court_code:
        new_name += court_suffix(court_code)
    return re.sub(r'[\\/*?:"<>|]', "", new_name)


def interactive_search_and_backfill(case_nums, counterparty, project, text, dirs_only=False):
    """
    ★ 统一的交互式搜索+回填流程 ★
    搜索 → 展示 → 让用户选 → 回填信息
    返回: (case_num, counterparty, project, selected_dir_or_None)
    """
    search_dirs = [D_DEFENDANT, D_PRESERVATION, D_EXECUTION, D_PLAINTIFF, D_JUDGMENT]

    # 构建多维度搜索关键词
    keywords = _build_search_keywords(case_nums, counterparty, project, text)

    if not keywords:
        print("[搜索] 没有可用的搜索关键词，跳过自动匹配。")
        return case_nums[0] if case_nums else "", counterparty, project, None

    print(f"\n[搜索关键词] {', '.join([f'{label}' for _, label in keywords])}")

    # 执行搜索
    results = search_case_folders(search_dirs, keywords, dirs_only=dirs_only)

    if not results:
        # ★★ 扩大搜索范围到整个 D_ROOT ★★
        print("[搜索] 诉讼目录下未找到，扩大到整个D盘根目录搜索...")
        results = search_case_folders([D_ROOT], keywords, dirs_only=dirs_only)

    if not results:
        print("[搜索] 未找到任何匹配的文件夹/文件。")
        return case_nums[0] if case_nums else "", counterparty, project, None

    # 展示结果
    print(f"\n[自动匹配] 找到 {len(results)} 个相关结果：")
    for i, (score, p) in enumerate(results):
        ptype = "[夹]" if p.is_dir() else "[文]"
        score_tag = "★★★" if score >= 100 else "★★" if score >= 50 else "★"
        try:
            rel = p.relative_to(D_ROOT)
        except ValueError:
            rel = p
        print(f"  {i+1}. {ptype} {score_tag} {rel}")

    print(f"  {len(results)+1}. 放弃匹配，不回填")

    choice = input(f"选择序号回填信息 (1-{len(results)+1}，回车跳过): ").strip()
    if choice.isdigit() and 1 <= int(choice) <= len(results):
        score, selected = results[int(choice)-1]

        # 如果选中的是文件，用其父文件夹名来提取信息
        info_name = selected.parent.name if selected.is_file() else selected.name
        folder_info = extract_info_from_folder_name(info_name)

        old_cn, old_cp, old_proj = (case_nums[0] if case_nums else ""), counterparty, project

        if folder_info["case_num"]:
            case_nums = [folder_info["case_num"]]
        if folder_info["project"]:
            project = folder_info["project"]
        if folder_info["counterparty"]:
            counterparty = folder_info["counterparty"]

        new_cn = case_nums[0] if case_nums else ""
        print(f"\n[回填完成]")
        if old_cn != new_cn: print(f"  案号: {old_cn or '(空)'} → {new_cn}")
        if old_cp != counterparty: print(f"  相对方: {old_cp or '(空)'} → {counterparty}")
        if old_proj != project: print(f"  项目: {old_proj or '(空)'} → {project}")

        # 返回选中的目录（仅当它是文件夹时）
        selected_dir = selected if selected.is_dir() else None
        return new_cn, counterparty, project, selected_dir

    return case_nums[0] if case_nums else "", counterparty, project, None


def route_and_move(file_path: Path, new_filename: str, text: str, prefix: str,
                   counterparty: str, case_num: str, project: str, is_second: bool,
                   pre_selected_dir=None, amount_wan=None, branch=None):
    """
    路由并移动文件。
    pre_selected_dir: 如果 process_file 阶段已经选了文件夹，直接用，不再搜索。
    amount_wan: (可选) 案件法定标的额（万元，str/float），用于 Z 盘待指派目录命名尾缀 `-{amount}万元`。
    branch: (可选) Z 盘目录中的分公司简称（默认 Z_BRANCH="中山"，其他如 "广分"/"总包" 时传入）。
    """
    is_judgment = prefix in ["一审判决", "二审判决", "调解书", "裁定"]
    target_dirs = []

    # 当已有 pre_selected_dir 时，跳过默认特殊路径，让 pre_selected_dir 唯一决定归档位置
    if pre_selected_dir:
        target_dirs.append(pre_selected_dir)
        print(f"\n[路由] 使用已选文件夹: {pre_selected_dir.relative_to(D_ROOT)}")
    else:
        # 默认特殊路径映射
        if "保全" in prefix or "解封" in prefix:
            if prefix == "解封申请":
                target_dirs.append(D_PRESERVATION / "解封申请")
            elif "保全" in prefix:
                target_dirs.append(D_PRESERVATION)

        if "执行" in prefix and "执行异议" not in prefix:
            target_dirs.append(D_EXECUTION)

        # ★ 重新搜索（只在诉讼子目录下搜文件夹） ★
        all_case_nums = extract_all_case_numbers(text)
        if case_num and case_num not in all_case_nums:
            all_case_nums.append(case_num)

        keywords = _build_search_keywords(all_case_nums, counterparty, project, text)
        search_dirs = [D_DEFENDANT, D_PRESERVATION, D_EXECUTION, D_PLAINTIFF]
        results = search_case_folders(search_dirs, keywords, dirs_only=True)

        if results:
            print(f"\n[路由搜索] 找到 {len(results)} 个近似案件文件夹：")
            for i, (score, p) in enumerate(results):
                score_tag = "★★★" if score >= 100 else "★★" if score >= 50 else "★"
                try:
                    print(f"  {i+1}. {score_tag} {p.relative_to(D_ROOT)}")
                except ValueError:
                    print(f"  {i+1}. {score_tag} {p}")
            print(f"  {len(results)+1}. 新建文件夹")

            choice = input(f"请选择要放入的文件夹序号 (1-{len(results)+1}，回车新建): ")
            if choice.isdigit() and 1 <= int(choice) <= len(results):
                _, selected_dir = results[int(choice)-1]

                # 回填信息并重新生成文件名
                folder_info = extract_info_from_folder_name(selected_dir.name)
                if not case_num and folder_info["case_num"]:
                    case_num = folder_info["case_num"]
                if not project and folder_info["project"]:
                    project = folder_info["project"]

                new_filename = generate_new_filename(prefix, counterparty, project, case_num, is_second, text)
                print(f"[更新命名] 回填后: {new_filename}{file_path.suffix}")

                if is_judgment and D_DEFENDANT not in selected_dir.parents and selected_dir != D_DEFENDANT:
                    defendant_results = search_case_folders([D_DEFENDANT],
                        _build_search_keywords([selected_dir.name], "", "", ""), dirs_only=True)
                    if defendant_results:
                        target_dirs.append(defendant_results[0][1])
                    else:
                        target_dirs.append(selected_dir)
                else:
                    target_dirs.append(selected_dir)
            else:
                # 新建
                new_dir_name = f"{datetime.now().strftime('%m%d')}【{counterparty}】{case_num}"
                new_dir = D_DEFENDANT / new_dir_name
                new_dir.mkdir(parents=True, exist_ok=True)
                target_dirs.append(new_dir)
        else:
            # 完全找不到
            new_dir_name = f"{datetime.now().strftime('%m%d')}【{counterparty}】{case_num}"
            new_dir = D_DEFENDANT / new_dir_name
            new_dir.mkdir(parents=True, exist_ok=True)
            target_dirs.append(new_dir)

    # 3. 处理判决/裁定的双路径
    if is_judgment:
        today = datetime.now()
        ym_dir = D_JUDGMENT / f"【判决裁定】{today.year}年{today.month}月收"
        ym_dir.mkdir(parents=True, exist_ok=True)
        target_dirs.append(ym_dir)

    # 4. 执行移动与复制
    print("\n--- 行动报告 ---")
    ext = file_path.suffix
    for i, target_dir in enumerate(target_dirs):
        final_name = new_filename
        if (target_dir == D_JUDGMENT or (D_JUDGMENT in target_dir.parents)) and prefix == "一审判决":
            final_name = f"({datetime.now().strftime('%Y%m%d')}收){new_filename}"

        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = target_dir / f"{final_name}{ext}"

        if target_path.exists() and input(f"[警告] {target_path.name} 已存在，是否覆盖？(y/n): ").lower() != 'y':
            continue

        if i == 0:
            shutil.move(str(file_path), str(target_path))
            print(f"[移动] {file_path.name} -> {target_path.relative_to(D_ROOT)}")
            file_path = target_path
        else:
            shutil.copy2(str(file_path), str(target_path))
            print(f"[复制] {file_path.name} -> {target_path.relative_to(D_ROOT)}")

    # 5. Z盘指派逻辑（新收起诉材料 → 02待指派律师的案件）
    # 目录命名: {MM.DD}开庭-{分公司简写}-{相对方简称}-{案号}-{标的额}万元
    if prefix == "起诉材料" and input("\n是否需要委托律师并归档到 Z 盘指派目录？(y/n): ").lower() == 'y':
        hearing_date = extract_hearing_date(text) or input("请输入开庭日期 (MM.DD，如08.06，回车填'待排期'): ") or "待排期"
        branch_s = branch or Z_BRANCH
        amount_s = f"-{amount_wan}万元" if amount_wan else ""
        z_dir = Z_ROOT / f"{hearing_date}开庭-{branch_s}-{counterparty}-{case_num}{amount_s}"
        z_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(file_path), str(z_dir / f"{new_filename}{ext}"))
        print(f"[指派] 已复制起诉状到 Z 盘: {z_dir.relative_to(Z_ROOT)}")


# ==========================================
# 5. 主交互流程 (Main CLI)
# ==========================================

def process_file(file_path_str: str):
    file_path = Path(file_path_str)
    if not file_path.exists():
        return print(f"[错误] 文件不存在: {file_path}")

    print(f"\n{'='*60}")
    print(f"=== 开始处理: {file_path.name} ===")
    print(f"{'='*60}")

    text = get_ocr_text(file_path)

    # ---- 第一步：OCR提取 ----
    all_case_nums = extract_all_case_numbers(text)
    case_num = all_case_nums[0] if all_case_nums else ""
    counterparty = extract_counterparty(text)
    project = extract_project_short_name(text)

    print(f"\n[OCR提取]")
    print(f"  案号: {case_num or '(未找到)'} (共发现{len(all_case_nums)}个)")
    print(f"  相对方: {counterparty or '(未找到)'}")
    print(f"  项目: {project or '(未找到)'}")

    # ---- 第二步：确认/修正提取结果 ----
    if not counterparty:
        counterparty = input("\n请手动输入相对方: ")
    elif input(f"\n相对方确认为 '{counterparty}' 吗？(y/n): ").lower() == 'n':
        counterparty = input("请手动输入相对方: ")

    if not project:
        project = input("请输入项目缩写 (若无关联项目直接回车，将使用'未知项目'): ")

    # ---- 第三步：判断材料类型 ----
    is_second = "二审" in text or "终" in case_num
    prefix = next((v for k, v in MATERIAL_PREFIXES.items() if k in text + " " + file_path.name), "其他材料")
    # 判决书需区分一审/二审：命中"判决书"但属二审时，纠正为"二审判决"
    if prefix == "一审判决" and is_second:
        prefix = "二审判决"
    print(f"  材料类型: {prefix}")
    print(f"  是否二审: {'是' if is_second else '否'}")

    # ---- 第四步：★ 多维度搜索 + 回填 ★ ----
    case_num, counterparty, project, selected_dir = interactive_search_and_backfill(
        case_nums=all_case_nums,
        counterparty=counterparty,
        project=project,
        text=text,
        dirs_only=False
    )

    # ---- 第五步：生成最终文件名 ----
    new_filename = generate_new_filename(prefix, counterparty, project, case_num, is_second, text)
    print(f"\n[最终命名] {new_filename}{file_path.suffix}")

    if input("确认使用此文件名吗？(y/n): ").lower() == 'n':
        new_filename = input("请手动输入新文件名 (不含后缀): ")

    # ---- 第六步：路由并移动（传入选好的文件夹） ----
    route_and_move(file_path, new_filename, text, prefix, counterparty, case_num, project, is_second,
                   pre_selected_dir=selected_dir)
    print(f"\n=== 处理完成 ===\n")


if __name__ == "__main__":
    print("=" * 60)
    print("  案件文件命名与归档工具 v2.0")
    print("  搜索范围：D:/【用户名】 全目录")
    print("=" * 60)
    while True:
        target = input("\n请输入要处理的文件路径 (或输入 q 退出): ").strip()
        if target.lower() == 'q':
            break
        if target:
            try:
                process_file(target)
            except Exception as e:
                import traceback
                print(f"[严重错误] {e}")
                traceback.print_exc()