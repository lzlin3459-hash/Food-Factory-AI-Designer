import streamlit as st
import ezdxf
from ezdxf.enums import TextEntityAlignment
import pandas as pd
import plotly.graph_objects as go
from openai import OpenAI
import io
import random
import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn

# --- 0. 图表全局防乱码配置 ---
font_path_cloud = "AlibabaPuHuiTi-3-115-Black.ttf"
if os.path.exists(font_path_cloud):
    prop = fm.FontProperties(fname=font_path_cloud)
else:
    prop = fm.FontProperties(family='SimHei')
plt.rcParams['font.family'] = prop.get_name()
plt.rcParams['axes.unicode_minus'] = False

# --- 1. 配置与 AI 初始化 ---
try: api_key = st.secrets["DEEPSEEK_API_KEY"]
except: api_key = "请配置您的 API Key"
client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

def fetch_market_data(): return round(16500 + random.uniform(-150, 400), 2), round(95 + random.uniform(-5, 15), 2)
def calculate_physics(cap, cat):
    f = {"饼干/烘焙": 1.0, "乳制品/液态奶": 1.5, "肉制品/切片": 1.3}.get(cat, 1.0)
    return {"elec": round(cap * 15 * f, 1), "water": round(cap * 3.5 * f, 1), "steam": round(cap * 0.8 * f, 1)}

# --- 2. 满血版 BOM (细化食堂、淋浴、办公等独立建筑) ---
def generate_detailed_bom(category, capacity, physics):
    base_bom = [
        {"专业": "给排水", "设备与材料": "RO反渗透与超滤水站", "规格参数": f"产水 {physics['water']}t/d", "数量": 1, "预估造价": 35.0, "合规说明": "设备接触水须达直饮标准", "智能运维建议": "每日自动检测水质硬度与RO膜压差"},
        {"专业": "给排水", "设备与材料": "SUS304防臭排水地沟", "规格参数": "宽20cm, 坡度1.5%", "数量": 1, "预估造价": 15.0, "合规说明": "水封地漏+6mm防鼠金属格栅", "智能运维建议": "沉渣槽配备超声波淤堵报警器"},
        {"专业": "电气", "设备与材料": "380V主配电与桥架管网", "规格参数": f"总装机 {physics['elec']}kW", "数量": 1, "预估造价": 28.0, "合规说明": "动力与照明强制分离", "智能运维建议": "接入能耗管理平台，监控峰谷电与功率因数"},
        {"专业": "仓储物流", "设备与材料": "原辅料立体高架库", "规格参数": "重型货架", "数量": 1, "预估造价": 45.0, "合规说明": "防鼠防虫防潮三级隔离", "智能运维建议": "温湿度传感器联动空调自动化"},
        {"专业": "仓储物流", "设备与材料": "成品发货与冷链穿梭库", "规格参数": "WMS系统", "数量": 1, "预估造价": 55.0, "合规说明": "月台需设防虫软帘", "智能运维建议": "叉车防撞预警与电量监控"},
        # 核心拆分：行政与生活基建
        {"专业": "建筑基建", "设备与材料": "综合行政办公与研发大楼", "规格参数": "独立建筑", "数量": 1, "预估造价": 60.0, "合规说明": "远离散发异味与粉尘区域", "智能运维建议": "中央空调智能分区控制"},
        {"专业": "建筑基建", "设备与材料": "独立员工食堂与后厨", "规格参数": "含油烟净化", "数量": 1, "预估造价": 30.0, "合规说明": "食堂排烟严禁串入生产区", "智能运维建议": "燃气泄漏联动电磁阀自动切断"},
        {"专业": "建筑基建", "设备与材料": "倒班宿舍楼", "规格参数": "4人间配置", "数量": 1, "预估造价": 50.0, "合规说明": "生活污水需独立化粪池处理", "智能运维建议": "宿舍区水电独立智能计费"},
        # 核心拆分：卫生防线
        {"专业": "卫生基建", "设备与材料": "外围独立卫生间", "规格参数": "非手接触水龙头", "数量": 1, "预估造价": 10.0, "合规说明": "卫生间门不得直对生产车间", "智能运维建议": "排风扇与照明红外感应联动"},
        {"专业": "卫生基建", "设备与材料": "一更/二更与独立淋浴间", "规格参数": "干湿分离", "数量": 1, "预估造价": 15.0, "合规说明": "高风险区员工必须强制沐浴更衣", "智能运维建议": "更衣柜臭氧定时消杀系统"},
        {"专业": "卫生基建", "设备与材料": "风淋室与强制洗手池", "规格参数": "互锁门禁", "数量": 1, "预估造价": 12.0, "合规说明": "未消毒无法推开进入车间的门", "智能运维建议": "HEPA滤网压差阻力在线监测报警"}
    ]
    if category == "饼干/烘焙":
        spec = [
            {"专业": "工艺", "设备与材料": "燃气隧道炉及预混成型线", "规格参数": f"{capacity}T/天", "数量": 1, "预估造价": 185.0, "合规说明": "排烟罩伸出40cm，废气余热回收", "智能运维建议": "配置红外扫描仪监测炉温曲线"},
            {"专业": "暖通", "设备与材料": "压差控制与排气系统", "规格参数": "热区排烟风速15m/s", "数量": 1, "预估造价": 18.0, "合规说明": "热加工区-5Pa负压，包装区+5Pa正压", "智能运维建议": "压差传感器实时在线监测报警"}
        ]
    elif category == "乳制品/液态奶":
        spec = [
            {"专业": "工艺", "设备与材料": "UHT杀菌与无菌灌装线", "规格参数": "百级净化", "数量": 1, "预估造价": 280.0, "合规说明": "蒸汽设2%下坡排冷凝水及疏水阀", "智能运维建议": "严格监控137℃保温时间，异常自动回流"},
            {"专业": "专项", "设备与材料": "全自动七步CIP清洗站", "规格参数": "管材SUS316L, 介质85℃", "数量": 1, "预估造价": 45.0, "合规说明": "清洗液温度达标，消除卫生死角", "智能运维建议": "电导率仪实时在线标定清洗液浓度"}
        ]
    else: 
        spec = [
            {"专业": "工艺", "设备与材料": "真空斩拌与深冷速冻库", "规格参数": "-35℃螺杆制冷", "数量": 1, "预估造价": 220.0, "合规说明": "动物/植物/水产清洗水池严格分开", "智能运维建议": "化霜周期根据霜层厚度智能调节"},
            {"专业": "给排水", "设备与材料": "热加工区含油废水处理站", "规格参数": "清洗水温75℃", "数量": 1, "预估造价": 25.0, "合规说明": "热加工区后置必须串联油水分离器", "智能运维建议": "隔油池油位超限自动提醒清掏"}
        ]
    return pd.DataFrame(base_bom + spec)

# --- 3. 终极一统 CAD 引擎：极度细化的总图规划 ---
def generate_ultimate_dxf(capacity, category):
    doc = ezdxf.new('R2010'); doc.styles.new('CHS', dxfattribs={'font': 'simsun.ttc'})
    doc.layers.add("BUILDING_LIFE", color=7)   # 生活办公区 (白)
    doc.layers.add("BUILDING_LOG", color=6)    # 仓储物流区 (紫)
    doc.layers.add("TECH_EQUIP", color=3)      # 工艺生产区 (绿)
    doc.layers.add("PLUMBING_DRAIN", color=5, linetype="DASHED") # 给排水 (蓝)
    doc.layers.add("ELEC_POWER", color=2, linetype="DASHED")     # 电气 (黄)
    doc.layers.add("HVAC_AIR", color=1, linetype="DASHED")       # 暖通 (红)
    
    msp = doc.modelspace()
    def draw_box(name, x, length, y, h, layer, color, text_size=150):
        msp.add_lwpolyline([(x, y), (x+length, y), (x+length, y+h), (x, y+h), (x, y)], dxfattribs={'layer': layer, 'color': color})
        msp.add_text(name, dxfattribs={'height': text_size, 'style': 'CHS', 'color': color}).set_placement((x+length/2, y+h/2), align=TextEntityAlignment.MIDDLE_CENTER)

    # 1. 核心工艺生产厂房 (TECH_EQUIP)
    py = 1000; ph = 3500; end_x = 15000
    if category == "饼干/烘焙":
        draw_box("原料预混打面车间", 0, 3500, py, ph, "TECH_EQUIP", 3)
        draw_box(f"燃气隧道炉与成型车间", 4000, 7000, py, ph, "TECH_EQUIP", 3)
        draw_box("冷却与无菌包装车间", 11500, 3500, py, ph, "TECH_EQUIP", 3)
    elif category == "乳制品/液态奶":
        draw_box("原奶化验与收储车间", 0, 3500, py, ph, "TECH_EQUIP", 3)
        draw_box("UHT杀菌机组车间", 4000, 4000, py, ph, "TECH_EQUIP", 3)
        draw_box("百级无菌灌装车间", 8500, 6500, py, ph, "TECH_EQUIP", 3)
        draw_box("CIP酸碱清洗站", 4000, 2500, py-3000, 2500, "TECH_EQUIP", 3)
    else: 
        draw_box("解冻斩拌与滚揉车间", 0, 4500, py, ph, "TECH_EQUIP", 3)
        draw_box("全自动烟熏蒸煮车间", 5000, 4500, py, ph, "TECH_EQUIP", 3)
        draw_box("-35℃速冻包装车间", 10000, 5000, py, ph, "TECH_EQUIP", 3)

    # 2. 厂区基建生态暴力拆解
    robot_x = end_x + 1500
    total_w = robot_x + 3000
    
    # 顶部：独立的仓储物流区
    draw_box("原料立体仓储库", -3000, 5000, py+ph+2000, 3500, "BUILDING_LOG", 6)
    draw_box("成品保温发货月台", robot_x, 5000, py+ph+2000, 3500, "BUILDING_LOG", 6)
    
    # 连接点：机械臂与 AGV 通道
    msp.add_circle((robot_x, py + 1700), radius=1500, dxfattribs={'layer': 'BUILDING_LOG', 'color': 2})
    msp.add_text("🤖自动码垛机械臂\n(衔接生产与仓储)", dxfattribs={'height': 120, 'style': 'CHS', 'color': 2}).set_placement((robot_x, py+3600), align=TextEntityAlignment.MIDDLE_CENTER)
    msp.add_line((-3000, py+ph+1000), (total_w, py+ph+1000), dxfattribs={'layer': 'BUILDING_LOG', 'color': 3, 'linetype': 'DASHED'})
    msp.add_text("🤖 AGV 全厂物流自动接驳通道", dxfattribs={'height': 150, 'style': 'CHS', 'color': 3}).set_placement((total_w/2, py+ph+1300))
    
    # 底部：极致细化的行政、生活与卫生区
    by = py - 4000; bh = 2500
    # -> 行政与生活群落
    draw_box("综合办公楼", -3000, 1800, by, bh, "BUILDING_LIFE", 7)
    draw_box("员工食堂", -1000, 1500, by, bh, "BUILDING_LIFE", 7)
    draw_box("倒班宿舍楼", 700, 1500, by, bh, "BUILDING_LIFE", 7)
    
    # -> 卫生隔离屏障 (严禁与食堂办公混流)
    draw_box("独立卫生间(带排风)", 2400, 1800, by, bh, "BUILDING_LIFE", 7, 120)
    draw_box("一更/二更/独立淋浴间", 4400, 2500, by, bh, "BUILDING_LIFE", 7, 120)
    draw_box("风淋室/强制洗手池", 7100, 2000, by, bh, "BUILDING_LIFE", 7, 120)
    
    # -> 动力枢纽
    draw_box("动力中心(配电/水处理/锅炉站)", 9300, 5700, by, bh, "BUILDING_LIFE", 7)

    # 3. 机电审查管线 (保留)
    drain_y = py - 300
    msp.add_line((0, drain_y), (end_x, drain_y), dxfattribs={'layer': 'PLUMBING_DRAIN', 'lineweight': 35})
    msp.add_text("💧 1.5% 坡度主地沟 (SUS304带水封)", dxfattribs={'height': 120, 'style': 'CHS', 'color': 5}).set_placement((1000, drain_y-200))
    if category in ["肉制品/切片", "饼干/烘焙"]: draw_box("油水分离器", end_x, 2000, drain_y-1500, 1500, "PLUMBING_DRAIN", 5)
    
    msp.add_line((4000, py+ph+300), (9000, py+ph+300), dxfattribs={'layer': 'HVAC_AIR', 'lineweight': 35})
    msp.add_text("♨️ 热加工负压 (-5 至 -10Pa)", dxfattribs={'height': 150, 'style': 'CHS', 'color': 1}).set_placement((6500, py+ph+500), align=TextEntityAlignment.MIDDLE_CENTER)
    
    corr_y = py - 800
    msp.add_line((-3000, corr_y), (total_w, corr_y), dxfattribs={'layer': 'ELEC_POWER', 'lineweight': 25})
    msp.add_text("⚡ 380V 主电缆桥架(独立回路)", dxfattribs={'height': 120, 'style': 'CHS', 'color': 2}).set_placement((-1500, corr_y+150))
    msp.add_line((-3000, corr_y-400), (total_w, corr_y-400), dxfattribs={'layer': 'HVAC_AIR', 'lineweight': 25})
    msp.add_text("♨️ 工业蒸汽主管 (带疏水阀/保温)", dxfattribs={'height': 120, 'style': 'CHS', 'color': 1}).set_placement((3000, corr_y-250))

    buf = io.StringIO(); doc.write(buf); return buf.getvalue()

# --- 4. 详版 Word 说明书 (精简展示代码，逻辑不变) ---
def generate_word_manual(project_name, category, capacity, ai_long_content):
    doc = Document(); style = doc.styles['Normal']; style.font.name = '宋体'; style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体'); style.font.size = Pt(11)
    doc.add_heading(f"{project_name}\n项目建议与工程审查详版说明书", 0).alignment = 1 
    
    doc.add_heading('一、 项目基准与总体概况', level=1)
    table = doc.add_table(rows=3, cols=2); table.style = 'Table Grid'
    table.cell(0, 0).text = "工艺品类"; table.cell(0, 1).text = category
    table.cell(1, 0).text = "处理产能"; table.cell(1, 1).text = f"{capacity} 吨/天"
    table.cell(2, 0).text = "验证标准"; table.cell(2, 1).text = "GB 14881 及机电强制国标"

    doc.add_heading('二、 全专业合规审查强制检查清单', level=1)
    specs = [
        "☑ 给排水：接触食品用水达直饮标准(RO/超滤)。主地沟深12-18cm，【1.5%下水坡度】。排污口强制配防鼠格栅与水封。",
        "☑ 电气动力：冷冻库、锅炉必须采用【独立回路】。照明设备加防护罩，应急照明即时启动。",
        "☑ 暖通压差：排烟与蒸汽物理分管。热加工区 -5至-10Pa(负压)，包装区 +3至+5Pa(正压)。",
        "☑ 建筑生态：行政办公、食堂、倒班宿舍必须与生产车间完全独立。车间入口必须经过【卫生间 -> 更衣/淋浴 -> 风淋】单向洗消防线。"
    ]
    for sp in specs: doc.add_paragraph(sp)

    doc.add_heading('三、 领域架构师深入工艺与运维建议', level=1)
    for p in ai_long_content.split('\n'):
        if p.strip() == "": continue
        if p.startswith('###') or p.startswith('**'): doc.add_heading(p.replace('#', '').replace('*', '').strip(), level=2)
        else: doc.add_paragraph(p.replace('*', '').strip())
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

# --- 5. Streamlit 界面构建 ---
st.set_page_config(page_title="总工机电基座", layout="wide")
s_price, c_price = fetch_market_data()
with st.sidebar:
    st.header("🗂️ 工艺与机电参数")
    proj_name = st.text_input("项目名称", "2026 高标机电一体化工厂")
    cat = st.selectbox("食品门类", ["饼干/烘焙", "乳制品/液态奶", "肉制品/切片"])
    cap = st.slider("日产吨数", 5, 50, 20)
    st.divider()

physics = calculate_physics(cap, cat); bom_df = generate_detailed_bom(cat, cap, physics)
total_cost = round(bom_df["预估造价"].sum() * 1.15, 1)

st.title("🏭 终极总包态：从车间走向智慧园区")

c1, c2, c3, c4 = st.columns(4)
c1.metric("含基建总投", f"￥{total_cost} 万"); c2.metric("日均耗水", f"{physics['water']} 吨")
c3.metric("总装机负荷", f"{physics['elec']} kW"); c4.metric("清单细项", len(bom_df))

st.subheader("📝 图纸生成与审查级说明书")
if st.button("🚀 撰写详版《全专业建议书》(输出超 1500 字)", use_container_width=True):
    with st.spinner('正在融合极度细化的宿舍、食堂与更衣淋浴路线...'):
        prompt = f"作为食品工程总工，为{cap}吨的{cat}工厂撰写机电与基建说明书。\n要求：包含【办公/宿舍/食堂的独立规划】、【强制更衣淋浴的风淋防线】、【1.5%排水与负压系统】、【机械臂与AGV物流衔接】。字数超1500字，纯文本格式。"
        st.session_state['word_content'] = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}]).choices[0].message.content
        st.success("1500+ 字硬核说明书已生成！")

st.divider(); st.subheader("📥 园区级总包一键下载")
e1, e2, e3 = st.columns(3)
excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer: bom_df.to_excel(writer, index=False)
e1.download_button("📊 导出极致 BOM(合规+运维)", excel_buf.getvalue(), "Granular_BOM.xlsx", use_container_width=True)
e2.download_button("📐 下载全景分离 CAD (建筑+机电)", generate_ultimate_dxf(cap, cat), "Granular_Layout.dxf", use_container_width=True)

if 'word_content' in st.session_state: e3.download_button("📄 下载详版说明书(Word)", generate_word_manual(proj_name, cat, cap, st.session_state['word_content']), "Granular_Manual.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else: e3.button("📄 请先生成说明书", disabled=True, use_container_width=True)
