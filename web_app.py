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

# --- 0. 图表全局强制防乱码配置 ---
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

# --- 2. 施工级 BOM 生成 (全面融入水电暖通审查要点) ---
def generate_detailed_bom(category, capacity, physics):
    base_bom = [
        {"专业": "给排水", "设备与材料": "RO反渗透处理站", "规格参数": f"产水 {physics['water']}t/d, 饮用水级", "数量": 1, "预估造价": 35.0, "合规说明": "设备直接接触水须达直饮标准"},
        {"专业": "给排水", "设备与材料": "车间防臭排水地沟系统", "规格参数": "SUS304, 坡度1.5%, 宽20cm", "数量": 1, "预估造价": 12.0, "合规说明": "水封地漏+6mm防鼠格栅+沉渣槽"},
        {"专业": "电气", "设备与材料": "三相380V主配电与独立回路箱", "规格参数": f"总负荷 {physics['elec']}kW", "数量": 1, "预估造价": 18.0, "合规说明": "锅炉与冷库必须采用独立回路控制"},
        {"专业": "电气", "设备与材料": "防水防尘防爆LED照明网", "规格参数": "IP65防护级", "数量": 1, "预估造价": 8.0, "合规说明": "食品正上方加防护罩，主电源失效秒级应急"},
        {"专业": "暖通", "设备与材料": "压差控制与排风系统", "规格参数": "热区排烟风速15m/s", "数量": 1, "预估造价": 25.0, "合规说明": "油烟与蒸汽严格分管，排烟罩伸出40cm"}
    ]
    if category == "饼干/烘焙":
        spec = [
            {"专业": "工艺", "设备与材料": "燃气隧道炉及预混线", "规格参数": f"{capacity}T/天", "数量": 1, "预估造价": 185.0, "合规说明": "废气余热回收"},
            {"专业": "给排水", "设备与材料": "耐高温隔油分离器", "规格参数": "热加工区专用", "数量": 1, "预估造价": 5.0, "合规说明": "防止油脂堵塞主干管"}
        ]
    elif category == "乳制品/液态奶":
        spec = [
            {"专业": "工艺", "设备与材料": "UHT杀菌与无菌灌装线", "规格参数": "百级净化", "数量": 1, "预估造价": 280.0, "合规说明": "蒸汽管道设2%下坡排冷凝水及疏水阀"},
            {"专业": "专项", "设备与材料": "全自动七步CIP清洗系统", "规格参数": "管道SUS316L, 介质85℃", "数量": 1, "预估造价": 45.0, "合规说明": "严格遵循乳制品卫生死角消除规范"}
        ]
    else: 
        spec = [
            {"专业": "工艺", "设备与材料": "解冻斩拌与深冷速冻库", "规格参数": "-35℃速冻", "数量": 1, "预估造价": 220.0, "合规说明": "动物性原料清洗池必须独立设置"},
            {"专业": "给排水", "设备与材料": "集中式热水洗涤站", "规格参数": "水温75℃, 给水1.5吋", "数量": 1, "预估造价": 15.0, "合规说明": "排水管径扩大至3吋防碎肉堵塞"}
        ]
    return pd.DataFrame(base_bom + spec)

# --- 3. 机电多图层 CAD 引擎 (一张图纸，多个专业层) ---
def generate_ultimate_dxf(capacity, category):
    doc = ezdxf.new('R2010'); doc.styles.new('CHS', dxfattribs={'font': 'simsun.ttc'})
    # 建立四大专业图层，支持施工队开关图层查看
    doc.layers.add("TECH_EQUIP", color=7) # 工艺设备 (白/黑)
    doc.layers.add("PLUMBING_DRAIN", color=5, linetype="DASHED") # 给排水 (蓝)
    doc.layers.add("ELEC_POWER", color=2, linetype="DASHED")  # 电气 (黄)
    doc.layers.add("HVAC_AIR", color=1, linetype="DASHED") # 暖通排气 (红)
    
    msp = doc.modelspace()
    def draw_box(name, x, length, y, h, layer, color, text_size=120):
        msp.add_lwpolyline([(x, y), (x+length, y), (x+length, y+h), (x, y+h), (x, y)], dxfattribs={'layer': layer, 'color': color})
        msp.add_text(name, dxfattribs={'height': text_size, 'style': 'CHS'}).set_placement((x+length/2, y+h/2), align=TextEntityAlignment.MIDDLE_CENTER)

    # 1. 核心工艺区规划 (TECH_EQUIP)
    py = 1000; ph = 3000; end_x = 14500
    if category == "饼干/烘焙":
        draw_box("原料预混区", 0, 3000, py, ph, "TECH_EQUIP", 7)
        draw_box("成型与隧道炉", 3500, 7000, py, ph, "TECH_EQUIP", 7)
        draw_box("冷却包装区", 11000, 3500, py, ph, "TECH_EQUIP", 7)
    elif category == "乳制品/液态奶":
        draw_box("原奶化验与收储", 0, 3000, py, ph, "TECH_EQUIP", 7)
        draw_box("UHT杀菌区", 3500, 4000, py, ph, "TECH_EQUIP", 7)
        draw_box("无菌灌装室", 8000, 6500, py, ph, "TECH_EQUIP", 7)
        draw_box("CIP清洗站", 3500, 2000, py-2500, 2000, "TECH_EQUIP", 7)
    else: 
        draw_box("解冻斩拌区", 0, 4000, py, ph, "TECH_EQUIP", 7)
        draw_box("烟熏蒸煮区", 4500, 4000, py, ph, "TECH_EQUIP", 7)
        draw_box("速冻冷库区", 9000, 5500, py, ph, "TECH_EQUIP", 7)

    # 2. 给排水设计意图 (PLUMBING_DRAIN)
    drain_y = py - 500
    msp.add_line((0, drain_y), (end_x, drain_y), dxfattribs={'layer': 'PLUMBING_DRAIN', 'lineweight': 30})
    msp.add_text("💧 主排污地沟 (SUS304/坡度1.5%) -> 流向低洁净区", dxfattribs={'height': 120, 'style': 'CHS', 'color': 5}).set_placement((1000, drain_y-250))
    if category in ["肉制品/切片", "饼干/烘焙"]:
        draw_box("隔油分离器", end_x+500, 1500, drain_y-1000, 1500, "PLUMBING_DRAIN", 5)

    # 3. 暖通与压差防线 (HVAC_AIR)
    air_y = py + ph + 500
    msp.add_line((0, air_y), (end_x, air_y), dxfattribs={'layer': 'HVAC_AIR', 'lineweight': 30})
    msp.add_text("♨️ 热加工负压区 (-5 至 -10Pa)", dxfattribs={'height': 150, 'style': 'CHS', 'color': 1}).set_placement((4000, air_y+300))
    msp.add_text("❄️ 冷却包装正压区 (+3 至 +5Pa)", dxfattribs={'height': 150, 'style': 'CHS', 'color': 1}).set_placement((11000, air_y+300))

    # 4. 电气点位布置 (ELEC_POWER)
    draw_box("⚡主配电室(380V)", -3500, 3000, py, ph, "ELEC_POWER", 2)
    # 分布式独立控制箱
    msp.add_circle((1500, py+ph), radius=300, dxfattribs={'layer': 'ELEC_POWER', 'color': 2})
    msp.add_circle((7000, py+ph), radius=300, dxfattribs={'layer': 'ELEC_POWER', 'color': 2})
    msp.add_text("独立设备配电箱/防水插座点", dxfattribs={'height': 100, 'style': 'CHS', 'color': 2}).set_placement((7500, py+ph+200))

    buf = io.StringIO(); doc.write(buf); return buf.getvalue()

# --- 4. 生成包含四大专业的公文级 Word ---
def generate_word_manual(project_name, category, capacity, ai_long_content):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = '宋体'; style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体'); style.font.size = Pt(11)

    title = doc.add_heading(f"{project_name}\n机电全专业审查与施工说明书", 0)
    title.alignment = 1 
    
    doc.add_heading('一、 项目基准参数', level=1)
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'
    records = [("涉及工艺品类", category), ("设计处理产能", f"{capacity} 吨/天"), ("施工验证标准", "GB 14881 及给排水/电气强制国标")]
    for i, (k, v) in enumerate(records):
        table.cell(i, 0).text = k; table.cell(i, 1).text = v

    doc.add_heading('二、 水电暖通四大专项规范', level=1)
    
    # 强行注入用户提供的硬核检查清单
    specs = [
        "1. 给排水系统：直接接触食品设备用水须达RO/超滤饮用标准。主管道采用不锈钢，支管为食品级PPR。车间主地沟深度12-18cm，宽度15-25cm，物理下坡度设计为 1.5%，严格从高洁净区流向低洁净区。所有排水口强制配置 6mm 以下网眼金属格栅及水封地漏防鼠防臭。热加工区后置必须串联油水分离器。",
        "2. 电气动力系统：实施三级供电（380V大型动力/220V常规/110V小型）。要求冷冻库、锅炉及高温线必须架设独立回路。暴露食品正上方灯具加装防护罩，冷链区强制使用防爆灯具，插座均需配备漏电保护且不得阻挡人车通行。",
        "3. 暖通与压差控制：油烟管道与蒸汽排风绝对物理隔离，排风罩外延至少 30-40cm，管道风速控制在 12-18m/s。车间实行压差屏障：热加工区维持 -5至-10Pa（负压），高洁净包装区维持 +3至+5Pa（正压）。蒸汽冷凝水管道需设置 2% 至 3% 的下坡并安装疏水阀。",
        "4. 特种与工艺辅助系统：乳制品/预制菜管路全面采用 SUS316L 材质。CIP清洗液循环温度不低于 85℃，必须支持全自动七步冲洗。全厂压缩空气端预留标准化快插接口。"
    ]
    for sp in specs: doc.add_paragraph(sp)

    doc.add_heading('三、 领域架构师综合建议', level=1)
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

physics = calculate_physics(cap, cat)
bom_df = generate_detailed_bom(cat, cap, physics)
# 彻底修复造价计算 BUG
total_cost = round(bom_df["预估造价"].sum() * 1.15, 1)

st.title("🏭 施工交付态：满足四大机电系统审查")

c1, c2, c3, c4 = st.columns(4)
c1.metric("含基建与机电总投", f"￥{total_cost} 万")
c2.metric("日均耗水预估", f"{physics['water']} 吨")
c3.metric("总装机负荷", f"{physics['elec']} kW")
c4.metric("施工材料与设备项", len(bom_df))

st.subheader("📝 施工图纸与技术标准生成 (支持审查清单验收)")
if st.button("🚀 依据机电标准撰写《全专业施工指导书》(约需 10 秒)", use_container_width=True):
    with st.spinner('正在融合给排水、电气、暖通与 CIP 标准...'):
        long_prompt = f"作为食品工程总工，为产能为{cap}吨的{cat}工厂补充一段指导建议。结合我们已经确定的水沟1.5%坡度、包装区+5Pa正压、316L管道材质等硬性规范，分析一下在实际施工穿管和设备进场时，最容易出现哪些安全和交叉污染隐患？请给出3点纯文本建议。"
        resp_long = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": long_prompt}])
        st.session_state['word_content'] = resp_long.choices[0].message.content
        st.success("机电参数与 AI 建议已成功写入 Word 缓冲池！")

st.divider()
st.subheader("📥 施工队发包一键下载")
e1, e2, e3 = st.columns(3)

excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer: bom_df.to_excel(writer, index=False)
e1.download_button("📊 导出含合规说明的 BOM", excel_buf.getvalue(), "MEP_BOM.xlsx", use_container_width=True)

e2.download_button("📐 下载多图层总图 (DXF)", generate_ultimate_dxf(cap, cat), "MEP_Layered_Layout.dxf", use_container_width=True)

if 'word_content' in st.session_state:
    word_bytes = generate_word_manual(proj_name, cat, cap, st.session_state['word_content'])
    e3.download_button("📄 下载全专业审查说明书(Word)", word_bytes, "MEP_Specifications.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    e3.button("📄 请先生成说明书", disabled=True, use_container_width=True)
