import streamlit as st
import ezdxf
from ezdxf.enums import TextEntityAlignment
import pandas as pd
import plotly.graph_objects as go
from openai import OpenAI
import io
import datetime
import random
from fpdf import FPDF
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn

# --- 0. 图表全局强制防乱码配置 ---
font_path_cloud = "AlibabaPuHuiTi-3-115-Black.ttf"
if os.path.exists(font_path_cloud):
    prop = fm.FontProperties(fname=font_path_cloud)
else:
    prop = fm.FontProperties(family='SimHei') # 本地后备
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

# --- 2. BOM 生成 (细化水电气与基建) ---
def generate_detailed_bom(category, capacity, physics):
    base_bom = [
        {"系统": "动力中心", "设备名称": "配电室与变压器组", "规格": f"{physics['elec']}kW", "数量": 1, "预估造价": 15.0, "智能运维": "监控峰谷电耗与功率因数"},
        {"系统": "动力中心", "设备名称": "RO反渗透与锅炉站", "规格": f"产汽 {physics['steam']}t/d", "数量": 1, "预估造价": 30.0, "智能运维": "水质硬度每日检测"},
        {"系统": "物流与基建", "设备名称": "立体仓储(原辅料/成品)", "规格": "WMS智能调度", "数量": 1, "预估造价": 65.0, "智能运维": "温湿度与防虫鼠三级控制"},
        {"系统": "卫生与生活", "设备名称": "三级更衣/风淋/办公宿舍", "规格": "防交叉隔离带", "数量": 1, "预估造价": 150.0, "智能运维": "人流通道紫外线定时消杀"}
    ]
    if category == "饼干/烘焙":
        spec = [{"系统": "核心工艺", "设备名称": "和面成型与隧道炉组", "规格": f"{capacity}T/天", "数量": 1, "预估造价": 185.0, "智能运维": "余热回收与辊筒清洁"}]
    elif category == "乳制品/液态奶":
        spec = [{"系统": "核心工艺", "设备名称": "UHT杀菌与无菌灌装线", "规格": "百级净化", "数量": 1, "预估造价": 280.0, "智能运维": "CIP原位清洗及滤网压差监控"}]
    else: 
        spec = [{"系统": "核心工艺", "设备名称": "斩拌蒸煮与全厂冷链", "规格": "-35℃速冻", "数量": 1, "预估造价": 220.0, "智能运维": "化霜周期调节与真空度监控"}]
    return pd.DataFrame(base_bom + spec)

# --- 3. 终极 CAD 引擎：细化分区与水电气管廊系统 ---
def generate_ultimate_dxf(capacity, category):
    doc = ezdxf.new('R2010'); doc.styles.new('CHS', dxfattribs={'font': 'simsun.ttc'})
    doc.layers.add("WATER", color=5, linetype="DASHED") # 蓝水
    doc.layers.add("STEAM", color=1, linetype="DASHED") # 红汽
    doc.layers.add("ELEC", color=2, linetype="DASHED")  # 黄电
    
    msp = doc.modelspace()
    def draw_equip(name, x, length, y=0, h=1500, color=0):
        msp.add_lwpolyline([(x, y), (x+length, y), (x+length, y+h), (x, y+h), (x, y)], dxfattribs={'color': color})
        msp.add_text(name, dxfattribs={'height': 150, 'style': 'CHS'}).set_placement((x+length/2, y+h/2), align=TextEntityAlignment.MIDDLE_CENTER)

    # 1. 动力与公用工程中心 (位于厂房北部顶部)
    uy = 4500; uh = 2000
    draw_equip("⚡主配电与空压站", 0, 3000, y=uy, h=uh, color=2)
    draw_equip("♨️燃气锅炉站", 3500, 3000, y=uy, h=uh, color=1)
    draw_equip("💧纯水净化与废水站", 7000, 4000, y=uy, h=uh, color=5)

    # 2. 核心主工艺区 (动态分配细致区间)
    end_x = 0; py = 1000; ph = 2500
    if category == "饼干/烘焙":
        draw_equip("辅料预混区", 0, 2500, y=py, h=ph, color=7); draw_equip("打面醒发区", 3000, 2500, y=py, h=ph, color=1)
        draw_equip("成型与隧道炉", 6000, 5000, y=py, h=ph, color=4); draw_equip("冷却与理料", 11500, 3000, y=py, h=ph, color=3)
        end_x = 14500
    elif category == "乳制品/液态奶":
        draw_equip("原奶化验与收储", 0, 3000, y=py, h=ph, color=7); draw_equip("脱气与均质区", 3500, 3000, y=py, h=ph, color=1)
        draw_equip("UHT杀菌管网", 7000, 3000, y=py, h=ph, color=4); draw_equip("百级无菌灌装室", 10500, 4000, y=py, h=ph, color=5)
        draw_equip("CIP酸碱罐区", 7000, 2000, y=py-2500, h=2000, color=6) # CIP 紧贴杀菌区下方
        end_x = 14500
    else: 
        draw_equip("解冻与缓化库", 0, 3000, y=py, h=ph, color=7); draw_equip("真空斩拌与滚揉", 3500, 3000, y=py, h=ph, color=1)
        draw_equip("全自动烟熏炉区", 7000, 3000, y=py, h=ph, color=4); draw_equip("深冷速冻包装区", 10500, 4000, y=py, h=ph, color=5)
        end_x = 14500

    # 3. 包装、机械臂与立体库
    robot_x = end_x + 1500
    msp.add_circle((robot_x, py + 1200), radius=1200, dxfattribs={'linetype': 'DASHED', 'color': 2})
    msp.add_text("🤖自动码垛机械臂", dxfattribs={'height': 120, 'style': 'CHS', 'color': 2}).set_placement((robot_x, py+2600), align=TextEntityAlignment.MIDDLE_CENTER)
    
    draw_equip("原料立体高架库", -4000, 3500, y=py, h=ph+1500, color=6)
    draw_equip("成品保温发货库", robot_x + 2000, 4000, y=py, h=ph+1500, color=6)

    # 4. 人员动线与生活区 (严格隔离在南部)
    hy = -3500; hh = 2000; total_width = robot_x + 6000
    draw_equip("综合办公与研发中心", -4000, 5000, y=hy, h=hh, color=7)
    draw_equip("员工食堂与宿舍楼", 1500, 4000, y=hy, h=hh, color=7)
    draw_equip("一更/二更/独立厕所/风淋", 6000, 6000, y=hy, h=hh, color=3) # 卫生隔离带
    msp.add_line((6000, hy+hh), (6000, py), dxfattribs={'color': 3})
    msp.add_text("👤洁净人流通道", dxfattribs={'height': 120, 'style': 'CHS'}).set_placement((5800, -500))

    # 5. 水电气公用工程“主管廊”排布 (工业设计的浪漫)
    corridor_y = 4000
    msp.add_line((0, corridor_y), (total_width-4000, corridor_y), dxfattribs={'layer': 'ELEC', 'lineweight': 30})
    msp.add_text("⚡ 380V 主电缆桥架", dxfattribs={'height': 120, 'style': 'CHS', 'color': 2}).set_placement((1500, corridor_y+150))
    
    msp.add_line((3500, corridor_y-300), (total_width-4000, corridor_y-300), dxfattribs={'layer': 'STEAM', 'lineweight': 30})
    msp.add_text("♨️ 0.6MPa 工业蒸汽主管", dxfattribs={'height': 120, 'style': 'CHS', 'color': 1}).set_placement((5000, corridor_y-150))
    
    msp.add_line((7000, corridor_y-600), (total_width-4000, corridor_y-600), dxfattribs={'layer': 'WATER', 'lineweight': 30})
    msp.add_text("💧 RO纯水与供水主管", dxfattribs={'height': 120, 'style': 'CHS', 'color': 5}).set_placement((8500, corridor_y-450))

    buf = io.StringIO(); doc.write(buf); return buf.getvalue()

# --- 4. 商业画册与原生高规格 Word 说明书生成器 ---
def generate_pdf_proposal(project_name, category, capacity, cost, roi, physics, ai_report):
    labels = ['核心工艺线', '厂房土建与基建', '冷链/空调辅助', '不可预见费']
    if category == "饼干/烘焙": sizes = [cost * 0.40, cost * 0.35, cost * 0.15, cost * 0.10]
    elif category == "乳制品/液态奶": sizes = [cost * 0.45, cost * 0.25, cost * 0.20, cost * 0.10]
    else: sizes = [cost * 0.35, cost * 0.30, cost * 0.25, cost * 0.10]

    fig, ax = plt.subplots(figsize=(6, 4))
    # 强制图表字体
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=['#ff9999','#66b3ff','#99ff99','#ffcc99'], textprops={'fontproperties': prop})
    ax.axis('equal')
    chart_filename = "temp_investment_chart.png"
    plt.savefig(chart_filename, bbox_inches='tight', transparent=True, dpi=150)
    plt.close(fig)

    pdf = FPDF(); pdf.add_page()
    has_zh = False
    if os.path.exists(font_path_cloud):
        try: pdf.add_font("ZH", style="", fname=font_path_cloud, uni=True); has_zh = True
        except: pass

    def safe_print(size, zh_text, en_text, align='L'):
        if has_zh: pdf.set_font("ZH", size=size); pdf.cell(0, 10, txt=zh_text, ln=1, align=align)
        else: pdf.set_font("Arial", size=size); pdf.cell(0, 10, txt=en_text, ln=1, align=align)

    if has_zh: pdf.set_font("ZH", size=18); pdf.cell(0, 15, txt="食品工厂总图规划与商业计划", ln=1, align='C')
    else: safe_print(18, "", "Factory Masterplan", 'C')
    
    safe_print(12, f"项目名称: {project_name} | 所属品类: {category} | 产能: {capacity}T", f"Category: {category}")
    safe_print(14, "一、 财务概算与全厂投资占比", "1. Financial Overview")
    safe_print(12, f"  • 预估总投资 : ￥{cost} 万 (含土建公用)", f"  - CAPEX: {cost}")
    
    if os.path.exists(chart_filename): pdf.image(chart_filename, x=45, y=pdf.get_y()+2, w=120)
    pdf.set_y(pdf.get_y() + 90)
    
    safe_print(14, "二、 领域架构师合规评估", "2. AI Assessment")
    if has_zh:
        pdf.set_font("ZH", size=11)
        pdf.multi_cell(0, 7, txt=ai_report.replace('💎', '').replace('🏗️', '').replace('🌿', ''))
    
    pdf.output("temp_pdf.pdf")
    with open("temp_pdf.pdf", "rb") as f: pdf_bytes = f.read()
    if os.path.exists(chart_filename): os.remove(chart_filename)
    if os.path.exists("temp_pdf.pdf"): os.remove("temp_pdf.pdf")
    return pdf_bytes

def generate_word_manual(project_name, category, capacity, ai_long_content):
    """极致专业的结构化原生 Word 引擎"""
    doc = Document()
    
    # 全局强制中文字体（宋体）
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(11)

    # 封面标题
    title = doc.add_heading(f"{project_name}\n项目建议书与全厂工艺规范", 0)
    title.alignment = 1 # 居中
    
    # 结构化表格：项目概况
    doc.add_heading('一、 项目基本概况', level=1)
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'
    records = [("项目归属领域", category), ("设计目标产能", f"{capacity} 吨/天"), ("设计执行标准", "GB 14881 食品安全国家标准")]
    for i, (k, v) in enumerate(records):
        table.cell(i, 0).text = k
        table.cell(i, 1).text = v

    # 核心专业内容
    doc.add_heading('二、 全厂工艺与工程说明', level=1)
    # 简单清洗并分段写入 AI 内容
    paragraphs = ai_long_content.split('\n')
    for p in paragraphs:
        if p.strip() == "": continue
        if p.startswith('###') or p.startswith('**'):
            doc.add_heading(p.replace('#', '').replace('*', '').strip(), level=2)
        else:
            doc.add_paragraph(p.replace('*', '').strip())

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# --- 5. Streamlit 界面构建 ---
st.set_page_config(page_title="总工程师基座", layout="wide")
s_price, c_price = fetch_market_data()

with st.sidebar:
    st.header("🗂️ 总图规划参数")
    proj_name = st.text_input("项目名称", "2026 标准智慧工厂")
    cat = st.selectbox("食品门类", ["饼干/烘焙", "乳制品/液态奶", "肉制品/切片"])
    cap = st.slider("日产吨数", 5, 50, 20)
    st.divider()
    st.success(f"钢价: ¥{s_price} | 碳价: ¥{c_price}")

physics = calculate_physics(cap, cat)
bom_df = generate_detailed_bom(cat, cap, physics)
total_cost = round(bom_df["预估造价(万)"].sum() * 1.15, 1)
roi = round(15.0 / (cap * 100 * 300 * 0.12 * 0.0035 + (cap*0.2*c_price)/10000 + 0.1), 1)

st.title("🏭 全要素基座：不仅是产线，更是工业生态")

c1, c2, c3, c4 = st.columns(4)
c1.metric("含基建总投", f"￥{total_cost} 万")
c2.metric("碳对冲 ROI", f"{roi} 年")
c3.metric("装机负荷", f"{physics['elec']} kW")
c4.metric("全厂清单项", len(bom_df))

st.subheader("📊 1. 商业路演准备 (面向资方)")
ai_content = "暂未生成。"
if st.button("✨ 启动路演级评估 (修复图表乱码)", use_container_width=True):
    with st.spinner('正在撰写商业企划文案...'):
        prompt = f"针对{proj_name}({cat})，产能{cap}吨。从全厂基建投资和ESG角度给出200字纯文本总结。"
        resp = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}])
        st.session_state['ai_report'] = resp.choices[0].message.content
        st.success("商业评估完成！豆腐块乱码已修复。")

st.subheader("📝 2. 工程技术交付 (生成标准公文级 Word)")
if st.button("🚀 撰写公文级《全厂项目建议书》(约需 15 秒)", use_container_width=True):
    with st.spinner('DeepSeek 正在严格按照国家工程标准推演厂房规范...'):
        long_prompt = f"作为食品工程总工，为{cap}吨/天的{cat}工厂撰写《项目建议书》。严禁废话，包含四个核心结构：1. 【核心工艺流程】详解设备；2. 【GB14881 全厂卫生分区】细化车间、一二更、风淋、厕所布局及人流路线；3. 【公用工程(水电气)消耗分配】说明动力中心的作用；4. 【三废处理方案】。不少于800字，纯文本格式输出。"
        resp_long = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": long_prompt}])
        st.session_state['word_content'] = resp_long.choices[0].message.content
        st.success("公文级 Word 说明书撰写完毕！")

st.divider()
st.subheader("📥 3. 工程师全家桶一键打包")
e1, e2, e3, e4 = st.columns(4)

excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer: bom_df.to_excel(writer, index=False)
e1.download_button("📊 导出总包 BOM", excel_buf.getvalue(), "Total_Factory_BOM.xlsx", use_container_width=True)

e2.download_button("📐 下载管网级 CAD", generate_ultimate_dxf(cap, cat), "Piping_Layout.dxf", use_container_width=True)

current_ai_report = st.session_state.get('ai_report', "请先运行 AI 评估。")
e3.download_button("📕 下载商业画册", generate_pdf_proposal(proj_name, cat, cap, total_cost, roi, physics, current_ai_report), "Business_Plan.pdf", mime="application/pdf", use_container_width=True)

if 'word_content' in st.session_state:
    word_bytes = generate_word_manual(proj_name, cat, cap, st.session_state['word_content'])
    e4.download_button("📄 下载项目建议书(Word)", word_bytes, "Project_Proposal.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    e4.button("📄 请先生成建议书", disabled=True, use_container_width=True)
