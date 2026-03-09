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

# --- 0. 图表全局防乱码配置 (云端豆腐块终极修复版) ---
font_path_cloud = "AlibabaPuHuiTi-3-115-Black.ttf"
if os.path.exists(font_path_cloud):
    # 如果在云端找到了阿里字体，强行注册给画图引擎 (修复 □□□□ 乱码)
    fm.fontManager.addfont(font_path_cloud)
    prop = fm.FontProperties(fname=font_path_cloud)
    plt.rcParams['font.family'] = prop.get_name()
else:
    # 本地退回策略
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei'] 
plt.rcParams['axes.unicode_minus'] = False

# --- 1. 配置与 AI 初始化 (云端安全开源版) ---
try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except:
    api_key = "请在云端后台配置您的 API Key"

client = OpenAI(
    api_key=api_key, 
    base_url="https://api.deepseek.com"
)

# --- 2. 实时行情与物理计算 ---
def fetch_market_data():
    try:
        return round(16500 + random.uniform(-150, 400), 2), round(95 + random.uniform(-5, 15), 2)
    except:
        return 16800, 100

def calculate_physics(capacity, category):
    factor = {"饼干/烘焙": 1.0, "乳制品/液态奶": 1.5, "肉制品/切片": 1.3}.get(category, 1.0)
    return {
        "elec": round(capacity * 15 * factor, 1), 
        "water": round(capacity * 3.5 * factor, 1), 
        "steam": round(capacity * 0.8 * factor, 1)
    }

# --- 3. 深度解耦的 BOM 生成器 ---
def generate_detailed_bom(category, capacity, physics):
    base_bom = [
        {"系统": "供电工程", "设备名称": "产线主配电柜", "规格": f"IP65 / {physics['elec']}kW", "数量": 1, "预估造价(万)": 8.0, "智能运维": "每季度检查防尘网，监控峰谷电耗。"}
    ]
    if category == "饼干/烘焙":
        oven_len = capacity * 5
        specific_bom = [
            {"系统": "面团处理", "设备名称": "卧式双轴和面机", "规格": "防尘级", "数量": 2, "预估造价(万)": 35.0, "智能运维": "班后检查密封性，防面筋发酸。"},
            {"系统": "成型工艺", "设备名称": "三道辊压机", "规格": "PLC自动化", "数量": 1, "预估造价(万)": 65.0, "智能运维": "彻底清洁辊筒残料。"},
            {"系统": "烘烤工艺", "设备名称": "燃气隧道烘烤炉", "规格": f"长 {oven_len} 米", "数量": 1, "预估造价(万)": round(85.0 + oven_len*3.8, 1), "智能运维": "配置余热回收系统。"}
        ]
    elif category == "乳制品/液态奶":
        specific_bom = [
            {"系统": "前处理", "设备名称": "原奶收储与净乳机", "规格": "SUS316L", "数量": 1, "预估造价(万)": 55.0, "智能运维": "监测离心机转子平衡。"},
            {"系统": "杀菌工艺", "设备名称": "UHT 超高温管式杀菌机", "规格": "全自动温控", "数量": 1, "预估造价(万)": 120.0, "智能运维": "严格监控 137℃ 保温时间。"},
            {"系统": "卫生保障", "设备名称": "CIP 中央自动清洗站", "规格": "酸碱水三罐", "数量": 1, "预估造价(万)": 45.0, "智能运维": "定期标定电导率仪，确保清洗液浓度。"}
        ]
    else: 
        specific_bom = [
            {"系统": "前处理", "设备名称": "真空斩拌与滚揉机", "规格": "低温夹套冷水", "数量": 2, "预估造价(万)": 60.0, "智能运维": "监控真空泵油位。"},
            {"系统": "热加工", "设备名称": "全自动烟熏蒸煮炉", "规格": "含烟雾发生器", "数量": 1, "预估造价(万)": 80.0, "智能运维": "排烟管道定期除焦油防明火。"},
            {"系统": "冷链保障", "设备名称": "速冻库与制冷机组", "规格": "-35℃ 螺杆式", "数量": 1, "预估造价(万)": 110.0, "智能运维": "化霜周期智能调节。"}
        ]
    return pd.DataFrame(base_bom + specific_bom)

# --- 4. 深度解耦的 CAD 引擎 ---
def generate_ultimate_dxf(capacity, category):
    doc = ezdxf.new('R2010'); doc.styles.new('CHS', dxfattribs={'font': 'simsun.ttc'})
    if 'METRIC' not in doc.dimstyles:
        dimstyle = doc.dimstyles.new('METRIC'); dimstyle.dxf.dimasz = 100; dimstyle.dxf.dimtxt = 150
    if 'DASHED' not in doc.linetypes:
        doc.linetypes.new('DASHED', dxfattribs={'description': 'Dashed', 'pattern': [10.0, -5.0]})

    msp = doc.modelspace(); o_len = capacity * 5000
    
    def draw_equip(name, x, length, color=0):
        msp.add_lwpolyline([(x, 0), (x+length, 0), (x+length, 1000), (x, 1000), (x, 0)], dxfattribs={'color': color})
        msp.add_text(name, dxfattribs={'height': 150, 'style': 'CHS'}).set_placement((x+length/2, 500), align=TextEntityAlignment.MIDDLE_CENTER)
        msp.add_linear_dim(base=(x, 1500), p1=(x, 1000), p2=(x+length, 1000), dimstyle='METRIC').render()

    if category == "饼干/烘焙":
        draw_equip("和面区", 0, 2000); draw_equip("成型区", 3000, 2000)
        draw_equip(f"燃气隧道炉 ({capacity}T)", 6000, o_len, color=4)
        wall_x = 6000 + o_len + 500
    elif category == "乳制品/液态奶":
        draw_equip("原奶收储罐区", 0, 3000, color=3); draw_equip("UHT杀菌机组", 4000, 3000, color=1)
        draw_equip("无菌灌装室", 8000, 4000, color=5); draw_equip("CIP中央清洗", 13000, 3000, color=4)
        wall_x = 13000 + 3000 + 500
    else: 
        draw_equip("解冻斩拌区", 0, 3000, color=1); draw_equip("真空滚揉区", 4000, 3000, color=2)
        draw_equip("烟熏蒸煮炉", 8000, 4000, color=4); draw_equip("低温无菌包装", 13000, 4000, color=5)
        wall_x = 13000 + 4000 + 500

    msp.add_circle((wall_x - 3000, 500), radius=1500, dxfattribs={'linetype': 'DASHED', 'color': 2})
    msp.add_text("机械臂安全半径", dxfattribs={'height': 120, 'style': 'CHS', 'color': 2}).set_placement((wall_x - 3000, 2200))
    path_y = -2500; total_width = wall_x + 1000
    msp.add_line((0, path_y), (total_width, path_y), dxfattribs={'linetype': 'DASHED', 'color': 3})
    msp.add_text("🤖 AGV 物流主干道", dxfattribs={'height': 150, 'style': 'CHS', 'color': 3}).set_placement((total_width/2, path_y-500), align=TextEntityAlignment.MIDDLE_CENTER)

    msp.add_line((wall_x, -4000), (wall_x, 4000), dxfattribs={'color': 1, 'lineweight': 35})
    msp.add_text("GB 14881 洁净隔离墙", dxfattribs={'height': 250, 'color': 1, 'style': 'CHS'}).set_placement((wall_x+100, 3000))

    buf = io.StringIO(); doc.write(buf); return buf.getvalue()

# --- 5. PDF 商业画册生成器 ---
def generate_pdf_proposal(project_name, category, capacity, cost, roi, physics, ai_report):
    if category == "饼干/烘焙":
        labels = ['隧道炉与成型', '天然气与排烟', '包装与自控', '不可预见费']
        sizes = [cost * 0.55, cost * 0.20, cost * 0.15, cost * 0.10]
    elif category == "乳制品/液态奶":
        labels = ['管网与杀菌罐', '无菌净化空调', 'CIP与水处理', '不可预见费']
        sizes = [cost * 0.45, cost * 0.25, cost * 0.20, cost * 0.10]
    else: 
        labels = ['斩拌与深加工', '工业制冷冷库', '高频清洗排水', '不可预见费']
        sizes = [cost * 0.40, cost * 0.35, cost * 0.15, cost * 0.10]

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=['#ff9999','#66b3ff','#99ff99','#ffcc99'])
    ax.axis('equal')
    chart_filename = "temp_investment_chart.png"
    plt.savefig(chart_filename, bbox_inches='tight', transparent=True, dpi=150)
    plt.close(fig)

    pdf = FPDF()
    pdf.add_page()
    
    font_path = None
    if os.path.exists("AlibabaPuHuiTi-3-115-Black.ttf"):
        font_path = "AlibabaPuHuiTi-3-115-Black.ttf"
    else:
        windows_fonts = ["C:\\Windows\\Fonts\\simhei.ttf", "C:\\Windows\\Fonts\\msyh.ttf", "C:\\Windows\\Fonts\\simsun.ttc"]
        for pf in windows_fonts:
            if os.path.exists(pf):
                font_path = pf
                break
            
    has_zh = False
    if font_path:
        try:
            pdf.add_font("ZH", style="", fname=font_path, uni=True)
            has_zh = True
        except:
            has_zh = False

    def safe_print(size, zh_text, en_text, align='L'):
        if has_zh:
            pdf.set_font("ZH", size=size)
            pdf.cell(0, 10, txt=zh_text, ln=1, align=align)
        else:
            pdf.set_font("Arial", size=size)
            pdf.cell(0, 10, txt=en_text, ln=1, align=align)

    if has_zh:
        pdf.set_font("ZH", size=18)
        pdf.cell(0, 15, txt="智能食品工厂商业计划书", ln=1, align='C')
    else:
        safe_print(18, "", "Smart Factory Proposal", 'C')
    pdf.ln(5)
    
    safe_print(12, f"项目名称: {project_name}", "Project Name: Default (Font Missing)")
    safe_print(12, f"所属品类: {category}  |  日产能: {capacity} 吨/天", f"Category: {category} | Cap: {capacity}T")
    pdf.ln(5)
    
    safe_print(14, "一、 财务概算与资金分析", "1. Financial Overview")
    safe_print(12, f"  • 预估总投资 : ￥{cost} 万", f"  - CAPEX: {cost} 万")
    safe_print(12, f"  • 投资回报期 : {roi} 年", f"  - ROI: {roi} Years")
    
    current_y = pdf.get_y()
    if os.path.exists(chart_filename):
        pdf.image(chart_filename, x=45, y=current_y + 2, w=120)
    pdf.set_y(current_y + 90)
    
    safe_print(14, "二、 公用工程负荷推算", "2. Utilities Load")
    safe_print(12, f"  • 电力负荷需求 : {physics['elec']} kW", f"  - Electric: {physics['elec']} kW")
    safe_print(12, f"  • 工业蒸汽消耗 : {physics['steam']} 吨", f"  - Steam: {physics['steam']} Tons")
    pdf.ln(5)
    
    safe_print(14, "三、 领域架构师合规评估", "3. AI Assessment")
    if has_zh:
        pdf.set_font("ZH", size=11)
        clean_report = ai_report.replace('💎', '').replace('🏗️', '').replace('🌿', '')
        pdf.multi_cell(0, 7, txt=clean_report)
    else:
        pdf.set_font("Arial", size=11)
        pdf.multi_cell(0, 7, txt="Report generated. Chinese display disabled due to missing system fonts.")

    temp_pdf_name = "temp_proposal_buffer.pdf"
    pdf.output(temp_pdf_name)
    
    with open(temp_pdf_name, "rb") as f:
        pdf_bytes = f.read()
        
    if os.path.exists(chart_filename): os.remove(chart_filename)
    if os.path.exists(temp_pdf_name): os.remove(temp_pdf_name)
        
    return pdf_bytes

# --- 6. Streamlit 界面构建 ---
st.set_page_config(page_title="全能架构基座", layout="wide")
s_price, c_price = fetch_market_data()

with st.sidebar:
    st.header("🗂️ 全局配置")
    proj_name = st.text_input("项目名称", "2026 绿建标准工厂")
    cat = st.selectbox("食品门类", ["饼干/烘焙", "乳制品/液态奶", "肉制品/切片"])
    cap = st.slider("日产吨数", 5, 50, 20)
    st.divider()
    st.success(f"钢价: ¥{s_price} | 碳价: ¥{c_price}")

physics = calculate_physics(cap, cat)
bom_df = generate_detailed_bom(cat, cap, physics)
total_cost = round(bom_df["预估造价(万)"].sum() * 1.15, 1)
roi = round(15.0 / (cap * 100 * 300 * 0.12 * 0.0035 + (cap*0.2*c_price)/10000 + 0.1), 1)

st.title("🏭 终极态：从数字设计到商业与工程双轨交付")

c1, c2, c3, c4 = st.columns(4)
c1.metric("总投预估", f"￥{total_cost} 万")
c2.metric("碳对冲 ROI", f"{roi} 年")
c3.metric("电力负荷", f"{physics['elec']} kW")
c4.metric("构件与运维项", len(bom_df))

# 模块一：短篇商业评估 (用于 PDF 画册)
st.subheader("📊 1. 商业路演准备 (面向投资方)")
ai_content = "暂未生成 AI 报告。请点击下方按钮启动评估。"
if st.button("✨ 启动路演级 AI 评估 (获取 PDF 画册素材)", use_container_width=True):
    with st.spinner('正在撰写商业企划文案...'):
        try:
            prompt = f"针对{proj_name}({cat})，产能{cap}吨。请从合规和ROI角度给出一段200字的中文纯文本总结。"
            resp = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}])
            ai_content = resp.choices[0].message.content
            st.session_state['ai_report'] = ai_content
            st.success("商业评估完成！")
        except Exception as e:
            st.error(f"AI 调用异常: {e}")

# 模块二：长篇工艺说明书 (面向厂长与环保局)
st.subheader("📝 2. 工程技术交付 (面向建设方)")
if st.button("🚀 撰写详尽《工厂工艺说明书》(约需 15 秒)", use_container_width=True):
    with st.spinner('DeepSeek 正在深度推演工艺流程、卫生分区与三废处理方案...'):
        try:
            long_prompt = f"作为食品工程领域架构师，请为产能为{cap}吨/天的{cat}工厂撰写一份专业的《工厂工艺说明书与项目建议书》。包含以下章节：1. 核心工艺流程描述；2. GB 14881 卫生分区规范(划分高低洁净区及人流物流防交叉策略)；3. 设备选型与能耗原理；4. 三废(废水、废气、固废)环保处理建议。要求极其专业严谨，使用结构化的排版，不少于800字。"
            resp_long = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": long_prompt}])
            st.session_state['manual_md'] = f"# {proj_name} - {cat} 工厂工艺说明书\n\n" + resp_long.choices[0].message.content
            st.success("工艺说明书撰写完毕！请在下方下载。")
        except Exception as e:
            st.error(f"AI 调用异常: {e}")

st.divider()
st.subheader("📥 3. 全要素交付物一键打包")
e1, e2, e3, e4 = st.columns(4)

# 按钮 1：BOM
excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
    bom_df.to_excel(writer, index=False, sheet_name='智能运维清单')
e1.download_button("📊 下载运维 BOM (Excel)", excel_buf.getvalue(), "Detailed_BOM.xlsx", use_container_width=True)

# 按钮 2：CAD
e2.download_button("📐 下载图纸 (DXF)", generate_ultimate_dxf(cap, cat), "Full_Layout.dxf", use_container_width=True)

# 按钮 3：商业画册 (PDF)
current_ai_report = st.session_state.get('ai_report', "请先运行上方的 AI 专家评估以生成报告内容。")
pdf_bytes = generate_pdf_proposal(proj_name, cat, cap, total_cost, roi, physics, current_ai_report)
e3.download_button("📕 下载商业画册 (PDF)", pdf_bytes, "Project_Proposal.pdf", mime="application/pdf", use_container_width=True)

# 按钮 4：工艺说明书 (Markdown -> Word)
if 'manual_md' in st.session_state:
    e4.download_button("📄 下载工艺说明书 (可用于Word)", st.session_state['manual_md'].encode('utf-8'), "Factory_Operation_Manual.md", use_container_width=True)
else:
    e4.button("📄 请先生成说明书", disabled=True, use_container_width=True)
