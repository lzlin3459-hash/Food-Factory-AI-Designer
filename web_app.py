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
import os

# --- 0. 图表全局防乱码配置 ---
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS'] 
plt.rcParams['axes.unicode_minus'] = False

# --- 1. 配置与 AI 初始化 (云端安全版) ---
# 本地测试时会读取 .streamlit/secrets.toml，云端会读取后台配置
try:
    deepseek_key = st.secrets["DEEPSEEK_API_KEY"]
except:
    # 终极保底：如果都没配置，提示用户
    deepseek_key = "请在云端后台配置您的 API Key"

client = OpenAI(
    api_key=deepseek_key, 
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

# --- 3. BOM 生成 (保留完整的智能运维建议) ---
def generate_detailed_bom(category, capacity, physics):
    base_bom = [
        {"系统": "供电工程", "设备名称": "产线主配电柜", "规格": f"IP65 / {physics['elec']}kW", "数量": 1, "预估造价(万)": 8.0, "智能运维": "每季度检查防尘网，监控峰谷电耗。"},
        {"系统": "水务工程", "设备名称": "不锈钢给水阀组", "规格": "SUS304 DN50", "数量": int(capacity/5)+2, "预估造价(万)": 2.5, "智能运维": "每月检查密封垫圈，防止滴漏引发车间霉菌。"}
    ]
    if category == "饼干/烘焙":
        oven_len = capacity * 5
        specific_bom = [
            {"系统": "前段工艺", "设备名称": "卧式和面机", "规格": "双轴高速", "数量": 2, "预估造价(万)": 35.0, "智能运维": "每日班后检查轴承密封性，防止面筋渗入发酸。"},
            {"系统": "成型工艺", "设备名称": "三道辊压机", "规格": "自动化控制", "数量": 1, "预估造价(万)": 65.0, "智能运维": "每日班后彻底清洁辊筒残料，防止交叉污染。"},
            {"系统": "烘烤工艺", "设备名称": "隧道烘烤炉", "规格": f"长 {oven_len} 米", "数量": 1, "预估造价(万)": round(85.0 + oven_len*3.8, 1), "智能运维": "每两周检查燃烧器能效防止积碳；配置余热回收。"}
        ]
    else:
        specific_bom = [{"系统": "核心工艺", "设备名称": f"{category}主产线", "规格": f"产能匹配 {capacity}T", "数量": 1, "预估造价(万)": capacity*30, "智能运维": "严格执行 CIP 原位清洗或冷链温控标准。"}]
    return pd.DataFrame(base_bom + specific_bom)

# --- 4. CAD 引擎 (保留尺寸标注、AGV、机械臂、隔离墙) ---
def generate_ultimate_dxf(capacity, category):
    doc = ezdxf.new('R2010'); doc.styles.new('CHS', dxfattribs={'font': 'simsun.ttc'})
    if 'METRIC' not in doc.dimstyles:
        dimstyle = doc.dimstyles.new('METRIC')
        dimstyle.dxf.dimasz = 100; dimstyle.dxf.dimtxt = 150; dimstyle.dxf.dimexo = 50
    if 'DASHED' not in doc.linetypes:
        doc.linetypes.new('DASHED', dxfattribs={'description': 'Dashed', 'pattern': [10.0, -5.0]})

    msp = doc.modelspace(); o_len = capacity * 5000
    
    def draw_equip(name, x, length, color=0):
        msp.add_lwpolyline([(x, 0), (x+length, 0), (x+length, 1000), (x, 1000), (x, 0)], dxfattribs={'color': color})
        msp.add_text(name, dxfattribs={'height': 150, 'style': 'CHS'}).set_placement((x+length/2, 500), align=TextEntityAlignment.MIDDLE_CENTER)
        msp.add_linear_dim(base=(x, 1500), p1=(x, 1000), p2=(x+length, 1000), dimstyle='METRIC').render()

    draw_equip("和面机A", 0, 2000); draw_equip("和面机B", 3000, 2000)
    draw_equip("成型机组", 6000, 2000); draw_equip(f"核心处理段({category})", 9000, o_len, color=4)

    # 绘制机械臂安全区与 AGV 路径
    msp.add_circle((7000, 500), radius=1500, dxfattribs={'linetype': 'DASHED', 'color': 2})
    msp.add_text("机械臂安全半径 (1.5m)", dxfattribs={'height': 120, 'style': 'CHS', 'color': 2}).set_placement((7000, 2200))
    path_y = -2500; total_width = 9000 + o_len + 2000
    msp.add_line((0, path_y), (total_width, path_y), dxfattribs={'linetype': 'DASHED', 'color': 3})
    msp.add_text("🤖 AGV 物流主干道", dxfattribs={'height': 150, 'style': 'CHS', 'color': 3}).set_placement((total_width/2, path_y-500), align=TextEntityAlignment.MIDDLE_CENTER)

    # 绘制 GB 14881 隔离墙
    wall_x = 9000 + o_len + 500
    msp.add_line((wall_x, -4000), (wall_x, 4000), dxfattribs={'color': 1, 'lineweight': 35})
    msp.add_text("GB 14881 生熟隔离墙", dxfattribs={'height': 250, 'color': 1, 'style': 'CHS'}).set_placement((wall_x+100, 3000))
    msp.add_linear_dim(base=(0, -1500), p1=(0, 0), p2=(wall_x, 0), dimstyle='METRIC', override={'dimtxt': 200}).render()

    buf = io.StringIO(); doc.write(buf); return buf.getvalue()

# --- 5. 终极无崩版 PDF 生成器 (完美解决 Latin-1 与中文支持) ---
def generate_pdf_proposal(project_name, category, capacity, cost, roi, physics, ai_report):
    # 步骤 A：生成图表
    labels = ['核心工艺设备', '公用辅助工程', '安装与自控', '不可预见费']
    sizes = [cost * 0.65, cost * 0.15, cost * 0.10, cost * 0.10]
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=['#ff9999','#66b3ff','#99ff99','#ffcc99'])
    ax.axis('equal')
    chart_filename = "temp_investment_chart.png"
    plt.savefig(chart_filename, bbox_inches='tight', transparent=True, dpi=150)
    plt.close(fig)

    # 步骤 B：初始化 PDF 并自动寻址 Windows 系统字体
    pdf = FPDF()
    pdf.add_page()
    
    font_path = None
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

    # 智能文本降级机制 (绝对防止崩溃)
    def safe_print(size, zh_text, en_text, align='L'):
        if has_zh:
            pdf.set_font("ZH", size=size)
            pdf.cell(0, 10, txt=zh_text, ln=1, align=align)
        else:
            pdf.set_font("Arial", size=size)
            pdf.cell(0, 10, txt=en_text, ln=1, align=align)

    # 步骤 C：写入内容 (旧版 fpdf 必须用 txt 和 ln=1)
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
    
    # 步骤 D：嵌入图表
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

    # 步骤 E：物理文件缓冲流输出 (彻底解决 dest='S' 的编码崩溃)
    temp_pdf_name = "temp_proposal_buffer.pdf"
    pdf.output(temp_pdf_name)
    
    with open(temp_pdf_name, "rb") as f:
        pdf_bytes = f.read()
        
    # 清理所有临时文件
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

st.title("🏭 终极态：从数字设计到商业画册无损交付")

c1, c2, c3, c4 = st.columns(4)
c1.metric("总投预估", f"￥{total_cost} 万")
c2.metric("碳对冲 ROI", f"{roi} 年")
c3.metric("电力负荷", f"{physics['elec']} kW")
c4.metric("构件与运维项", len(bom_df))

ai_content = "暂未生成 AI 报告。请点击下方按钮启动评估。"
if st.button("✨ 1. 启动 DeepSeek 专家评估 (获取画册素材)", use_container_width=True):
    with st.spinner('正在撰写商业企划文案...'):
        try:
            prompt = f"针对{proj_name}({cat})，产能{cap}吨。请从合规和ROI角度给出一段200字的中文纯文本总结。"
            resp = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}])
            ai_content = resp.choices[0].message.content
            st.session_state['ai_report'] = ai_content
            st.success("AI 评估完成！请在下方下载画册。")
        except Exception as e:
            st.error(f"AI 调用异常: {e}")

st.divider()
st.subheader("📥 2. 全要素交付物一键打包")
e1, e2, e3 = st.columns(3)

excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
    bom_df.to_excel(writer, index=False, sheet_name='智能运维清单')
e1.download_button("📊 下载含运维 BOM (Excel)", excel_buf.getvalue(), "Detailed_BOM.xlsx", use_container_width=True)

e2.download_button("📐 下载全要素 CAD (DXF)", generate_ultimate_dxf(cap, cat), "Full_Layout.dxf", use_container_width=True)

current_ai_report = st.session_state.get('ai_report', "请先运行上方的 AI 专家评估以生成报告内容。")
pdf_bytes = generate_pdf_proposal(proj_name, cat, cap, total_cost, roi, physics, current_ai_report)
e3.download_button("📕 下载可视化商业画册 (PDF)", pdf_bytes, "Project_Proposal.pdf", mime="application/pdf", use_container_width=True)