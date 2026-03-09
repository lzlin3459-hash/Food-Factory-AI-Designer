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
from docx import Document  # 新增：原生 Word 引擎

# --- 0. 图表全局防乱码配置 ---
font_path_cloud = "AlibabaPuHuiTi-3-115-Black.ttf"
if os.path.exists(font_path_cloud):
    fm.fontManager.addfont(font_path_cloud)
    prop = fm.FontProperties(fname=font_path_cloud)
    plt.rcParams['font.family'] = prop.get_name()
else:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei'] 
plt.rcParams['axes.unicode_minus'] = False

# --- 1. 配置与 AI 初始化 ---
try:
    api_key = st.secrets["DEEPSEEK_API_KEY"]
except:
    api_key = "请在云端后台配置您的 API Key"

client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

def fetch_market_data():
    try: return round(16500 + random.uniform(-150, 400), 2), round(95 + random.uniform(-5, 15), 2)
    except: return 16800, 100

def calculate_physics(capacity, category):
    factor = {"饼干/烘焙": 1.0, "乳制品/液态奶": 1.5, "肉制品/切片": 1.3}.get(category, 1.0)
    return {"elec": round(capacity * 15 * factor, 1), "water": round(capacity * 3.5 * factor, 1), "steam": round(capacity * 0.8 * factor, 1)}

# --- 2. BOM 生成器 (新增全厂基础设施) ---
def generate_detailed_bom(category, capacity, physics):
    # 基础公用与全厂基建 (不管什么品类都必须有)
    base_bom = [
        {"系统": "动力工程", "设备名称": "全厂主配电与变压器", "规格": f"{physics['elec']}kW", "数量": 1, "预估造价(万)": 15.0, "智能运维": "监控峰谷电耗与功率因数。"},
        {"系统": "仓储物流", "设备名称": "原辅料与成品立体仓", "规格": "重型货架+WMS", "数量": 1, "预估造价(万)": 65.0, "智能运维": "定期盘点，严格防潮防鼠防虫。"},
        {"系统": "建筑卫生", "设备名称": "男/女更衣室及卫生间", "规格": "GB14881 隔离级", "数量": 1, "预估造价(万)": 25.0, "智能运维": "风淋室互锁，强制洗手消毒门禁。"},
        {"系统": "生活后勤", "设备名称": "综合办公与员工宿舍楼", "规格": "含食堂", "数量": 1, "预估造价(万)": 180.0, "智能运维": "生活污水与工业废水分流排放。"}
    ]
    
    if category == "饼干/烘焙":
        oven_len = capacity * 5
        specific_bom = [
            {"系统": "核心工艺", "设备名称": "和面与成型机组", "规格": "自动化", "数量": 1, "预估造价(万)": 100.0, "智能运维": "彻底清洁辊筒防面筋发酸。"},
            {"系统": "热力工艺", "设备名称": "燃气隧道烘烤炉", "规格": f"长 {oven_len} 米", "数量": 1, "预估造价(万)": round(85.0 + oven_len*3.8, 1), "智能运维": "配置余热回收系统。"}
        ]
    elif category == "乳制品/液态奶":
        specific_bom = [
            {"系统": "前处理", "设备名称": "原奶收储与均质机组", "规格": "SUS316L", "数量": 1, "预估造价(万)": 85.0, "智能运维": "监测离心机转子平衡。"},
            {"系统": "杀菌工艺", "设备名称": "UHT 管式杀菌与无菌灌装", "规格": "百级净化", "数量": 1, "预估造价(万)": 180.0, "智能运维": "严格监控 137℃ 保温时间与 HEPA 滤网。"},
            {"系统": "卫生保障", "设备名称": "CIP 中央自动清洗站", "规格": "酸碱水三罐", "数量": 1, "预估造价(万)": 45.0, "智能运维": "标定电导率仪确保浓度。"}
        ]
    else: 
        specific_bom = [
            {"系统": "前处理", "设备名称": "解冻与真空斩拌滚揉线", "规格": "低温夹套冷水", "数量": 1, "预估造价(万)": 90.0, "智能运维": "监控真空泵油位。"},
            {"系统": "热加工", "设备名称": "全自动烟熏蒸煮炉", "规格": "含烟雾发生", "数量": 1, "预估造价(万)": 80.0, "智能运维": "排烟管道定期除焦油。"},
            {"系统": "冷链保障", "设备名称": "全厂速冻库与制冷机组", "规格": "-35℃ 螺杆式", "数量": 1, "预估造价(万)": 130.0, "智能运维": "化霜周期智能调节。"}
        ]
    return pd.DataFrame(base_bom + specific_bom)

# --- 3. CAD 引擎 (动态坐标 + 全厂基建布局) ---
def generate_ultimate_dxf(capacity, category):
    doc = ezdxf.new('R2010'); doc.styles.new('CHS', dxfattribs={'font': 'simsun.ttc'})
    if 'METRIC' not in doc.dimstyles:
        dimstyle = doc.dimstyles.new('METRIC'); dimstyle.dxf.dimasz = 100; dimstyle.dxf.dimtxt = 150
    if 'DASHED' not in doc.linetypes:
        doc.linetypes.new('DASHED', dxfattribs={'description': 'Dashed', 'pattern': [10.0, -5.0]})

    msp = doc.modelspace()
    
    # 增加灵活的 Y 和高度控制
    def draw_equip(name, x, length, y=0, h=1500, color=0):
        msp.add_lwpolyline([(x, y), (x+length, y), (x+length, y+h), (x, y+h), (x, y)], dxfattribs={'color': color})
        msp.add_text(name, dxfattribs={'height': 200, 'style': 'CHS'}).set_placement((x+length/2, y+h/2), align=TextEntityAlignment.MIDDLE_CENTER)

    # A. 动态绘制主工艺段
    end_x = 0
    if category == "饼干/烘焙":
        oven_len = capacity * 5000
        draw_equip("和面成型区", 0, 4000, color=1); draw_equip(f"隧道炉 ({capacity}T)", 4500, oven_len, color=4)
        end_x = 4500 + oven_len
    elif category == "乳制品/液态奶":
        draw_equip("原奶收储区", 0, 4000, color=3); draw_equip("UHT杀菌与CIP", 4500, 5000, color=1)
        draw_equip("百级无菌灌装室", 10000, 6000, color=5)
        end_x = 16000
    else: 
        draw_equip("解冻斩拌区", 0, 4000, color=1); draw_equip("烟熏蒸煮炉区", 4500, 5000, color=4)
        draw_equip("低温速冻包装区", 10000, 6000, color=5)
        end_x = 16000

    # B. 动态锚定：机械臂永远放置在主工艺的最尾端（码垛打包区）
    robot_x = end_x + 1500
    msp.add_circle((robot_x, 750), radius=1500, dxfattribs={'linetype': 'DASHED', 'color': 2})
    msp.add_text("码垛机械臂", dxfattribs={'height': 150, 'style': 'CHS', 'color': 2}).set_placement((robot_x, 2400), align=TextEntityAlignment.MIDDLE_CENTER)

    # C. 全厂基础设施布局 (在工艺区上方与下方排布)
    total_width = robot_x + 3000
    # 上方：物流与仓储
    draw_equip("原辅料立体仓储区", 0, total_width*0.4, y=3000, h=3000, color=6)
    draw_equip("成品发货库与月台", total_width*0.45, total_width*0.55, y=3000, h=3000, color=6)
    # 下方：人流与生活后勤
    draw_equip("卫生间/更衣室/风淋系统", 0, 4000, y=-3500, h=2000, color=2)
    draw_equip("综合办公与员工宿舍/食堂", 4500, total_width-4500, y=-3500, h=2000, color=7)

    # D. 贯穿全厂的隔离与动线
    path_y = 2000 # 产线与仓库之间的物流主干道
    msp.add_line((-1000, path_y), (total_width, path_y), dxfattribs={'linetype': 'DASHED', 'color': 3})
    msp.add_text("🤖 AGV 全厂物流自动接驳通道", dxfattribs={'height': 150, 'style': 'CHS', 'color': 3}).set_placement((total_width/2, path_y+300), align=TextEntityAlignment.MIDDLE_CENTER)

    buf = io.StringIO(); doc.write(buf); return buf.getvalue()

# --- 4. PDF 商业画册与原生 Word 说明书生成器 ---
def generate_pdf_proposal(project_name, category, capacity, cost, roi, physics, ai_report):
    labels = ['核心工艺线', '厂房土建与基建', '冷链/空调辅助', '不可预见费']
    if category == "饼干/烘焙": sizes = [cost * 0.40, cost * 0.35, cost * 0.15, cost * 0.10]
    elif category == "乳制品/液态奶": sizes = [cost * 0.45, cost * 0.25, cost * 0.20, cost * 0.10]
    else: sizes = [cost * 0.35, cost * 0.30, cost * 0.25, cost * 0.10]

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=['#ff9999','#66b3ff','#99ff99','#ffcc99'])
    ax.axis('equal')
    chart_filename = "temp_investment_chart.png"
    plt.savefig(chart_filename, bbox_inches='tight', transparent=True, dpi=150)
    plt.close(fig)

    pdf = FPDF(); pdf.add_page()
    font_path = "AlibabaPuHuiTi-3-115-Black.ttf" if os.path.exists("AlibabaPuHuiTi-3-115-Black.ttf") else ("C:\\Windows\\Fonts\\simhei.ttf" if os.path.exists("C:\\Windows\\Fonts\\simhei.ttf") else None)
    has_zh = False
    if font_path:
        try: pdf.add_font("ZH", style="", fname=font_path, uni=True); has_zh = True
        except: pass

    def safe_print(size, zh_text, en_text, align='L'):
        if has_zh: pdf.set_font("ZH", size=size); pdf.cell(0, 10, txt=zh_text, ln=1, align=align)
        else: pdf.set_font("Arial", size=size); pdf.cell(0, 10, txt=en_text, ln=1, align=align)

    if has_zh: pdf.set_font("ZH", size=18); pdf.cell(0, 15, txt="食品工厂总图规划与商业计划", ln=1, align='C')
    else: safe_print(18, "", "Factory Masterplan", 'C')
    
    safe_print(12, f"项目名称: {project_name} | 所属品类: {category} | 产能: {capacity}T", f"Category: {category}")
    safe_print(14, "一、 财务概算与全厂投资占比", "1. Financial Overview")
    safe_print(12, f"  • 预估总投资 : ￥{cost} 万 (含土建基建)", f"  - CAPEX: {cost}")
    
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
    """真·原生 Word 文档生成引擎"""
    doc = Document()
    doc.add_heading(f"{project_name} - 详细工程与工艺说明书", 0)
    doc.add_paragraph(f"项目类型：{category}    |    设计产能：{capacity} 吨/天")
    
    doc.add_heading('第一章：全厂总体规划与土建要求', level=1)
    doc.add_paragraph('本工厂布局严格遵守 GB 14881 食品安全国家标准，已全面规划以下区域：\n'
                      '1. 生产核心区：高/低洁净度物理隔离区。\n'
                      '2. 物流仓储区：原辅料立体库、成品发货月台。\n'
                      '3. 卫生配套区：风淋室、强制洗手池、独立排气卫生间。\n'
                      '4. 生活后勤区：办公楼、倒班宿舍及员工食堂（与生产区严格物理阻断）。')
    
    doc.add_heading('第二章：核心工艺与环保建议', level=1)
    doc.add_paragraph(ai_long_content.replace('*', '').replace('#', '')) # 简单清洗大模型的 markdown 符号
    
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# --- 5. Streamlit 界面构建 ---
st.set_page_config(page_title="总工程师基座", layout="wide")
s_price, c_price = fetch_market_data()

with st.sidebar:
    st.header("🗂️ 总图规划参数")
    proj_name = st.text_input("项目名称", "2026 绿建标准工厂")
    cat = st.selectbox("食品门类", ["饼干/烘焙", "乳制品/液态奶", "肉制品/切片"])
    cap = st.slider("日产吨数", 5, 50, 20)
    st.divider()
    st.success(f"钢价: ¥{s_price} | 碳价: ¥{c_price}")

physics = calculate_physics(cap, cat)
bom_df = generate_detailed_bom(cat, cap, physics)
total_cost = round(bom_df["预估造价(万)"].sum() * 1.15, 1)
roi = round(15.0 / (cap * 100 * 300 * 0.12 * 0.0035 + (cap*0.2*c_price)/10000 + 0.1), 1)

st.title("🏭 全要素基座：不仅是产线，更是生态")

c1, c2, c3, c4 = st.columns(4)
c1.metric("含基建总投", f"￥{total_cost} 万")
c2.metric("碳对冲 ROI", f"{roi} 年")
c3.metric("装机负荷", f"{physics['elec']} kW")
c4.metric("全厂清单项", len(bom_df))

st.subheader("📊 1. 商业路演准备 (面向资方)")
ai_content = "暂未生成。"
if st.button("✨ 启动路演级 AI 评估 (获取 PDF 画册素材)", use_container_width=True):
    with st.spinner('正在撰写商业企划文案...'):
        prompt = f"针对{proj_name}({cat})，产能{cap}吨。从全厂投资(含宿舍仓储)和ROI角度给出200字纯文本总结。"
        resp = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt}])
        st.session_state['ai_report'] = resp.choices[0].message.content
        st.success("商业评估完成！")

st.subheader("📝 2. 工程技术交付 (生成可编辑的 Word 说明书)")
if st.button("🚀 撰写详尽《全厂工程说明书》(约需 15 秒)", use_container_width=True):
    with st.spinner('DeepSeek 正在推演厂房、物流、工艺及生活区生态...'):
        long_prompt = f"为产能为{cap}吨/天的{cat}工厂撰写《全厂工艺说明书》。包含：1. 核心工艺；2. GB14881 全厂卫生分区(含厕所宿舍物流防交叉)；3. 环保与能耗。要求专业严谨，不少于600字，纯文本结构。"
        resp_long = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": long_prompt}])
        st.session_state['word_content'] = resp_long.choices[0].message.content
        st.success("Word 说明书核心内容撰写完毕！")

st.divider()
st.subheader("📥 3. 全厂交付物一键打包")
e1, e2, e3, e4 = st.columns(4)

excel_buf = io.BytesIO()
with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer: bom_df.to_excel(writer, index=False)
e1.download_button("📊 导出总包 BOM (Excel)", excel_buf.getvalue(), "Total_Factory_BOM.xlsx", use_container_width=True)

e2.download_button("📐 下载总平布置图 (DXF)", generate_ultimate_dxf(cap, cat), "Factory_Masterplan.dxf", use_container_width=True)

current_ai_report = st.session_state.get('ai_report', "请先运行 AI 评估。")
e3.download_button("📕 下载商业画册 (PDF)", generate_pdf_proposal(proj_name, cat, cap, total_cost, roi, physics, current_ai_report), "Business_Plan.pdf", mime="application/pdf", use_container_width=True)

# 核心更新：原生 Word 下载按钮
if 'word_content' in st.session_state:
    word_bytes = generate_word_manual(proj_name, cat, cap, st.session_state['word_content'])
    e4.download_button("📄 下载全厂说明书 (Word docx)", word_bytes, "Factory_Manual.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    e4.button("📄 请先生成说明书", disabled=True, use_container_width=True)
