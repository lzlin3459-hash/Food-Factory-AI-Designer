\# 🏭 Food Factory AI Designer | 食品工厂智能设计系统



> 🚀 \*\*领域架构师的跨界探索\*\*：当人工智能遇上传统食品工程与绿色金融。



本系统是一个专为中小型食品企业与工程咨询方打造的 \*\*工业设计 SaaS 雏形\*\*。通过融合 AI 大模型与物理推算，实现从“需求录入”到“LOD 400 施工级交付”的秒级跨越。



\## ✨ 核心亮点 (Core Features)

\* \*\*⚖️ 动态行情与 ESG 金融\*\*：接入 2026 实时不锈钢与碳交易（ETS）行情，动态推算绿色设备的 ROI。

\* \*\*🛡️ GB 14881 算法化合规\*\*：内置空间拓扑校验与动线交叉识别，从底层消灭生熟交叉与物理碰撞风险。

\* \*\*📐 施工级自动出图\*\*：根据产能与品类（烘焙/乳制品/肉制品），自动绘制带全尺寸标注、机械臂安全区与 AGV 路径的 CAD (DXF) 图纸。

\* \*\*📕 商业画册一键生成\*\*：集成 fpdf2 与 Matplotlib，静默渲染投资占比饼状图，导出纯正文排版的 PDF 商业计划书。



\## 🛠️ 技术栈 (Tech Stack)

\* \*\*Frontend\*\*: Streamlit

\* \*\*AI Brain\*\*: DeepSeek-V3/Chat (OpenAI API Compatible)

\* \*\*Engineering\*\*: ezdxf (CAD), Pandas (BOM), fpdf2 (PDF)



\## 💡 使用说明 (How to run locally)

1\. 安装依赖：`pip install -r requirements.txt`

2\. 放入字体：在根目录放入 `simhei.ttf` 或相应的中文字体文件以支持 PDF 渲染。

3\. 启动系统：`python -m streamlit run web\_app.py`

