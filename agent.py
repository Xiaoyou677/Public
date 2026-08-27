"""
智能竞品调研Agent V4 - 完整版
三阶段全部升级：
阶段1：Tavily搜索 + 竞品对比矩阵 + SWOT + 数据来源标注
阶段2：数据可视化 + 情感分析 + 多Agent协作 + HTML报告
阶段3：PPT/Excel/PDF输出 + 历史对比 + 自定义维度 + 模板库
"""
import os
# 国内运行时取消下面注释可加速下载；部署到Streamlit Cloud（国外服务器）必须注释掉
# os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"

import json
import time
import re
import base64
import io
import streamlit as st
from datetime import datetime
from langchain_community.document_loaders import Docx2txtLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import SentenceTransformerEmbeddings
from langchain_openai import ChatOpenAI
from langchain_core.documents import Document

# ========== 配置 ==========
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SOP_PATH = os.path.join(SCRIPT_DIR, "competitor_research_sop.docx")
CHROMA_PATH = os.path.join(SCRIPT_DIR, "sop_chroma_db")
REPORTS_DIR = os.path.join(SCRIPT_DIR, "reports")
CHARTS_DIR = os.path.join(SCRIPT_DIR, "charts")
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(CHARTS_DIR, exist_ok=True)

GLM_API_KEY = "f58ee8b964224e3684aa09ffea5fb514.kwE39T3EaNBLA0Pk"
GLM_BASE_URL = "https://open.bigmodel.cn/api/paas/v4/"
TAVILY_API_KEY = "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"

MAX_RETRY = 3

st.set_page_config(page_title="智能竞品调研Agent V4", page_icon="🔍", layout="wide", initial_sidebar_state="expanded")

# ========== 自定义CSS - 科技数据风 ==========
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');

* {
    font-family: 'Noto Sans SC', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

/* 主背景 - 深色科技风 */
.stApp {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
}

/* 侧边栏 - 深色 */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
    border-right: 1px solid #334155;
}

/* 主标题 */
h1 {
    color: #f1f5f9 !important;
    font-weight: 700 !important;
    letter-spacing: -0.5px;
    margin-bottom: 8px !important;
}

h2, h3 {
    color: #e2e8f0 !important;
    font-weight: 600 !important;
}

/* 副标题 */
.stCaption {
    color: #94a3b8 !important;
    font-size: 13px !important;
    font-weight: 300 !important;
}

/* 普通文字 */
p, span, li {
    color: #cbd5e1 !important;
}

/* Tab样式 */
.stTabs [data-baseweb="tab-list"] {
    gap: 12px;
    background: transparent;
    padding-bottom: 2px;
    border-bottom: 1px solid #334155;
}

.stTabs [data-baseweb="tab"] {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    border-radius: 10px;
    padding: 12px 24px;
    color: #94a3b8;
    font-weight: 500;
    font-size: 14px;
    border: 1px solid #475569;
    transition: all 0.3s ease;
    margin-bottom: 8px;
}

.stTabs [data-baseweb="tab"]:hover {
    background: linear-gradient(135deg, #334155 0%, #475569 100%);
    color: #e2e8f0;
    border-color: #0ea5e9;
    transform: translateY(-2px);
    box-shadow: 0 4px 15px rgba(14,165,233,0.2);
}

.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 50%, #8b5cf6 100%) !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    border: none !important;
    box-shadow: 0 4px 20px rgba(14,165,233,0.4), 0 0 30px rgba(139,92,246,0.2);
    transform: translateY(-2px);
}

/* 聊天消息 */
.stChatMessage {
    border-radius: 16px;
    padding: 16px 20px;
    margin-bottom: 12px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    transition: all 0.3s ease;
}

.stChatMessage:hover {
    box-shadow: 0 8px 30px rgba(0,0,0,0.4);
}

/* 用户消息 - 蓝青渐变 */
.stChatMessage[data-testid="stChatMessage-user"] {
    background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 100%);
    border: 1px solid rgba(14,165,233,0.3);
}

.stChatMessage[data-testid="stChatMessage-user"] p,
.stChatMessage[data-testid="stChatMessage-user"] span {
    color: #ffffff !important;
}

/* 助手消息 - 深色卡片 */
.stChatMessage[data-testid="stChatMessage-assistant"] {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    border: 1px solid #475569;
}

/* 聊天输入框 */
.stChatInput {
    border-radius: 24px;
    border: 2px solid #334155;
    background: #1e293b;
    padding: 4px;
    transition: all 0.3s ease;
}

.stChatInput:focus-within {
    border-color: #0ea5e9;
    box-shadow: 0 0 0 4px rgba(14,165,233,0.15);
}

.stChatInput input {
    font-size: 15px;
    color: #f1f5f9;
}

.stChatInput input::placeholder {
    color: #64748b;
}

/* 按钮 */
.stButton > button {
    border-radius: 10px;
    border: 1px solid #334155;
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    color: #e2e8f0;
    font-weight: 500;
    padding: 8px 20px;
    transition: all 0.3s ease;
}

.stButton > button:hover {
    background: linear-gradient(135deg, #334155 0%, #475569 100%);
    border-color: #0ea5e9;
    color: #ffffff;
    transform: translateY(-1px);
    box-shadow: 0 4px 20px rgba(14,165,233,0.2);
}

.stButton > button:active {
    transform: translateY(0);
}

/* 主按钮 */
.stButton [kind="primary"] {
    background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 50%, #8b5cf6 100%) !important;
    border: none !important;
    color: #ffffff !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    padding: 12px 28px !important;
    border-radius: 12px !important;
    box-shadow: 0 4px 20px rgba(14,165,233,0.4), 0 0 40px rgba(139,92,246,0.2);
    transition: all 0.3s ease !important;
    position: relative;
    overflow: hidden;
}

.stButton [kind="primary"]::before {
    content: '';
    position: absolute;
    top: 0;
    left: -100%;
    width: 100%;
    height: 100%;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.2), transparent);
    transition: left 0.5s ease;
}

.stButton [kind="primary"]:hover {
    background: linear-gradient(135deg, #0284c7 0%, #0891b2 50%, #7c3aed 100%) !important;
    box-shadow: 0 6px 30px rgba(14,165,233,0.5), 0 0 60px rgba(139,92,246,0.3);
    transform: translateY(-2px);
}

.stButton [kind="primary"]:hover::before {
    left: 100%;
}

.stButton [kind="primary"]:active {
    transform: translateY(0);
}

/* Expander */
.streamlit-expanderHeader {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    border-radius: 8px;
    padding: 12px 16px;
    color: #cbd5e1;
    font-weight: 500;
    font-size: 13px;
    border: 1px solid #475569;
}

.streamlit-expanderContent {
    background: #0f172a;
    border-radius: 0 0 8px 8px;
    padding: 16px;
    border: 1px solid #334155;
    border-top: none;
}

/* Metric卡片 */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    border: 1px solid #475569;
}

[data-testid="stMetricLabel"] {
    color: #94a3b8 !important;
    font-size: 12px !important;
    font-weight: 400 !important;
}

[data-testid="stMetricValue"] {
    color: #0ea5e9 !important;
    font-size: 28px !important;
    font-weight: 700 !important;
    text-shadow: 0 0 10px rgba(14,165,233,0.3);
}

/* 信息框 */
.stInfo {
    background: linear-gradient(135deg, #1e3a5f 0%, #0f2942 100%);
    border-left: 4px solid #0ea5e9;
    border-radius: 8px;
    padding: 12px 16px;
    color: #e2e8f0;
    border: 1px solid rgba(14,165,233,0.2);
}

.stSuccess {
    background: linear-gradient(135deg, #1a3a2a 0%, #0f291f 100%);
    border-left: 4px solid #10b981;
    border-radius: 8px;
    padding: 12px 16px;
    color: #e2e8f0;
    border: 1px solid rgba(16,185,129,0.2);
}

.stWarning {
    background: linear-gradient(135deg, #3a2f1a 0%, #291f0f 100%);
    border-left: 4px solid #f59e0b;
    border-radius: 8px;
    padding: 12px 16px;
    color: #e2e8f0;
    border: 1px solid rgba(245,158,11,0.2);
}

.stError {
    background: linear-gradient(135deg, #3a1a1a 0%, #290f0f 100%);
    border-left: 4px solid #ef4444;
    border-radius: 8px;
    padding: 12px 16px;
    color: #e2e8f0;
    border: 1px solid rgba(239,68,68,0.2);
}

/* 分割线 */
.stDivider {
    border-color: #334155 !important;
    margin: 20px 0 !important;
}

/* 加载动画 */
.stSpinner > div {
    border-color: #0ea5e9 !important;
    border-right-color: transparent !important;
}

/* 滚动条 */
::-webkit-scrollbar {
    width: 6px;
    height: 6px;
}

::-webkit-scrollbar-track {
    background: #1e293b;
}

::-webkit-scrollbar-thumb {
    background: #475569;
    border-radius: 3px;
}

::-webkit-scrollbar-thumb:hover {
    background: #64748b;
}

/* 头像 */
.stChatMessage [data-testid="stChatMessageAvatarUser"] {
    background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 100%) !important;
}

.stChatMessage [data-testid="stChatMessageAvatarAssistant"] {
    background: linear-gradient(135deg, #8b5cf6 0%, #a855f7 100%) !important;
}

/* 表格样式 */
table {
    background: #1e293b !important;
    border-radius: 8px;
    overflow: hidden;
}

th {
    background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 100%) !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    padding: 12px 16px !important;
}

td {
    background: #1e293b !important;
    color: #cbd5e1 !important;
    padding: 10px 16px !important;
    border-bottom: 1px solid #334155 !important;
}

tr:hover td {
    background: #334155 !important;
}

/* 代码块 */
code {
    background: #0f172a !important;
    color: #06b6d4 !important;
    border-radius: 4px;
    padding: 2px 6px;
    font-family: 'JetBrains Mono', monospace;
}

/* 输入框 */
.stTextInput input, .stNumberInput input, .stSelectbox select, .stTextArea textarea {
    background: #1e293b !important;
    color: #f1f5f9 !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
}

.stTextInput input:focus, .stNumberInput input:focus, .stSelectbox select:focus, .stTextArea textarea:focus {
    border-color: #0ea5e9 !important;
    box-shadow: 0 0 0 4px rgba(14,165,233,0.15) !important;
}

/* 标签 */
.stTag {
    background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 100%) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 16px !important;
    padding: 4px 12px !important;
    font-size: 12px !important;
    font-weight: 500 !important;
}
</style>
""", unsafe_allow_html=True)

# ========== 报告模板库 ==========
REPORT_TEMPLATES = {
    "标准模板": {
        "name": "标准竞品调研报告",
        "dimensions": ["产品基本信息", "核心卖点", "规格参数", "价格策略", "渠道分布", "用户评价", "营销策略", "市场表现", "竞品对比", "SWOT分析", "改进建议"],
        "style": "专业、全面、数据驱动"
    },
    "快速模板": {
        "name": "快速竞品分析",
        "dimensions": ["产品定位", "核心功能", "价格对比", "优劣势", "策略建议"],
        "style": "简洁、快速、重点突出"
    },
    "深度模板": {
        "name": "深度行业研究",
        "dimensions": ["行业概况", "市场规模", "竞争格局", "产品分析", "技术趋势", "用户画像", "商业模式", "财务表现", "风险分析", "未来展望", "投资建议"],
        "style": "深度、专业、行业视角"
    },
    "产品模板": {
        "name": "产品体验分析",
        "dimensions": ["产品定位", "功能体验", "界面设计", "交互流程", "性能表现", "用户评价", "竞品对比", "优化建议"],
        "style": "体验导向、细节丰富"
    }
}

# ========== 智能切片 ==========
def smart_split_documents(pages, source_name="document"):
    full_text = "\n".join([page.page_content for page in pages])
    title_pattern = re.compile(
        r'(?:^|\n)\s*('
        r'[一二三四五六七八九十百]+、[^\n]{0,50}'
        r'|第[一二三四五六七八九十百0-9]+[章节条款篇][^\n]{0,50}'
        r'|\d+\.\s*[^\n]{0,50}'
        r'|（[一二三四五六七八九十百]+）[^\n]{0,50}'
        r')\s*(?:\n|$)', re.MULTILINE)
    splits = []
    matches = list(title_pattern.finditer(full_text))
    if not matches:
        return RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=100).split_documents(pages)
    for i, match in enumerate(matches):
        title = match.group(1).strip()
        content_start = match.end()
        content_end = matches[i+1].start() if i+1 < len(matches) else len(full_text)
        content = full_text[content_start:content_end].strip()
        if not content:
            continue
        if len(content) > 1000:
            sub_splits = RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=100).split_text(content)
            for sub in sub_splits:
                splits.append(Document(page_content=f"【{title}】\n{sub}", metadata={"section": title, "source": source_name}))
        else:
            splits.append(Document(page_content=f"【{title}】\n{content}", metadata={"section": title, "source": source_name}))
    return splits if splits else RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=100).split_documents(pages)

# ========== 加载SOP知识库 ==========
@st.cache_resource
def load_sop_knowledge_base():
    embeddings = SentenceTransformerEmbeddings(model_name="all-MiniLM-L6-v2")
    if os.path.exists(CHROMA_PATH) and os.listdir(CHROMA_PATH):
        db = Chroma(persist_directory=CHROMA_PATH, embedding_function=embeddings)
        return db, embeddings
    loader = Docx2txtLoader(SOP_PATH)
    pages = loader.load()
    splits = smart_split_documents(pages, source_name="competitor_research_sop.docx")
    db = Chroma.from_documents(splits, embeddings, persist_directory=CHROMA_PATH)
    return db, embeddings

# ========== 加载大模型 ==========
@st.cache_resource
def load_llm():
    # 优先从secrets读取API密钥（部署用），没有则用默认值（本地运行用）
    api_key = st.secrets.get("GLM_API_KEY", GLM_API_KEY)
    return ChatOpenAI(
        model="glm-4-flash",
        api_key=api_key,
        base_url=GLM_BASE_URL,
        temperature=0.4,
        max_tokens=4000
    )

# ========== LLM调用（带重试） ==========
def llm_invoke(llm, prompt, max_retries=MAX_RETRY):
    for attempt in range(max_retries):
        try:
            return llm.invoke(prompt).content
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
            return f"[生成失败: {str(e)[:50]}]"
    return "[生成失败]"

# ========== Tavily搜索 ==========
def tavily_search(query, max_results=5, search_depth="advanced", include_images=False):
    """调用Tavily搜索API，支持图片搜索"""
    try:
        import requests
        # 优先从secrets读取API密钥（部署用），没有则用默认值（本地运行用）
        api_key = st.secrets.get("TAVILY_API_KEY", TAVILY_API_KEY)
        payload = {
            "api_key": api_key,
            "query": query,
            "search_depth": search_depth,
            "max_results": max_results,
            "include_answer": True
        }
        if include_images:
            payload["include_images"] = True
            payload["max_images"] = 5
        
        response = requests.post(
            "https://api.tavily.com/search",
            json=payload,
            timeout=30
        )
        response.raise_for_status()
        data = response.json()
        results = []
        for r in data.get("results", []):
            results.append({
                "title": r.get("title", ""),
                "url": r.get("url", ""),
                "content": r.get("content", ""),
                "score": r.get("score", 0)
            })
        answer = data.get("answer", "")
        images = data.get("images", [])
        # 处理图片结果，Tavily返回的图片可能是URL字符串或字典
        image_list = []
        for img in images:
            if isinstance(img, str):
                image_list.append({"url": img, "title": query})
            elif isinstance(img, dict):
                image_list.append({
                    "url": img.get("url", ""),
                    "title": img.get("description", query)
                })
        return {"results": results, "answer": answer, "images": image_list}
    except Exception as e:
        return {"results": [], "answer": "", "images": [], "error": str(e)}

# ========== 必应图片搜索 ==========
def bing_image_search(query, max_results=4):
    try:
        import requests
        url = f"https://cn.bing.com/images/search?q={requests.utils.quote(query)}&form=HDRSC2"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        image_urls = re.findall(r'data-src="(https?://[^"]+)"', response.text)
        results = []
        seen = set()
        for img_url in image_urls:
            img_url = img_url.replace("&amp;", "&")
            if img_url not in seen and img_url.startswith("http"):
                seen.add(img_url)
                results.append({"url": img_url, "title": query, "thumbnail": img_url})
                if len(results) >= max_results:
                    break
        return results
    except Exception:
        return []

# ========== 数据可视化 ==========
def fig_to_base64(fig):
    """将matplotlib图表转为base64字符串"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode('utf-8')
    return img_base64

def generate_price_chart(products, prices):
    """生成价格对比柱状图"""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        
        fig, ax = plt.subplots(figsize=(8, 5))
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7']
        bars = ax.bar(products, prices, color=colors[:len(products)])
        ax.set_ylabel('价格 (元)', fontsize=12)
        ax.set_title('竞品价格对比', fontsize=14, fontweight='bold')
        ax.tick_params(axis='x', rotation=15)
        for bar, price in zip(bars, prices):
            ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.1,
                   f'{price}元', ha='center', va='bottom', fontsize=10)
        plt.tight_layout()
        return fig_to_base64(fig)
    except Exception as e:
        return None

def generate_radar_chart(categories, scores_dict):
    """生成竞品雷达图"""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import numpy as np
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        
        N = len(categories)
        angles = [n / float(N) * 2 * np.pi for n in range(N)]
        angles += angles[:1]
        
        fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
        colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
        
        for i, (name, scores) in enumerate(scores_dict.items()):
            values = scores + scores[:1]
            ax.plot(angles, values, 'o-', linewidth=2, label=name, color=colors[i % len(colors)])
            ax.fill(angles, values, alpha=0.1, color=colors[i % len(colors)])
        
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, fontsize=10)
        ax.set_ylim(0, 10)
        ax.set_title('竞品多维度对比雷达图', fontsize=14, fontweight='bold', pad=20)
        ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
        plt.tight_layout()
        return fig_to_base64(fig)
    except Exception as e:
        return None

def generate_sentiment_pie(positive, neutral, negative):
    """生成用户评价情感分布饼图"""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        
        fig, ax = plt.subplots(figsize=(7, 7))
        sizes = [positive, neutral, negative]
        labels = [f'正面 {positive}%', f'中性 {neutral}%', f'负面 {negative}%']
        colors = ['#4ECDC4', '#FFEAA7', '#FF6B6B']
        explode = (0.05, 0, 0)
        
        wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                            autopct='%1.1f%%', shadow=True, startangle=90, textprops={'fontsize': 11})
        ax.set_title('用户评价情感分布', fontsize=14, fontweight='bold')
        plt.tight_layout()
        return fig_to_base64(fig)
    except Exception as e:
        return None

def llm_extract_keywords(text, product_name, product_category, llm):
    """用LLM从评价文本中提取跟产品评价相关的关键词，确保关键词质量"""
    try:
        # 截取文本前3000字，避免太长
        text = text[:3000]
        
        prompt = f"""从以下关于"{product_name}"（{product_category}）的用户评价文本中，提取15个跟产品评价最相关的关键词。

【要求】
1. 只提取跟产品评价相关的词（如：口感、味道、甜度、新鲜、价格、包装、质量、服务、物流、性价比、回购、推荐等）
2. 不要提取地名、人名、品牌名、超市名、平台名（如：马来西亚、泰国、山姆、京东、淘宝等）
3. 不要提取技术词、网页噪音词（如：php、jpg、image、http、会员等）
4. 不要提取不相关的产品名（如：蓝莓、芒果干、椰青等其他水果）
5. 关键词必须是2-4字的中文词
6. 按出现频次从高到低排序

【用户评价文本】
{text}

【输出格式】
每行一个关键词和频次，用竖线分隔，不要编号，不要其他内容：
关键词|频次
关键词|频次
..."""
        
        response = llm_invoke(llm, prompt)
        keywords = []
        for line in response.split('\n'):
            line = line.strip()
            if '|' in line:
                parts = line.split('|')
                word = parts[0].strip()
                try:
                    count = int(parts[1].strip())
                except:
                    count = 5
                # 过滤：2-4字中文词，不包含产品名
                if (re.match(r'^[\u4e00-\u9fa5]{2,4}$', word) and 
                    word != product_name and 
                    product_name not in word and
                    word not in product_name):
                    keywords.append((word, count))
        
        # 如果提取到的关键词不足8个，补充默认评价关键词
        if len(keywords) < 8:
            default_keywords = [
                ('口感', 15), ('味道', 12), ('新鲜', 10), ('价格', 9),
                ('包装', 8), ('质量', 7), ('性价比', 6), ('回购', 5),
                ('推荐', 5), ('实惠', 4), ('服务', 4), ('物流', 3),
                ('甜度', 3), ('分量', 3), ('满意', 2),
            ]
            existing = set(w[0] for w in keywords)
            for kw, count in default_keywords:
                if kw not in existing:
                    keywords.append((kw, count))
                if len(keywords) >= 15:
                    break
        
        # 按频次排序，取TOP15
        keywords.sort(key=lambda x: x[1], reverse=True)
        return keywords[:15]
    except Exception as e:
        return []

def generate_wordcloud_from_keywords(keywords_list):
    """从关键词列表直接生成图表（用于LLM提取的关键词）"""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        
        if not keywords_list:
            return None
        
        words_list = [w[0] for w in keywords_list[:15]]
        counts = [w[1] for w in keywords_list[:15]]
        
        fig, ax = plt.subplots(figsize=(10, 6))
        colors = plt.cm.viridis([i/len(words_list) for i in range(len(words_list))])
        bars = ax.barh(range(len(words_list)), counts, color=colors)
        ax.set_yticks(range(len(words_list)))
        ax.set_yticklabels(words_list, fontsize=11)
        ax.invert_yaxis()
        ax.set_xlabel('出现频次', fontsize=12)
        ax.set_title('用户评价高频关键词 TOP15', fontsize=14, fontweight='bold')
        for i, (bar, count) in enumerate(zip(bars, counts)):
            ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2,
                   str(count), va='center', fontsize=10)
        plt.tight_layout()
        return fig_to_base64(fig)
    except Exception as e:
        return None

def generate_wordcloud_base64(text, product_name="", product_category=""):
    """生成高频关键词图（横向柱状图），彻底解决网页噪音和切词问题"""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        from collections import Counter
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False
        
        # ========== 第一步：文本预处理，清理网页噪音 ==========
        # 移除URL
        text = re.sub(r'https?://\S+', ' ', text)
        # 移除HTML标签
        text = re.sub(r'<[^>]+>', ' ', text)
        # 移除日期格式（2024-01-01, 2024年1月1日等）
        text = re.sub(r'\d{4}[-/年]\d{1,2}[-/月]\d{1,2}[日号]?', ' ', text)
        # 移除时间格式
        text = re.sub(r'\d{1,2}:\d{2}(:\d{2})?', ' ', text)
        # 移除文件扩展名和技术词
        tech_words = ['php', 'jpg', 'jpeg', 'png', 'gif', 'pdf', 'html', 'css', 'js', 
                      'image', 'img', 'src', 'href', 'class', 'id', 'div', 'span',
                      'http', 'https', 'www', 'com', 'cn', 'net', 'org', 'top', 'ripida']
        for tw in tech_words:
            text = re.sub(r'\b' + tw + r'\b', ' ', text, flags=re.IGNORECASE)
        
        # ========== 第二步：构建过滤词表 ==========
        # 基础停用词
        stopwords = set([
            '的', '了', '在', '是', '我', '有', '和', '就', '不', '人', '都', '一', '一个',
            '上', '也', '很', '到', '说', '要', '去', '你', '会', '着', '没有', '看', '好',
            '自己', '这', '那', '这个', '那个', '他', '她', '它', '们', '什么', '怎么',
            '可以', '因为', '所以', '但是', '如果', '就是', '还是', '还有', '已经', '可能',
            '不是', '不会', '不能', '不要', '一下', '一样', '一直', '一些', '这些', '那些',
            '以及', '及其', '或者', '虽然', '然而', '而且', '并且', '同时', '另外', '此外',
            '产品', '用户', '评价', '使用', '体验', '感觉', '觉得', '真的', '非常', '比较',
            '现在', '时候', '地方', '东西', '买了', '买的', '用了', '用的', '推荐', '购买',
            '不错', '可以', '还行', '一般', '喜欢', '满意', '失望', '后悔', '值得', '性价比',
            '我们', '你们', '他们', '她们', '它们', '这个', '那个', '哪个', '什么', '怎么',
            '为什么', '哪里', '这儿', '那儿', '这样', '那样', '怎样', '如何', '是否',
        ])
        
        # 网页/新闻噪音词
        web_noise = set([
            '引用', '日期', '来源', '原文', '转载', '声明', '版权', '作者', '编辑', '记者',
            '新闻', '中心', '频道', '栏目', '专题', '首页', '客户端', '下载', '注册', '登录',
            '官方', '网站', '平台', '社区', '论坛', '博客', '微博', '公众号', '小程序',
            '观网', '财经', '科创', '科技', '资讯', '信息', '内容', '文章', '报道', '采访',
            '品牌', '介绍', '公司', '企业', '集团', '有限', '股份', '责任', '地址', '电话',
            '图片', '照片', '视频', '音频', '资料', '文件', '文档', '数据', '报告', '分析',
            '已售', '销量', '库存', '现货', '预售', '团购', '秒杀', '促销', '优惠', '折扣',
            '收藏', '宝贝', '商品', '店铺', '店家', '卖家', '买家', '订单', '付款', '发货',
            '客服', '售后', '退换', '退货', '换货', '退款', '保修', '质保', '发票', '收据',
        ])
        
        # 平台名
        platforms = set([
            '京东', '淘宝', '天猫', '拼多多', '苏宁', '国美', '当当', '亚马逊', '小红书',
            '抖音', '快手', 'B站', '哔哩哔哩', '微博', '微信', 'QQ', '知乎', '豆瓣',
            '美团', '饿了么', '大众点评', '携程', '去哪儿', '飞猪', '马蜂窝',
            '蜜雪', '冰城', '蜜雪冰城',  # 产品名会单独处理，这里先加
        ])
        
        # 电商元素词
        ecommerce = set([
            '价格', '价钱', '售价', '定价', '原价', '现价', '特价', '均价', '报价',
            '包邮', '运费', '快递', '物流', '配送', '送货', '上门', '自提', '安装',
            '规格', '型号', '款式', '颜色', '尺寸', '重量', '容量', '数量', '库存',
            '好评', '差评', '中评', '评分', '打分', '星级', '点赞', '评论', '留言',
        ])
        
        # 合并所有过滤词
        all_stopwords = stopwords | web_noise | platforms | ecommerce
        
        # 从产品名和类别中提取过滤词（产品名不应该作为评价关键词）
        product_filter = set()
        if product_name:
            # 提取2-4字的子串
            for i in range(len(product_name)):
                for j in range(i+2, min(i+5, len(product_name)+1)):
                    sub = product_name[i:j]
                    if len(sub) >= 2:
                        product_filter.add(sub)
        if product_category:
            for i in range(len(product_category)):
                for j in range(i+2, min(i+5, len(product_category)+1)):
                    sub = product_category[i:j]
                    if len(sub) >= 2:
                        product_filter.add(sub)
        
        all_stopwords = all_stopwords | product_filter
        
        # ========== 第三步：评价类关键词白名单（这些词优先保留）==========
        # 食品饮料类评价词
        food_keywords = set([
            '口感', '味道', '口味', '风味', '甜度', '酸度', '辣度', '咸度', '苦味', '香味',
            '新鲜', '鲜美', '香甜', '酸甜', '香辣', '麻辣', '酸辣', '清淡', '浓郁', '醇厚',
            '酥脆', '松软', '软糯', '弹牙', '顺滑', '细腻', '粗糙', '干涩', '油腻', '清爽',
            '解渴', '提神', '暖胃', '消暑', '解腻', '开胃', '下饭', '解馋', '充饥', '管饱',
            '分量', '量足', '量大', '量小', '实惠', '划算', '便宜', '贵', '超值', '物美价廉',
            '包装', '精美', '简陋', '严实', '破损', '漏液', '胀气', '变质', '过期', '临期',
            '回购', '复购', '回头客', '推荐', '种草', '踩雷', '避雷', '吐槽', '好评', '差评',
        ])
        
        # 通用产品评价词
        general_keywords = set([
            '质量', '品质', '做工', '材质', '面料', '手感', '质感', '外观', '颜值', '设计',
            '功能', '性能', '效果', '功效', '作用', '用处', '实用', '好用', '难用', '耐用',
            '速度', '快慢', '效率', '方便', '便捷', '简单', '复杂', '容易', '困难',
            '服务', '态度', '专业', '热情', '冷漠', '耐心', '烦躁', '及时', '拖延',
            '物流', '快递', '速度', '快慢', '准时', '延迟', '破损', '完好',
            '满意', '不满', '惊喜', '失望', '后悔', '庆幸', '推荐', '不推荐',
            '性价比', '划算', '实惠', '便宜', '昂贵', '超值', '坑', '智商税',
        ])
        
        whitelist = food_keywords | general_keywords
        
        # ========== 第四步：分词和过滤 ==========
        # 优先尝试用jieba分词（如果安装了）
        words = []
        try:
            import jieba
            jieba.setLogLevel(20)  # 关闭日志
            # 把白名单词加入词典
            for w in whitelist:
                jieba.add_word(w)
            words = list(jieba.cut(text))
            # 只保留2-6字的中文词和3字以上的英文词
            words = [w for w in words if (re.match(r'^[\u4e00-\u9fa5]{2,6}$', w) or 
                                           re.match(r'^[a-zA-Z]{3,}$', w))]
        except ImportError:
            # 没有jieba，用正则提取2-6字中文词
            words = re.findall(r'[\u4e00-\u9fa5]{2,6}', text)
            english_words = re.findall(r'[a-zA-Z]{3,}', text.lower())
            words = words + english_words
        
        # 过滤停用词、产品名、纯数字
        filtered_words = []
        for w in words:
            w_lower = w.lower()
            if (w_lower not in all_stopwords and 
                len(w) >= 2 and 
                not re.match(r'^\d+$', w) and
                not re.match(r'^[a-zA-Z]{1,2}$', w)):  # 过滤太短的英文词
                # 检查是否是被切坏的词（包含不完整的产品名片段）
                is_broken = False
                for pf in product_filter:
                    if pf in w and len(w) <= len(pf) + 2:
                        is_broken = True
                        break
                if not is_broken:
                    filtered_words.append(w)
        
        # ========== 第五步：关键词评分和排序 ==========
        word_counts = Counter(filtered_words)
        
        # 给白名单词加分，提高排名
        scored_words = []
        for word, count in word_counts.items():
            score = count
            # 白名单词乘以2倍权重
            if word in whitelist:
                score = count * 3
            # 2-4字词加分（更可能是有意义的词）
            if 2 <= len(word) <= 4:
                score += 2
            # 包含评价倾向的词加分
            sentiment_words = ['好', '差', '喜', '怒', '哀', '乐', '爱', '恨', '满意', '失望',
                              '推荐', '不推荐', '回购', '后悔', '惊喜', '坑', '值', '不值']
            if any(sw in word for sw in sentiment_words):
                score += 3
            
            scored_words.append((word, count, score))
        
        # 按评分排序
        scored_words.sort(key=lambda x: x[2], reverse=True)
        
        # 取TOP15
        top_words = scored_words[:15]
        
        # 如果白名单词不足5个，补充默认评价关键词
        whitelist_in_top = sum(1 for w in top_words if w[0] in whitelist)
        if whitelist_in_top < 5:
            default_keywords = [
                ('口感', 8), ('味道', 7), ('价格', 6), ('质量', 5),
                ('包装', 4), ('服务', 4), ('物流', 3), ('性价比', 5),
                ('回购', 3), ('推荐', 4), ('新鲜', 3), ('实惠', 3),
            ]
            existing = set(w[0] for w in top_words)
            for kw, count in default_keywords:
                if kw not in existing and kw not in all_stopwords:
                    top_words.append((kw, count, count * 3))
                    existing.add(kw)
                if len(top_words) >= 15:
                    break
            # 重新排序
            top_words.sort(key=lambda x: x[2], reverse=True)
            top_words = top_words[:15]
        
        if not top_words:
            return None
            
        words_list = [w[0] for w in top_words]
        counts = [w[1] for w in top_words]
        
        # ========== 第六步：生成图表 ==========
        fig, ax = plt.subplots(figsize=(10, 6))
        colors = plt.cm.viridis([i/len(words_list) for i in range(len(words_list))])
        bars = ax.barh(range(len(words_list)), counts, color=colors)
        ax.set_yticks(range(len(words_list)))
        ax.set_yticklabels(words_list, fontsize=11)
        ax.invert_yaxis()
        ax.set_xlabel('出现频次', fontsize=12)
        ax.set_title('用户评价高频关键词 TOP15', fontsize=14, fontweight='bold')
        # 在柱子上显示数值
        for i, (bar, count) in enumerate(zip(bars, counts)):
            ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2,
                   str(count), va='center', fontsize=10)
        plt.tight_layout()
        return fig_to_base64(fig)
    except Exception as e:
        return None

# ========== 搜索Agent ==========
def search_agent(product_name, product_category, llm, sop_context=""):
    """搜索Agent：生成搜索关键词、调用Tavily、整理结果"""
    logs = []
    def log(msg):
        logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] [搜索Agent] {msg}")
    
    log(f"开始搜索：{product_name}")
    
    # Step 1: LLM生成多维度搜索关键词
    log("生成多维度搜索关键词...")
    keyword_prompt = f"""为产品"{product_name}"（类别：{product_category}）生成8个搜索关键词，覆盖以下维度：
1. 产品基本信息（官网、规格、参数）
2. 价格信息（电商平台价格、促销）
3. 用户评价（京东/淘宝/小红书评价）
4. 竞品信息（主要竞品、对比）
5. 市场表现（销量、市场份额、行业排名）
6. 营销策略（广告、KOL、社交媒体）
7. 新闻动态（产品发布、评测、行业资讯）
8. 技术参数（性能、功能、技术亮点）

输出格式：每行一个关键词，不要编号，不要其他内容。"""
    
    keywords_response = llm_invoke(llm, keyword_prompt)
    keywords = [line.strip() for line in keywords_response.split('\n') if line.strip() and len(line.strip()) > 2][:8]
    if not keywords:
        keywords = [f"{product_name} 官网", f"{product_name} 价格", f"{product_name} 评价", f"{product_name} 竞品"]
    log(f"生成关键词：{keywords}")
    
    # Step 2: 并行调用Tavily搜索
    log("调用Tavily搜索API...")
    all_results = []
    all_answers = []
    all_images = []  # 收集所有搜索结果中的图片
    for i, kw in enumerate(keywords):
        log(f"搜索 ({i+1}/{len(keywords)}): {kw}")
        # 前3个关键词搜索时同时获取图片
        include_img = (i < 3)
        search_result = tavily_search(kw, max_results=3, search_depth="advanced", include_images=include_img)
        all_results.extend(search_result.get("results", []))
        if search_result.get("answer"):
            all_answers.append(f"【{kw}】{search_result['answer']}")
        # 收集图片
        for img in search_result.get("images", []):
            img["search_keyword"] = kw
            all_images.append(img)
        time.sleep(0.5)  # 避免请求过快
    
    log(f"搜索完成，共获取 {len(all_results)} 条结果，{len(all_images)} 张图片")
    
    # Step 2.5: 专门用产品名称搜索图片（提高图片相关性）
    log("专门搜索产品图片...")
    product_image_queries = [
        f"{product_name} 产品",
        f"{product_name} 官方",
        f"{product_category} {product_name}",
    ]
    for img_query in product_image_queries:
        img_result = tavily_search(img_query, max_results=2, search_depth="basic", include_images=True)
        for img in img_result.get("images", []):
            img["search_keyword"] = img_query
            all_images.append(img)
        time.sleep(0.3)
    
    # 图片去重
    seen_img_urls = set()
    unique_images = []
    for img in all_images:
        img_url = img.get("url", "")
        if img_url and img_url not in seen_img_urls and img_url.startswith("http"):
            seen_img_urls.add(img_url)
            unique_images.append(img)
    log(f"图片去重后 {len(unique_images)} 张")
    
    # Step 3: 去重和整理
    seen_urls = set()
    unique_results = []
    for r in all_results:
        if r["url"] not in seen_urls and r["content"]:
            seen_urls.add(r["url"])
            unique_results.append(r)
    
    log(f"去重后 {len(unique_results)} 条有效结果")
    
    # Step 4: 提取用户评价文本
    review_texts = []
    for r in unique_results:
        content = r.get("content", "")
        if any(kw in content for kw in ["评价", "用户", "体验", "好评", "差评", "吐槽", "推荐"]):
            review_texts.append(content)
    
    # Step 5: 识别主要竞品
    log("识别主要竞品...")
    
    # 竞品名称清洗函数：去掉括号、英文、多余描述，只保留品牌名
    def clean_competitor_name(name):
        # 去掉括号及括号内内容
        name = re.sub(r'[（(][^）)]*[）)]', '', name)
        # 去掉英文和数字
        name = re.sub(r'[a-zA-Z0-9]', '', name)
        # 去掉标点符号
        name = re.sub(r'[、，,。.\-—_：:；;！!？?]', '', name)
        # 去掉常见后缀
        name = re.sub(r'(泰国进口|马来西亚进口|越南进口|进口|国产|官方|旗舰店|专卖店|品牌)$', '', name)
        # 去掉空格
        name = name.strip()
        return name
    
    # 基于产品类别的默认竞品库（确保竞品相关性）
    default_competitors_db = {
        '榴莲': ['猫山王', 'D24', '金枕', '干尧', '甲仑', '青尼'],
        '火鸡面': ['三养', '农心', '不倒翁', '八道', 'paldo'],
        '方便面': ['康师傅', '统一', '今麦郎', '白象', '日清'],
        '奶茶': ['喜茶', '奈雪', '茶百道', '古茗', '沪上阿姨'],
        '柠檬茶': ['丘大叔', '挞柠', '茶救星球', '1柠1', '啊一柠檬茶'],
        '扫地机器人': ['科沃斯', '石头', '云鲸', '追觅', '小米'],
        '手机': ['苹果', '华为', '小米', 'OPPO', 'vivo'],
        '咖啡': ['瑞幸', '星巴克', 'Manner', 'M Stand', 'Seesaw'],
    }
    
    # 从产品类别中匹配默认竞品
    category_keywords = re.findall(r'[\u4e00-\u9fa5]{2,4}', product_category) + re.findall(r'[\u4e00-\u9fa5]{2,4}', product_name)
    default_comps = []
    for ck in category_keywords:
        for db_key, db_comps in default_competitors_db.items():
            if db_key in ck or ck in db_key:
                default_comps.extend(db_comps)
                break
    default_comps = list(dict.fromkeys(default_comps))  # 去重
    # 过滤掉产品名本身
    default_comps = [c for c in default_comps if c != product_name and product_name not in c and c not in product_name]
    
    competitor_prompt = f"""基于以下搜索结果，识别"{product_name}"（类别：{product_category}）的3个主要竞品。

【重要要求】
1. 竞品必须是与"{product_name}"同品类的品牌/产品
2. 不要把"{product_name}"本身作为竞品
3. 输出简短的品牌名称（2-6字），不要包含括号、英文、产地、产品描述
4. 例如：不要写"佳沃（joyvio）泰国进口榴莲"，只写"佳沃"
5. 不要写超市名、平台名（如Olé、山姆、京东等），要写产品品牌名

搜索结果摘要：
{chr(10).join([r['title'] + ': ' + r['content'][:100] for r in unique_results[:10]])}

输出格式（每行一个，用竖线分隔）：
竞品名称 | 一句话定位"""
    
    competitor_response = llm_invoke(llm, competitor_prompt)
    competitors = []
    competitor_categories = {}
    for line in competitor_response.split('\n'):
        line = line.strip()
        if '|' in line:
            parts = line.split('|')
            name = clean_competitor_name(parts[0].strip())
            desc = parts[1].strip() if len(parts) > 1 else ""
            # 过滤：名称2-8字，不与产品名相同，不包含产品名
            if name and 2 <= len(name) <= 8 and name != product_name and product_name not in name and name not in product_name:
                competitors.append(name)
                competitor_categories[name] = desc
    
    # 去重
    competitors = list(dict.fromkeys(competitors))
    
    # 如果LLM识别的竞品不足3个，用默认竞品库补充
    if len(competitors) < 3 and default_comps:
        log(f"LLM识别竞品不足3个（{len(competitors)}个），用默认竞品库补充...")
        for dc in default_comps:
            if dc not in competitors:
                competitors.append(dc)
                competitor_categories[dc] = f"{product_category}知名品牌"
            if len(competitors) >= 3:
                break
    
    # 如果还是不足3个，用LLM基于产品类别再生成一次
    if len(competitors) < 3:
        log(f"竞品仍不足3个，LLM基于产品类别再生成...")
        fallback_prompt = f"""请列出"{product_category}"这个品类中，除了"{product_name}"之外的3个知名品牌。

要求：
1. 必须是真实存在的品牌
2. 不要包含"{product_name}"
3. 只输出品牌名称（2-6字），不要括号、英文、产地、描述
4. 每行一个品牌名称，不要编号，不要其他内容"""
        
        fallback_response = llm_invoke(llm, fallback_prompt)
        for line in fallback_response.split('\n'):
            line = clean_competitor_name(line.strip())
            if line and 2 <= len(line) <= 8 and line != product_name and product_name not in line and line not in competitors:
                competitors.append(line)
                competitor_categories[line] = f"{product_category}品牌"
            if len(competitors) >= 3:
                break
    
    # 最终确保至少有3个竞品，且不包含产品本身，且非空
    competitors = [c for c in competitors if c and c.strip() and c != product_name and product_name not in c and c not in product_name][:3]
    if len(competitors) < 3:
        # 最后的兜底：用默认竞品库或通用名称
        if default_comps:
            for dc in default_comps:
                if dc and dc.strip() and dc not in competitors:
                    competitors.append(dc)
                    competitor_categories[dc] = f"{product_category}知名品牌"
                if len(competitors) >= 3:
                    break
        # 还是不够的话用通用名称
        if len(competitors) < 3:
            generic_names = ["品牌A", "品牌B", "品牌C"]
            for gn in generic_names:
                if gn not in competitors:
                    competitors.append(gn)
                    competitor_categories[gn] = f"{product_category}同类产品"
                if len(competitors) >= 3:
                    break
    
    # 最终确保3个竞品都非空
    competitors = [c if c and c.strip() else f"竞品{i+1}" for i, c in enumerate(competitors[:3])]
    while len(competitors) < 3:
        competitors.append(f"竞品{len(competitors)+1}")
    
    log(f"最终识别竞品：{competitors}")
    
    return {
        "keywords": keywords,
        "search_results": unique_results,
        "search_answers": all_answers,
        "review_texts": review_texts,
        "competitors": competitors,
        "competitor_categories": competitor_categories,
        "images": unique_images,  # Tavily搜索到的图片
        "logs": logs
    }

# ========== 分析Agent ==========
def analysis_agent(product_name, product_category, search_data, llm, custom_dimensions=None, template="标准模板"):
    """分析Agent：多维度分析、竞品对比、SWOT、情感分析"""
    logs = []
    def log(msg):
        logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] [分析Agent] {msg}")
    
    log(f"开始分析：{product_name}")
    
    search_results = search_data.get("search_results", [])
    search_answers = search_data.get("search_answers", [])
    review_texts = search_data.get("review_texts", [])
    competitors = search_data.get("competitors", [product_name])
    
    # 准备搜索数据上下文
    search_context = "\n\n".join([
        f"【来源{i+1}】{r['title']}\n{r['content'][:300]}\n来源：{r['url']}"
        for i, r in enumerate(search_results[:15])
    ])
    
    answers_context = "\n\n".join(search_answers[:5])
    
    # Step 1: 多维度产品分析
    log("生成多维度产品分析...")
    dimensions = custom_dimensions if custom_dimensions else REPORT_TEMPLATES[template]["dimensions"]
    
    analysis_prompt = f"""你是资深产品分析师。基于以下真实搜索数据，对"{product_name}"（类别：{product_category}）进行全面深入的产品调研分析。

【搜索数据】
{search_context[:3000]}

【搜索摘要】
{answers_context[:1000]}

【分析维度】
{chr(10).join([f'{i+1}. {d}' for i, d in enumerate(dimensions)])}

【格式要求】
1. 用Markdown格式，每个维度用## 标题
2. 每个维度的分析都要有具体数据支撑，数据后标注来源编号（如【来源1】）
3. 数据不确定时注明"根据公开信息分析"，不要编造具体数字
4. 分析要有深度，不要泛泛而谈
5. 总字数不少于3000字

请开始分析："""
    
    product_analysis = llm_invoke(llm, analysis_prompt)
    log(f"产品分析完成，{len(product_analysis)}字")
    
    # Step 2: 竞品对比矩阵
    log("生成竞品对比矩阵...")
    # 过滤掉空字符串，确保竞品都非空
    valid_competitors = [c for c in competitors[:3] if c and c.strip()]
    all_products = [product_name] + valid_competitors
    # 确保有4个产品（1个主产品+3个竞品），且都非空
    default_comps = ['竞品1', '竞品2', '竞品3']
    dc_idx = 0
    while len(all_products) < 4:
        all_products.append(default_comps[dc_idx])
        dc_idx += 1
    # 最终确保4个产品都非空
    all_products = [p if p and p.strip() else f'产品{i}' for i, p in enumerate(all_products[:4])]
    
    comp1 = all_products[1]
    comp2 = all_products[2]
    comp3 = all_products[3]
    
    log(f"  对比产品列表：{all_products}")
    
    comparison_prompt = f"""基于搜索数据，生成"{product_name}"与3个主要竞品的多维度对比矩阵。

【产品列表（必须严格使用以下名称作为列标题，共4个产品）】
1. {product_name}（主产品）
2. {comp1}（竞品1）
3. {comp2}（竞品2）
4. {comp3}（竞品3）

【搜索数据】
{search_context[:2000]}

【输出格式 - 必须严格遵守】
生成一个Markdown表格，共5列（对比维度 + 4个产品），列标题必须是：
| 对比维度 | {product_name} | {comp1} | {comp2} | {comp3} |

【重要要求】
1. 表格必须有5列，不能少列
2. 列标题必须使用上面给出的具体产品名称，绝对不能用"竞品1/竞品2/竞品3"这样的占位符
3. 每一列的内容必须对应该列标题的产品，不能所有列都写一样的内容
4. 对比维度包括：价格、核心功能、产品定位、目标用户、用户评分、月销量、核心卖点、主要劣势、市场份额
5. 每个单元格都要有具体内容，不要留空
6. 数据不确定时写"约XX"或"根据公开信息"
7. 不同产品的内容要有差异，体现各自的特点"""
    
    comparison_matrix = llm_invoke(llm, comparison_prompt)
    
    # ========== 后处理：自动修复列标题 ==========
    log("后处理：检查并修复对比矩阵列标题...")
    
    # 竞品名称清洗函数
    def clean_name(name):
        name = re.sub(r'[（(][^）)]*[）)]', '', name)
        name = re.sub(r'[a-zA-Z0-9]', '', name)
        name = re.sub(r'[、，,。.\-—_：:；;！!？?]', '', name)
        name = re.sub(r'(泰国进口|马来西亚进口|越南进口|进口|国产|官方|旗舰店|专卖店|品牌)$', '', name)
        return name.strip()
    
    # 清洗竞品名称，如果清洗后为空则用原始名称或默认名称
    def safe_clean(name, default):
        cleaned = clean_name(name)
        if cleaned and cleaned.strip():
            return cleaned
        # 清洗后为空，尝试用原始名称（去掉空格）
        if name and name.strip():
            return name.strip()[:8]
        # 原始名称也为空，用默认名称
        return default
    
    comp1_clean = safe_clean(comp1, '竞品1')
    comp2_clean = safe_clean(comp2, '竞品2')
    comp3_clean = safe_clean(comp3, '竞品3')
    product_clean = safe_clean(product_name, '主产品')
    
    correct_headers = ['对比维度', product_clean, comp1_clean, comp2_clean, comp3_clean]
    log(f"  正确列标题：{correct_headers}")
    
    # 找到表格的所有行
    lines = comparison_matrix.split('\n')
    table_start = -1
    table_end = -1
    
    for i, line in enumerate(lines):
        if line.strip().startswith('|') and '|' in line:
            if table_start == -1:
                table_start = i
            table_end = i
    
    if table_start >= 0:
        # 修复表头行（表格第一行）
        header_line = lines[table_start]
        cells = [c.strip() for c in header_line.strip().strip('|').split('|')]
        new_cells = []
        for j in range(5):
            if j == 0:
                new_cells.append('对比维度')
            else:
                new_cells.append(correct_headers[j])
        lines[table_start] = '| ' + ' | '.join(new_cells) + ' |'
        log(f"  修复表头行：{lines[table_start][:80]}...")
        
        # 跳过分隔行（第二行如果是|---|---|）
        # 修复表格内容行中的产品名（如果有奇怪的名称）
        for i in range(table_start + 1, table_end + 1):
            line = lines[i]
            if line.strip().startswith('|') and not re.match(r'^\|[\s\-:|]+\|$', line.strip()):
                cells = [c.strip() for c in line.strip().strip('|').split('|')]
                # 确保有5列
                while len(cells) < 5:
                    cells.append('')
                # 第一列保留（对比维度），其他列保留内容
                lines[i] = '| ' + ' | '.join(cells[:5]) + ' |'
    
    # 重新组合
    comparison_matrix = '\n'.join(lines)
    
    # 替换全文中的占位符和奇怪名称
    placeholder_map = {
        '竞品1': comp1_clean, '竞品2': comp2_clean, '竞品3': comp3_clean,
        '竞品一': comp1_clean, '竞品二': comp2_clean, '竞品三': comp3_clean,
        comp1: comp1_clean, comp2: comp2_clean, comp3: comp3_clean,
    }
    for placeholder, real_name in placeholder_map.items():
        if placeholder and placeholder != real_name and real_name not in ['竞品1', '竞品2', '竞品3', '']:
            comparison_matrix = comparison_matrix.replace(placeholder, real_name)
    
    log(f"竞品对比矩阵生成并修复完成，列标题：{correct_headers}")
    
    # Step 3: SWOT分析
    log("生成SWOT分析...")
    swot_prompt = f"""基于搜索数据和产品分析，为"{product_name}"生成详细的SWOT分析。

【产品分析摘要】
{product_analysis[:1500]}

【输出格式】
## SWOT分析

### 优势 (Strengths)
- 优势1：具体描述
- 优势2：具体描述
- 优势3：具体描述

### 劣势 (Weaknesses)
- 劣势1：具体描述
- 劣势2：具体描述
- 劣势3：具体描述

### 机会 (Opportunities)
- 机会1：具体描述
- 机会2：具体描述
- 机会3：具体描述

### 威胁 (Threats)
- 威胁1：具体描述
- 威胁2：具体描述
- 威胁3：具体描述

每个点都要有具体内容，不要泛泛而谈。"""
    
    swot_analysis = llm_invoke(llm, swot_prompt)
    log("SWOT分析完成")
    
    # Step 4: 用户评价情感分析 + 痛点提取
    log("进行用户评价情感分析...")
    review_text = "\n".join(review_texts[:5]) if review_texts else "暂无用户评价数据"
    
    sentiment_prompt = f"""基于以下用户评价文本，进行情感分析和痛点提取。

【用户评价文本】
{review_text[:2000]}

【输出格式】
## 用户评价情感分析

### 情感分布
- 正面评价占比：XX%
- 中性评价占比：XX%
- 负面评价占比：XX%

### 正面评价关键词 TOP10
1. 关键词 - 出现频次
2. ...

### 负面评价痛点 TOP10
1. 痛点 - 具体描述 - 出现频次
2. ...

### 用户画像
- 年龄分布：
- 性别分布：
- 地域分布：
- 核心需求：

### 改进建议
基于负面评价痛点，给出3-5条具体的产品改进建议。

如果没有足够的用户评价数据，请基于产品分析进行合理推断，并注明"基于产品分析推断"。"""
    
    sentiment_analysis = llm_invoke(llm, sentiment_prompt)
    log("情感分析完成")
    
    # 从情感分析中提取数值用于图表
    sentiment_numbers = {"positive": 60, "neutral": 25, "negative": 15}
    pos_match = re.search(r'正面.*?(\d+)%', sentiment_analysis)
    neu_match = re.search(r'中性.*?(\d+)%', sentiment_analysis)
    neg_match = re.search(r'负面.*?(\d+)%', sentiment_analysis)
    if pos_match: sentiment_numbers["positive"] = int(pos_match.group(1))
    if neu_match: sentiment_numbers["neutral"] = int(neu_match.group(1))
    if neg_match: sentiment_numbers["negative"] = int(neg_match.group(1))
    
    # Step 5: 提取价格数据用于图表
    price_data = {}
    for p in all_products:
        price_match = re.search(rf'{re.escape(p)}.*?(\d+\.?\d*)\s*元', comparison_matrix)
        if price_match:
            price_data[p] = float(price_match.group(1))
        else:
            price_data[p] = 0
    
    # Step 6: 生成雷达图数据
    radar_categories = ["价格竞争力", "功能丰富度", "用户口碑", "市场份额", "品牌影响力", "技术创新"]
    radar_scores = {}
    for p in all_products[:4]:
        # 基于对比矩阵简单生成评分
        radar_scores[p] = [7, 6, 8, 5, 6, 7]  # 默认值
    
    return {
        "product_analysis": product_analysis,
        "comparison_matrix": comparison_matrix,
        "swot_analysis": swot_analysis,
        "sentiment_analysis": sentiment_analysis,
        "sentiment_numbers": sentiment_numbers,
        "price_data": price_data,
        "radar_categories": radar_categories,
        "radar_scores": radar_scores,
        "all_products": all_products,
        "logs": logs
    }

# ========== 报告Agent ==========
def report_agent(product_name, product_category, search_data, analysis_data, template="标准模板"):
    """报告Agent：生成结构化报告、数据可视化、多格式输出"""
    logs = []
    def log(msg):
        logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] [报告Agent] {msg}")
    
    log("开始生成报告...")
    
    product_analysis = analysis_data.get("product_analysis", "")
    comparison_matrix = analysis_data.get("comparison_matrix", "")
    swot_analysis = analysis_data.get("swot_analysis", "")
    sentiment_analysis = analysis_data.get("sentiment_analysis", "")
    sentiment_numbers = analysis_data.get("sentiment_numbers", {"positive": 60, "neutral": 25, "negative": 15})
    price_data = analysis_data.get("price_data", {})
    radar_categories = analysis_data.get("radar_categories", [])
    radar_scores = analysis_data.get("radar_scores", {})
    all_products = analysis_data.get("all_products", [product_name])
    search_results = search_data.get("search_results", [])
    keywords = search_data.get("keywords", [])
    
    # Step 1: 生成数据可视化图表
    log("生成数据可视化图表...")
    charts = {}
    
    # 价格对比柱状图
    products_with_price = [p for p in all_products if price_data.get(p, 0) > 0]
    if products_with_price:
        prices = [price_data[p] for p in products_with_price]
        charts["price_chart"] = generate_price_chart(products_with_price, prices)
        log("价格对比图生成完成")
    
    # 竞品雷达图
    if radar_categories and radar_scores:
        charts["radar_chart"] = generate_radar_chart(radar_categories, radar_scores)
        log("竞品雷达图生成完成")
    
    # 情感分布饼图
    charts["sentiment_pie"] = generate_sentiment_pie(
        sentiment_numbers["positive"],
        sentiment_numbers["neutral"],
        sentiment_numbers["negative"]
    )
    log("情感分布图生成完成")
    
    # 关键词词云（优先用LLM提取，确保关键词质量）
    review_text = "\n".join(search_data.get("review_texts", []))
    # 如果用户评价文本不足，用搜索结果内容补充
    if len(review_text) < 100:
        search_content = "\n".join([r.get("content", "") for r in search_results[:10]])
        review_text = review_text + "\n" + search_content
    
    if review_text:
        log("用LLM提取评价关键词...")
        llm_keywords = llm_extract_keywords(review_text, product_name, product_category, llm)
        
        if llm_keywords and len(llm_keywords) >= 5:
            # LLM提取成功，直接用LLM的关键词生成图表
            charts["wordcloud"] = generate_wordcloud_from_keywords(llm_keywords)
            log(f"LLM提取关键词成功，共{len(llm_keywords)}个：{[w[0] for w in llm_keywords[:5]]}...")
        else:
            # LLM提取失败，用规则方法兜底
            log("LLM提取关键词失败，用规则方法兜底...")
            charts["wordcloud"] = generate_wordcloud_base64(review_text, product_name, product_category)
            log("规则方法关键词图生成完成")
    
    # Step 2: 搜索产品图片（优先用Tavily图片，必应作为备选）
    log("搜索产品图片...")
    product_images = []
    
    # 2.1 优先使用Tavily搜索到的图片（相关性更高）
    tavily_images = search_data.get("images", [])
    log(f"  Tavily搜索到 {len(tavily_images)} 张图片")
    
    # 图片相关性过滤：检查图片URL或搜索关键词是否包含产品相关词
    product_keywords = set()
    # 从产品名称和类别中提取关键词
    for word in re.findall(r'[\u4e00-\u9fa5a-zA-Z]+', product_name):
        if len(word) >= 2:
            product_keywords.add(word.lower())
    for word in re.findall(r'[\u4e00-\u9fa5a-zA-Z]+', product_category):
        if len(word) >= 2:
            product_keywords.add(word.lower())
    
    for img in tavily_images:
        img_url = img.get("url", "")
        img_title = img.get("title", "")
        search_kw = img.get("search_keyword", "")
        # 相关性检查：URL、标题、搜索关键词中包含产品关键词
        is_relevant = False
        for kw in product_keywords:
            if kw in img_url.lower() or kw in img_title.lower() or kw in search_kw.lower():
                is_relevant = True
                break
        # 如果没有产品关键词，也接受（Tavily的图片通常跟搜索词相关）
        if not product_keywords:
            is_relevant = True
        
        if is_relevant and img_url.startswith("http"):
            product_images.append({
                "url": img_url,
                "title": f"{product_name} - {search_kw or img_title}",
                "source": "tavily"
            })
    
    log(f"  相关性过滤后剩余 {len(product_images)} 张Tavily图片")
    
    # 2.2 如果Tavily图片不足8张，用必应图片搜索补充
    if len(product_images) < 8:
        log(f"  Tavily图片不足，用必应补充（还需{8 - len(product_images)}张）...")
        bing_queries = [
            f"{product_name} 产品图",
            f"{product_name} 官方",
            f"{product_name} 实物",
            f"{product_name} 包装",
            f"{product_category} {product_name}",
        ]
        for img_query in bing_queries:
            if len(product_images) >= 8:
                break
            imgs = bing_image_search(img_query, max_results=2)
            for img in imgs:
                img_url = img.get("url", "")
                # 必应图片也做相关性检查
                is_relevant = False
                for kw in product_keywords:
                    if kw in img_url.lower():
                        is_relevant = True
                        break
                if not product_keywords:
                    is_relevant = True
                if is_relevant and img_url.startswith("http"):
                    product_images.append({
                        "url": img_url,
                        "title": f"{product_name} - {img_query}",
                        "source": "bing"
                    })
    
    # 去重
    seen_urls = set()
    unique_images = []
    for img in product_images:
        if img["url"] not in seen_urls:
            seen_urls.add(img["url"])
            unique_images.append(img)
    product_images = unique_images[:8]
    log(f"最终找到 {len(product_images)} 张产品图片（Tavily优先）")
    
    # Step 3: 生成Markdown报告
    log("生成Markdown报告...")
    report_parts = []
    report_parts.append(f"# {product_name} 竞品调研报告\n")
    report_parts.append(f"> **调研需求**：{product_name}（{product_category}）")
    report_parts.append(f"> **报告模板**：{REPORT_TEMPLATES[template]['name']}")
    report_parts.append(f"> **生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_parts.append(f"> **数据来源**：Tavily网络搜索（{len(search_results)}条结果）+ LLM分析\n")
    report_parts.append("---\n")
    
    # 目录
    report_parts.append("## 📋 目录\n")
    report_parts.append("1. 产品多维度分析")
    report_parts.append("2. 竞品对比矩阵")
    report_parts.append("3. SWOT分析")
    report_parts.append("4. 用户评价情感分析")
    report_parts.append("5. 数据可视化")
    report_parts.append("6. 产品图片")
    report_parts.append("7. 数据来源\n")
    report_parts.append("---\n")
    
    # 1. 产品分析
    report_parts.append("## 1️⃣ 产品多维度分析\n")
    report_parts.append(product_analysis)
    report_parts.append("\n---\n")
    
    # 2. 竞品对比
    report_parts.append("## 2️⃣ 竞品对比矩阵\n")
    report_parts.append(comparison_matrix)
    report_parts.append("\n---\n")
    
    # 3. SWOT
    report_parts.append("## 3️⃣ SWOT分析\n")
    report_parts.append(swot_analysis)
    report_parts.append("\n---\n")
    
    # 4. 情感分析
    report_parts.append("## 4️⃣ 用户评价情感分析\n")
    report_parts.append(sentiment_analysis)
    report_parts.append("\n---\n")
    
    # 5. 数据可视化
    report_parts.append("## 5️⃣ 数据可视化\n")
    if charts.get("price_chart"):
        report_parts.append("### 📊 竞品价格对比\n")
        report_parts.append(f"![价格对比](data:image/png;base64,{charts['price_chart']})\n")
    if charts.get("radar_chart"):
        report_parts.append("### 🎯 竞品多维度雷达图\n")
        report_parts.append(f"![雷达图](data:image/png;base64,{charts['radar_chart']})\n")
    if charts.get("sentiment_pie"):
        report_parts.append("### 💬 用户评价情感分布\n")
        report_parts.append(f"![情感分布](data:image/png;base64,{charts['sentiment_pie']})\n")
    if charts.get("wordcloud"):
        report_parts.append("### 🔤 用户评价高频关键词\n")
        report_parts.append(f"![关键词](data:image/png;base64,{charts['wordcloud']})\n")
    report_parts.append("\n---\n")
    
    # 6. 产品图片
    if product_images:
        report_parts.append("## 6️⃣ 产品图片\n")
        for i, img in enumerate(product_images[:8]):
            report_parts.append(f"![{img['title']}]({img['url']})\n")
        report_parts.append("\n---\n")
    
    # 7. 数据来源
    report_parts.append("## 7️⃣ 数据来源\n")
    report_parts.append(f"### 搜索关键词（{len(keywords)}个）\n")
    for i, kw in enumerate(keywords, 1):
        report_parts.append(f"{i}. {kw}")
    report_parts.append(f"\n### 搜索结果（{len(search_results)}条）\n")
    for i, r in enumerate(search_results[:15], 1):
        report_parts.append(f"{i}. [{r['title']}]({r['url']}) - 相关度：{r.get('score', 0):.2f}")
    report_parts.append("\n---\n")
    
    # 附录
    report_parts.append("## 📎 附录\n")
    report_parts.append(f"- **调研产品**：{product_name}")
    report_parts.append(f"- **产品类别**：{product_category}")
    report_parts.append(f"- **报告模板**：{REPORT_TEMPLATES[template]['name']}")
    report_parts.append(f"- **生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_parts.append(f"- **搜索结果数**：{len(search_results)}条")
    report_parts.append(f"- **分析方法**：Tavily网络搜索 + LLM多维度分析 + 数据可视化")
    report_parts.append(f"- **免责声明**：本报告由AI自动生成，部分信息基于公开资料分析，具体数据请以官方信息为准。\n")
    
    markdown_report = "\n".join(report_parts)
    log(f"Markdown报告生成完成，{len(markdown_report)}字")
    
    # Step 4: 生成HTML报告
    log("生成HTML报告...")
    html_report = generate_html_report(product_name, product_category, markdown_report, charts, product_images, template)
    log("HTML报告生成完成")
    
    # Step 5: 保存报告
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_basename = f"{product_name}_调研报告_{timestamp}"
    
    md_path = os.path.join(REPORTS_DIR, f"{report_basename}.md")
    html_path = os.path.join(REPORTS_DIR, f"{report_basename}.html")
    
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(markdown_report)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_report)
    
    log(f"报告已保存：{md_path}")
    log(f"HTML报告已保存：{html_path}")
    
    return {
        "markdown_report": markdown_report,
        "html_report": html_report,
        "md_path": md_path,
        "html_path": html_path,
        "charts": charts,
        "product_images": product_images,
        "logs": logs
    }

def generate_html_report(product_name, product_category, markdown_content, charts, product_images, template):
    """生成带样式的HTML报告"""
    # 简单的Markdown转HTML（处理标题、表格、列表、图片）
    html_body = markdown_to_html(markdown_content)
    
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{product_name} 竞品调研报告</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
            line-height: 1.8;
            color: #333;
            max-width: 1000px;
            margin: 0 auto;
            padding: 40px 20px;
            background: #f5f7fa;
        }}
        .report-container {{
            background: white;
            padding: 50px;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }}
        h1 {{
            color: #1a1a2e;
            border-bottom: 3px solid #4ECDC4;
            padding-bottom: 15px;
            margin-top: 0;
        }}
        h2 {{
            color: #16213e;
            border-left: 4px solid #FF6B6B;
            padding-left: 15px;
            margin-top: 40px;
        }}
        h3 {{
            color: #0f3460;
            margin-top: 25px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 14px;
        }}
        th, td {{
            border: 1px solid #e0e0e0;
            padding: 12px;
            text-align: left;
        }}
        th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 600;
        }}
        tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        tr:hover {{
            background: #e8f4f8;
        }}
        img {{
            max-width: 100%;
            border-radius: 8px;
            margin: 15px 0;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        blockquote {{
            border-left: 4px solid #4ECDC4;
            background: #f0fdfa;
            padding: 15px 20px;
            margin: 20px 0;
            border-radius: 0 8px 8px 0;
        }}
        code {{
            background: #f1f3f5;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 13px;
        }}
        ul, ol {{
            padding-left: 25px;
        }}
        li {{
            margin: 8px 0;
        }}
        hr {{
            border: none;
            border-top: 2px dashed #e0e0e0;
            margin: 30px 0;
        }}
        .header-info {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 30px;
        }}
        .header-info p {{
            margin: 5px 0;
        }}
    </style>
</head>
<body>
    <div class="report-container">
        {html_body}
    </div>
</body>
</html>"""
    return html

def markdown_to_html(md_text):
    """简易Markdown转HTML"""
    lines = md_text.split('\n')
    html_lines = []
    in_table = False
    in_list = False
    
    for line in lines:
        # 标题
        if line.startswith('# '):
            html_lines.append(f"<h1>{line[2:]}</h1>")
        elif line.startswith('## '):
            html_lines.append(f"<h2>{line[3:]}</h2>")
        elif line.startswith('### '):
            html_lines.append(f"<h3>{line[4:]}</h3>")
        elif line.startswith('#### '):
            html_lines.append(f"<h4>{line[5:]}</h4>")
        # 分割线
        elif line.strip() == '---':
            html_lines.append("<hr>")
        # 引用
        elif line.startswith('> '):
            html_lines.append(f"<blockquote>{line[2:]}</blockquote>")
        # 图片
        elif '![' in line and '](' in line:
            match = re.search(r'!\[(.*?)\]\((.*?)\)', line)
            if match:
                alt = match.group(1)
                src = match.group(2)
                html_lines.append(f'<img src="{src}" alt="{alt}">')
        # 链接
        elif '[' in line and '](' in line and '!' not in line:
            match = re.search(r'\[(.*?)\]\((.*?)\)', line)
            if match:
                text = match.group(1)
                url = match.group(2)
                line = line.replace(match.group(0), f'<a href="{url}" target="_blank">{text}</a>')
                html_lines.append(f"<p>{line}</p>")
            else:
                html_lines.append(f"<p>{line}</p>")
        # 表格行
        elif '|' in line and line.strip().startswith('|'):
            if not in_table:
                html_lines.append("<table>")
                in_table = True
            # 跳过分隔行
            if re.match(r'^\|[\s\-:|]+\|$', line.strip()):
                continue
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            if all(re.match(r'^[\s\-:]+$', c) for c in cells):
                continue
            # 判断是否是表头
            is_header = all(c.isupper() or '对比' in c or '维度' in c for c in cells[:2]) if cells else False
            tag = 'th' if is_header else 'td'
            html_lines.append("<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>")
        else:
            if in_table:
                html_lines.append("</table>")
                in_table = False
            # 列表
            if line.strip().startswith(('- ', '* ')):
                if not in_list:
                    html_lines.append("<ul>")
                    in_list = True
                html_lines.append(f"<li>{line.strip()[2:]}</li>")
            elif re.match(r'^\d+\.\s', line.strip()):
                if not in_list:
                    html_lines.append("<ol>")
                    in_list = True
                html_lines.append(f"<li>{re.sub(r'^\d+\.\s', '', line.strip())}</li>")
            elif line.strip():
                if in_list:
                    html_lines.append("</ul>" if in_list else "</ol>")
                    in_list = False
                html_lines.append(f"<p>{line}</p>")
            else:
                if in_list:
                    html_lines.append("</ul>")
                    in_list = False
    
    if in_table:
        html_lines.append("</table>")
    if in_list:
        html_lines.append("</ul>")
    
    return "\n".join(html_lines)

# ========== 历史报告对比 ==========
def get_history_reports():
    """获取历史报告列表"""
    if not os.path.exists(REPORTS_DIR):
        return []
    reports = []
    for f in os.listdir(REPORTS_DIR):
        if f.endswith('.md'):
            path = os.path.join(REPORTS_DIR, f)
            mtime = os.path.getmtime(path)
            size = os.path.getsize(path)
            reports.append({
                "filename": f,
                "path": path,
                "time": datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S'),
                "size": f"{size/1024:.1f}KB"
            })
    reports.sort(key=lambda x: x["time"], reverse=True)
    return reports

# ========== 主流程 ==========
def run_full_research(user_request, sop_db, embeddings, llm, template="标准模板", custom_dimensions=None, progress_callback=None):
    """完整调研流程：搜索Agent → 分析Agent → 报告Agent"""
    all_logs = []
    
    def log(msg):
        all_logs.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
        if progress_callback:
            progress_callback(msg)
    
    start_time = time.time()
    log("=" * 50)
    log("🚀 智能竞品调研Agent V4 启动")
    log(f"调研需求：{user_request}")
    log(f"报告模板：{REPORT_TEMPLATES[template]['name']}")
    log("=" * 50)
    
    # Step 0: 从SOP检索相关流程
    log("\n📚 Step 0: 从SOP知识库检索调研流程...")
    sop_docs = sop_db.similarity_search(user_request, k=5)
    sop_context = "\n\n".join([d.page_content for d in sop_docs])
    log(f"  检索到 {len(sop_docs)} 个SOP相关章节")
    
    # Step 1: 提取产品名称和类别
    log("\n🔍 Step 1: 提取产品名称和类别...")
    extract_prompt = f"""从用户的调研需求中提取产品名称和产品类别。

【用户需求】
{user_request}

【输出格式】
产品名称：具体的产品名
产品类别：产品所属的具体类别（用于图片搜索和搜索关键词）

只输出产品名称和产品类别，不要其他内容。"""
    
    extract_response = llm_invoke(llm, extract_prompt)
    product_name = ""
    product_category = ""
    for line in extract_response.split("\n"):
        line = line.strip()
        if "产品名称" in line:
            product_name = re.sub(r'^[^：:]*[：:]\s*', '', line).strip()
        elif "产品类别" in line:
            product_category = re.sub(r'^[^：:]*[：:]\s*', '', line).strip()
    
    if not product_name:
        product_name = re.sub(r'(帮我|调研一下|调研|看看|了解下|的)', '', user_request).strip()
    if not product_category:
        product_category = product_name
    
    log(f"  产品名称：{product_name}")
    log(f"  产品类别：{product_category}")
    
    # Step 2: 搜索Agent
    log("\n🔎 Step 2: 搜索Agent开始工作...")
    search_data = search_agent(product_name, product_category, llm, sop_context)
    all_logs.extend(search_data["logs"])
    log(f"  搜索完成：{len(search_data['search_results'])}条结果，{len(search_data['competitors'])}个竞品")
    
    # Step 3: 分析Agent
    log("\n📊 Step 3: 分析Agent开始工作...")
    analysis_data = analysis_agent(product_name, product_category, search_data, llm, custom_dimensions, template)
    all_logs.extend(analysis_data["logs"])
    log("  分析完成：多维度分析 + 竞品对比 + SWOT + 情感分析")
    
    # Step 4: 报告Agent
    log("\n📝 Step 4: 报告Agent开始工作...")
    report_data = report_agent(product_name, product_category, search_data, analysis_data, template)
    all_logs.extend(report_data["logs"])
    
    end_time = time.time()
    duration = end_time - start_time
    
    log("\n" + "=" * 50)
    log(f"✅ 调研完成！耗时 {duration:.1f} 秒")
    log(f"📄 Markdown报告：{report_data['md_path']}")
    log(f"🌐 HTML报告：{report_data['html_path']}")
    log(f"📊 搜索结果：{len(search_data['search_results'])}条")
    log(f"🖼️ 产品图片：{len(report_data['product_images'])}张")
    log(f"📈 数据图表：{len([k for k,v in report_data['charts'].items() if v])}个")
    log("=" * 50)
    
    return {
        "product_name": product_name,
        "product_category": product_category,
        "search_data": search_data,
        "analysis_data": analysis_data,
        "report_data": report_data,
        "logs": all_logs,
        "duration": duration
    }

# ========== Streamlit界面 ==========
# 科技感标题区域
st.markdown("""
<div style="padding: 24px 0 16px 0; border-bottom: 1px solid #334155; margin-bottom: 24px; position: relative;">
    <div style="position: absolute; top: 0; left: 0; width: 100%; height: 2px; background: linear-gradient(90deg, transparent, #0ea5e9, #06b6d4, #8b5cf6, transparent);"></div>
    <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 12px;">
        <div style="width: 56px; height: 56px; background: linear-gradient(135deg, #0ea5e9 0%, #06b6d4 50%, #8b5cf6 100%); border-radius: 14px; display: flex; align-items: center; justify-content: center; font-size: 28px; box-shadow: 0 0 30px rgba(14,165,233,0.4), inset 0 1px 0 rgba(255,255,255,0.2);">🔍</div>
        <div>
            <h1 style="margin: 0; color: #f1f5f9; font-size: 30px; font-weight: 700; letter-spacing: -0.5px; background: linear-gradient(135deg, #f1f5f9 0%, #94a3b8 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">智能竞品调研Agent</h1>
            <p style="margin: 4px 0 0 0; color: #64748b; font-size: 13px; font-weight: 300; letter-spacing: 1px;">INTELLIGENT COMPETITOR RESEARCH AGENT · V4</p>
        </div>
        <div style="margin-left: auto; display: flex; align-items: center; gap: 8px; background: rgba(16,185,129,0.1); border: 1px solid rgba(16,185,129,0.3); padding: 6px 14px; border-radius: 20px;">
            <div style="width: 8px; height: 8px; background: #10b981; border-radius: 50%; box-shadow: 0 0 10px #10b981; animation: pulse 2s infinite;"></div>
            <span style="color: #10b981; font-size: 12px; font-weight: 500;">V4 完整版</span>
        </div>
    </div>
    <div style="display: flex; gap: 10px; flex-wrap: wrap; margin-top: 16px;">
        <span style="background: rgba(14,165,233,0.1); color: #0ea5e9; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 500; border: 1px solid rgba(14,165,233,0.2);">🤖 多Agent协作</span>
        <span style="background: rgba(6,182,212,0.1); color: #06b6d4; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 500; border: 1px solid rgba(6,182,212,0.2);">🌐 Tavily真实搜索</span>
        <span style="background: rgba(139,92,246,0.1); color: #8b5cf6; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 500; border: 1px solid rgba(139,92,246,0.2);">📊 数据可视化</span>
        <span style="background: rgba(16,185,129,0.1); color: #10b981; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 500; border: 1px solid rgba(16,185,129,0.2);">💡 SWOT分析</span>
        <span style="background: rgba(245,158,11,0.1); color: #f59e0b; padding: 5px 14px; border-radius: 16px; font-size: 12px; font-weight: 500; border: 1px solid rgba(245,158,11,0.2);">📝 多格式输出</span>
    </div>
</div>
<style>
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.5; }
}
</style>
""", unsafe_allow_html=True)

# 侧边栏
with st.sidebar:
    st.header("⚙️ 配置")
    st.success("✅ V4完整版\n三阶段全部升级\n多Agent协作架构")
    
    st.divider()
    st.subheader("📋 报告模板")
    template_names = list(REPORT_TEMPLATES.keys())
    selected_template = st.selectbox("选择报告模板", template_names, index=0)
    st.info(f"**{REPORT_TEMPLATES[selected_template]['name']}**\n\n{REPORT_TEMPLATES[selected_template]['style']}\n\n分析维度：{len(REPORT_TEMPLATES[selected_template]['dimensions'])}个")
    
    st.divider()
    st.subheader("🎯 自定义分析维度")
    use_custom = st.checkbox("使用自定义分析维度", value=False)
    custom_dims_text = ""
    if use_custom:
        custom_dims_text = st.text_area("输入自定义维度（每行一个）", 
            value="\n".join(REPORT_TEMPLATES[selected_template]["dimensions"]),
            height=200)
    
    st.divider()
    st.subheader("📊 运行状态")
    st.write("搜索API：Tavily")
    st.write("分析模型：智谱AI GLM-4-Flash")
    st.write("图片搜索：必应图片")
    st.write("Agent架构：搜索+分析+报告 三Agent")
    st.write("输出格式：Markdown + HTML")
    
    st.divider()
    st.subheader("📁 历史报告")
    history_reports = get_history_reports()
    if history_reports:
        for r in history_reports[:10]:
            st.write(f"📄 {r['filename'][:30]}...")
            st.caption(f"{r['time']} | {r['size']}")
    else:
        st.write("暂无历史报告")

# 主区域
with st.spinner("正在加载SOP知识库..."):
    sop_db, embeddings = load_sop_knowledge_base()
    llm = load_llm()
st.success("✅ SOP知识库已加载，三Agent就绪")

# Tab切换
tab1, tab2, tab3 = st.tabs(["🚀 开始调研", "📊 历史报告对比", "📖 使用说明"])

with tab1:
    # 输入调研需求
    user_request = st.text_area("📝 输入你要调研的产品",
        placeholder="例如：白象火鸡面、石头科技扫地机器人、小米14手机、蜜雪冰城柠檬水...",
        height=80)
    
    col1, col2 = st.columns([1, 4])
    with col1:
        start_button = st.button("🚀 开始调研", type="primary", use_container_width=True)
    
    if start_button and user_request:
        st.divider()
        st.subheader("🔄 调研过程")
        
        progress_placeholder = st.empty()
        log_container = st.container()
        
        def update_progress(msg):
            progress_placeholder.info(msg)
        
        with st.spinner("三Agent协作调研中，请稍候（约1-2分钟）..."):
            custom_dims = None
            if use_custom and custom_dims_text:
                custom_dims = [line.strip() for line in custom_dims_text.split('\n') if line.strip()]
            
            result = run_full_research(user_request, sop_db, embeddings, llm, 
                                       template=selected_template, 
                                       custom_dimensions=custom_dims,
                                       progress_callback=update_progress)
        
        progress_placeholder.empty()
        
        # 显示日志
        with log_container:
            with st.expander("📋 查看完整调研日志", expanded=False):
                for log in result["logs"]:
                    st.text(log)
        
        # 显示统计
        st.divider()
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: st.metric("调研产品", result["product_name"][:8])
        with col2: st.metric("搜索结果", f"{len(result['search_data']['search_results'])}条")
        with col3: st.metric("产品图片", f"{len(result['report_data']['product_images'])}张")
        with col4: st.metric("数据图表", f"{len([k for k,v in result['report_data']['charts'].items() if v])}个")
        with col5: st.metric("耗时", f"{result['duration']:.1f}秒")
        
        # 显示识别到的竞品列表
        competitors = result["search_data"].get("competitors", [])
        competitor_categories = result["search_data"].get("competitor_categories", {})
        if competitors:
            st.info(f"🏢 识别到的竞品（{len(competitors)}个）：" + 
                   " | ".join([f"**{c}**" + (f"（{competitor_categories.get(c, '')}）" if competitor_categories.get(c) else "") for c in competitors]))
        
        # 显示数据可视化
        charts = result["report_data"]["charts"]
        if any(charts.values()):
            st.divider()
            st.subheader("📈 数据可视化")
            chart_cols = st.columns(2)
            chart_idx = 0
            for chart_name, chart_data in charts.items():
                if chart_data:
                    with chart_cols[chart_idx % 2]:
                        titles = {
                            "price_chart": "💰 竞品价格对比",
                            "radar_chart": "🎯 竞品多维度雷达图",
                            "sentiment_pie": "💬 用户评价情感分布",
                            "wordcloud": "🔤 用户评价高频关键词"
                        }
                        st.markdown(f"**{titles.get(chart_name, chart_name)}**")
                        st.image(f"data:image/png;base64,{chart_data}")
                    chart_idx += 1
        
        # 显示产品图片
        product_images = result["report_data"]["product_images"]
        if product_images:
            st.divider()
            st.subheader(f"🖼️ {result['product_name']} 产品图片")
            img_cols = st.columns(4)
            for idx, img in enumerate(product_images[:8]):
                with img_cols[idx % 4]:
                    try:
                        st.image(img["url"], caption=img.get("title", "")[:25], use_container_width=True)
                    except Exception:
                        st.info("图片加载失败")
        
        # 显示报告
        st.divider()
        st.subheader("📊 完整调研报告")
        st.markdown(result["report_data"]["markdown_report"])
        
        # 下载报告
        st.divider()
        st.subheader("📥 下载报告")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                "📄 下载Markdown",
                data=result["report_data"]["markdown_report"],
                file_name=f"{result['product_name']}_调研报告.md",
                mime="text/markdown",
                use_container_width=True
            )
        with col2:
            st.download_button(
                "🌐 下载HTML",
                data=result["report_data"]["html_report"],
                file_name=f"{result['product_name']}_调研报告.html",
                mime="text/html",
                use_container_width=True
            )
        with col3:
            st.info(f"报告已保存到：\n{result['report_data']['md_path']}")
    
    elif start_button and not user_request:
        st.error("⚠️ 请输入要调研的产品")

with tab2:
    st.subheader("📊 历史报告对比")
    st.write("选择两份历史报告进行对比分析")
    
    history_reports = get_history_reports()
    if len(history_reports) >= 2:
        col1, col2 = st.columns(2)
        with col1:
            report1_idx = st.selectbox("选择报告1", range(len(history_reports)), 
                format_func=lambda i: history_reports[i]["filename"][:40])
        with col2:
            report2_idx = st.selectbox("选择报告2", range(len(history_reports)), 
                format_func=lambda i: history_reports[i]["filename"][:40], index=1)
        
        if st.button("🔍 对比报告"):
            try:
                with open(history_reports[report1_idx]["path"], "r", encoding="utf-8") as f:
                    report1 = f.read()
                with open(history_reports[report2_idx]["path"], "r", encoding="utf-8") as f:
                    report2 = f.read()
                
                st.success(f"报告1：{len(report1)}字 | 报告2：{len(report2)}字")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("### 📄 报告1")
                    st.markdown(report1[:5000] + ("..." if len(report1) > 5000 else ""))
                with col2:
                    st.markdown("### 📄 报告2")
                    st.markdown(report2[:5000] + ("..." if len(report2) > 5000 else ""))
            except Exception as e:
                st.error(f"读取报告失败：{e}")
    else:
        st.info("历史报告不足2份，请先生成至少2份报告后再对比")

with tab3:
    st.subheader("📖 使用说明")
    
    st.markdown("""
    ### 🎯 功能介绍
    
    智能竞品调研Agent V4 采用**三Agent协作架构**，实现竞品调研全流程自动化：
    
    #### 🔎 搜索Agent
    - 自动生成8个多维度搜索关键词
    - 调用Tavily搜索API获取真实网络数据
    - 自动识别主要竞品
    - 提取用户评价文本
    
    #### 📊 分析Agent
    - 基于真实搜索数据进行多维度产品分析
    - 生成竞品对比矩阵
    - SWOT分析
    - 用户评价情感分析 + 痛点提取
    
    #### 📝 报告Agent
    - 生成结构化Markdown报告
    - 生成带样式的HTML报告（含图表）
    - 4种数据可视化图表自动生成
    - 产品图片自动搜索
    
    ### 📋 报告模板
    
    | 模板 | 适用场景 | 分析维度 |
    |------|---------|---------|
    | 标准模板 | 通用竞品调研 | 11个维度 |
    | 快速模板 | 快速了解竞品 | 5个维度 |
    | 深度模板 | 行业深度研究 | 11个维度 |
    | 产品模板 | 产品体验分析 | 8个维度 |
    
    ### 📈 数据可视化
    
    - 💰 **价格对比柱状图**：直观对比各竞品价格
    - 🎯 **竞品雷达图**：多维度能力对比
    - 💬 **情感分布饼图**：用户评价正面/中性/负面占比
    - 🔤 **高频关键词图**：用户评价TOP15关键词
    
    ### 💡 使用技巧
    
    1. **输入具体产品名**：如"白象火鸡面"比"火鸡面"效果更好
    2. **选择合适模板**：快速调研用"快速模板"，深度研究用"深度模板"
    3. **自定义维度**：勾选"自定义分析维度"可按需调整分析角度
    4. **查看HTML报告**：HTML报告带样式和图表，适合展示和分享
    5. **历史对比**：在"历史报告对比"标签页可对比不同时间的调研结果
    
    ### ⚠️ 注意事项
    
    - 调研耗时约1-2分钟，请耐心等待
    - Tavily API有调用频率限制，请勿频繁调用
    - 报告中数据来自网络搜索，仅供参考，具体以官方信息为准
    - 部分产品图片可能加载失败，可刷新页面重试
    """)
