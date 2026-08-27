import os
os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"

import streamlit as st
import re
import tempfile
import pickle
import json
import time
import uuid
import shutil
from datetime import datetime
from rank_bm25 import BM25Okapi
from langchain_community.document_loaders import Docx2txtLoader, PyPDFLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import SentenceTransformerEmbeddings
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.documents import Document

# ========== 页面配置 ==========
st.set_page_config(page_title="企业员工智能服务助手", page_icon="👔", layout="wide", initial_sidebar_state="expanded")

# ========== 自定义CSS - 莫兰迪极简商务风 ==========
st.markdown("""
<style>
/* 全局样式 */
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700&display=swap');

* {
    font-family: 'Noto Sans SC', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

/* 主背景 - 米白色 */
.stApp {
    background: linear-gradient(135deg, #faf9f7 0%, #f5f3f0 100%);
}

/* 侧边栏 - 浅灰褐 */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ebe8e3 0%, #e5e2dd 100%);
    border-right: 1px solid #d4d0c9;
}

/* 主标题样式 */
h1 {
    color: #3d3833 !important;
    font-weight: 700 !important;
    letter-spacing: -0.5px;
    margin-bottom: 8px !important;
}

/* 副标题 */
.stCaption {
    color: #8a847c !important;
    font-size: 13px !important;
    font-weight: 300 !important;
}

/* Tab样式 */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: transparent;
    padding-bottom: 2px;
}

.stTabs [data-baseweb="tab"] {
    background: #ebe8e3;
    border-radius: 8px 8px 0 0;
    padding: 10px 20px;
    color: #6b6560;
    font-weight: 500;
    font-size: 14px;
    border: none;
    transition: all 0.3s ease;
}

.stTabs [data-baseweb="tab"]:hover {
    background: #e0ddd8;
    color: #3d3833;
}

.stTabs [aria-selected="true"] {
    background: #ffffff !important;
    color: #3d3833 !important;
    font-weight: 700 !important;
    box-shadow: 0 -2px 8px rgba(0,0,0,0.06);
}

/* 聊天消息样式 */
.stChatMessage {
    border-radius: 16px;
    padding: 16px 20px;
    margin-bottom: 12px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.04);
    transition: all 0.3s ease;
}

.stChatMessage:hover {
    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
}

/* 用户消息 - 莫兰迪蓝渐变 */
.stChatMessage[data-testid="stChatMessage-user"] {
    background: linear-gradient(135deg, #a8b5c4 0%, #96a5b8 100%);
    color: #ffffff;
    box-shadow: 0 2px 12px rgba(168,181,196,0.3);
    position: relative;
}

.stChatMessage[data-testid="stChatMessage-user"]::before {
    content: '';
    position: absolute;
    top: 0;
    right: 0;
    width: 4px;
    height: 100%;
    background: linear-gradient(180deg, #ffffff 0%, rgba(255,255,255,0.5) 100%);
    border-radius: 0 16px 16px 0;
}

.stChatMessage[data-testid="stChatMessage-user"] p,
.stChatMessage[data-testid="stChatMessage-user"] span {
    color: #ffffff !important;
}

/* 助手消息 - 莫兰迪灰绿精致卡片 */
.stChatMessage[data-testid="stChatMessage-assistant"] {
    background: linear-gradient(135deg, #e8ede8 0%, #dde5dd 100%);
    border: 1px solid #c8d4c8;
    box-shadow: 0 2px 12px rgba(168,196,168,0.2), 0 1px 3px rgba(0,0,0,0.02);
    position: relative;
}

.stChatMessage[data-testid="stChatMessage-assistant"]::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 4px;
    height: 100%;
    background: linear-gradient(180deg, #a8c4a8 0%, #8fb08f 100%);
    border-radius: 16px 0 0 16px;
}

/* 聊天输入框 */
.stChatInput {
    border-radius: 24px;
    border: 2px solid #d4d0c9;
    background: #ffffff;
    padding: 4px;
    transition: all 0.3s ease;
}

.stChatInput:focus-within {
    border-color: #a8b5c4;
    box-shadow: 0 0 0 4px rgba(168,181,196,0.15);
}

.stChatInput input {
    font-size: 15px;
    color: #3d3833;
}

/* 按钮样式 */
.stButton > button {
    border-radius: 10px;
    border: 1px solid #d4d0c9;
    background: #ffffff;
    color: #3d3833;
    font-weight: 500;
    padding: 8px 20px;
    transition: all 0.3s ease;
}

.stButton > button:hover {
    background: #ebe8e3;
    border-color: #a8b5c4;
    color: #3d3833;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}

.stButton > button:active {
    transform: translateY(0);
}

/* 主按钮 */
.stButton [kind="primary"] {
    background: linear-gradient(135deg, #a8b5c4 0%, #96a5b8 100%) !important;
    border: none !important;
    color: #ffffff !important;
}

.stButton [kind="primary"]:hover {
    background: linear-gradient(135deg, #96a5b8 0%, #8495aa 100%) !important;
}

/* Expander样式 */
.streamlit-expanderHeader {
    background: #f5f3f0;
    border-radius: 8px;
    padding: 10px 16px;
    color: #6b6560;
    font-weight: 500;
    font-size: 13px;
}

.streamlit-expanderContent {
    background: #faf9f7;
    border-radius: 0 0 8px 8px;
    padding: 12px 16px;
}

/* Metric卡片 */
[data-testid="stMetric"] {
    background: #ffffff;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.04);
    border: 1px solid #e8e5e0;
}

[data-testid="stMetricLabel"] {
    color: #8a847c !important;
    font-size: 12px !important;
    font-weight: 400 !important;
}

[data-testid="stMetricValue"] {
    color: #3d3833 !important;
    font-size: 28px !important;
    font-weight: 700 !important;
}

/* 信息框 */
.stInfo {
    background: linear-gradient(135deg, #e8eef5 0%, #dde6f0 100%);
    border-left: 4px solid #a8b5c4;
    border-radius: 8px;
    padding: 12px 16px;
    color: #3d3833;
}

.stSuccess {
    background: linear-gradient(135deg, #e8f0e8 0%, #dde8dd 100%);
    border-left: 4px solid #a8c4a8;
    border-radius: 8px;
    padding: 12px 16px;
    color: #3d3833;
}

.stWarning {
    background: linear-gradient(135deg, #f5f0e8 0%, #f0e8dd 100%);
    border-left: 4px solid #c4b8a8;
    border-radius: 8px;
    padding: 12px 16px;
    color: #3d3833;
}

.stError {
    background: linear-gradient(135deg, #f5e8e8 0%, #f0dddd 100%);
    border-left: 4px solid #c4a8a8;
    border-radius: 8px;
    padding: 12px 16px;
    color: #3d3833;
}

/* 分割线 */
.stDivider {
    border-color: #e0ddd8 !important;
    margin: 20px 0 !important;
}

/* 加载动画 */
.stSpinner > div {
    border-color: #a8b5c4 !important;
    border-right-color: transparent !important;
}

/* 滚动条美化 */
::-webkit-scrollbar {
    width: 6px;
    height: 6px;
}

::-webkit-scrollbar-track {
    background: #f5f3f0;
}

::-webkit-scrollbar-thumb {
    background: #c4c0b9;
    border-radius: 3px;
}

::-webkit-scrollbar-thumb:hover {
    background: #a8a49d;
}

/* 头像样式 */
.stChatMessage [data-testid="stChatMessageAvatarUser"] {
    background: linear-gradient(135deg, #a8b5c4 0%, #96a5b8 100%) !important;
}

.stChatMessage [data-testid="stChatMessageAvatarAssistant"] {
    background: linear-gradient(135deg, #c4b8a8 0%, #b8a896 100%) !important;
}
</style>
""", unsafe_allow_html=True)

# ========== 路径配置 ==========
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "data")
DOC_VERSIONS_DIR = os.path.join(SCRIPT_DIR, "doc_versions")
FEEDBACK_FILE = os.path.join(DATA_DIR, "feedback.json")
EVAL_FILE = os.path.join(DATA_DIR, "eval_set_v2.json")
QA_LOG_FILE = os.path.join(DATA_DIR, "qa_log.json")
DOC_VERSIONS_FILE = os.path.join(DATA_DIR, "doc_versions.json")

# 确保目录存在
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(DOC_VERSIONS_DIR, exist_ok=True)

# ========== 初始化 session_state ==========
if 'messages' not in st.session_state:
    st.session_state.messages = []
if 'doc_list' not in st.session_state:
    st.session_state.doc_list = []
if 'current_qa_id' not in st.session_state:
    st.session_state.current_qa_id = None

# ========== JSON 工具函数 ==========
def load_json(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default
    return default

def save_json(filepath, data):
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ========== 功能1：智能切片 ==========
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

# ========== 功能2：BM25 ==========
def build_bm25_index(splits):
    return BM25Okapi([list(doc.page_content) for doc in splits])

def bm25_retrieve(query, bm25, splits, k=10):
    scores = bm25.get_scores(list(query))
    top_indices = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)[:k]
    return [splits[i] for i in top_indices if scores[i] > 0]

# ========== 功能3：混合检索 ==========
def hybrid_retrieve(query, db, bm25, splits, k=10):
    vector_docs = db.similarity_search(query, k=k)
    keyword_docs = bm25_retrieve(query, bm25, splits, k=k)
    seen, merged = set(), []
    for doc in vector_docs + keyword_docs:
        key = doc.page_content[:100]
        if key not in seen:
            seen.add(key)
            merged.append(doc)
    return merged

# ========== 功能4：Rerank ==========
def rerank_docs(query, docs, embeddings, top_k=3):
    if not docs:
        return []
    query_emb = embeddings.embed_query(query)
    scored = []
    for doc in docs:
        doc_emb = embeddings.embed_query(doc.page_content)
        sim = sum(a*b for a,b in zip(query_emb, doc_emb))
        scored.append((doc, sim))
    scored.sort(key=lambda x: x[1], reverse=True)
    return [doc for doc,_ in scored[:top_k]]

# ========== 功能5：查询改写 ==========
def rewrite_query(query, history, llm):
    if not history:
        return query
    recent = history[-6:]
    history_text = "\n".join([f"{'用户' if m['role']=='user' else '助手'}：{m['content'][:200]}" for m in recent])
    rewrite_prompt = f"""你是查询改写助手。根据对话历史，把用户当前问题改写成适合文档检索的问题。
规则：1.只补全指代，不改变核心意图 2.不要过度具体化 3.保留关键词多样性 4.只输出改写后的问题
对话历史：{history_text}
当前问题：{query}
改写后："""
    try:
        result = llm.invoke(rewrite_prompt).content.strip()
        return result if len(result) >= 2 else query
    except Exception:
        return query

# ========== 网络检查 ==========
def check_api_connection():
    import urllib.request, urllib.error
    try:
        req = urllib.request.Request("https://open.bigmodel.cn")
        with urllib.request.urlopen(req, timeout=5) as resp:
            return True
    except urllib.error.HTTPError as e:
        return e.code in (401, 403)
    except Exception:
        return False

# ========== 知识库 ==========
@st.cache_resource
def get_knowledge_base():
    doc_path = os.path.join(SCRIPT_DIR, "employee_handbook.docx")
    chroma_path = os.path.join(SCRIPT_DIR, "chroma_db_persist")
    bm25_path = os.path.join(SCRIPT_DIR, "bm25_index.pkl")
    splits_path = os.path.join(SCRIPT_DIR, "splits.pkl")
    if not os.path.exists(doc_path):
        return None, None, None, "未找到 company_rule.docx"
    try:
        embeddings = SentenceTransformerEmbeddings(model_name="all-MiniLM-L6-v2")
        if os.path.exists(chroma_path) and os.listdir(chroma_path) and os.path.exists(bm25_path) and os.path.exists(splits_path):
            db = Chroma(persist_directory=chroma_path, embedding_function=embeddings)
            with open(bm25_path, "rb") as f: bm25 = pickle.load(f)
            with open(splits_path, "rb") as f: all_splits = pickle.load(f)
            return db, bm25, all_splits, {"name": "employee_handbook.docx", "chunks": len(all_splits)}
        loader = Docx2txtLoader(doc_path)
        pages = loader.load()
        all_splits = smart_split_documents(pages, source_name="company_rule.docx")
        if os.path.exists(chroma_path): shutil.rmtree(chroma_path, ignore_errors=True)
        db = Chroma.from_documents(all_splits, embeddings, persist_directory=chroma_path)
        bm25 = build_bm25_index(all_splits)
        with open(bm25_path, "wb") as f: pickle.dump(bm25, f)
        with open(splits_path, "wb") as f: pickle.dump(all_splits, f)
        return db, bm25, all_splits, {"name": "company_rule.docx", "chunks": len(all_splits)}
    except Exception as e:
        return None, None, None, f"知识库构建失败：{str(e)}"

# ========== 模型 ==========
@st.cache_resource
def load_models():
    embeddings = SentenceTransformerEmbeddings(model_name="all-MiniLM-L6-v2")
    # 优先从secrets读取API密钥（部署用），没有则用默认值（本地运行用）
    api_key = st.secrets.get("GLM_API_KEY", "f58ee8b964224e3684aa09ffea5fb514.kwE39T3EaNBLA0Pk")
    llm = ChatOpenAI(model="glm-4-flash", api_key=api_key,
        base_url="https://open.bigmodel.cn/api/paas/v4/", temperature=0.3, max_tokens=2000)
    return embeddings, llm

# ========== 问答日志记录 ==========
def log_qa(qa_id, question, answer, rewritten, retrieved_docs, full_prompt, response_time, sources):
    logs = load_json(QA_LOG_FILE, [])
    logs.append({
        "id": qa_id,
        "question": question,
        "answer": answer,
        "rewritten": rewritten,
        "retrieved_docs": [{"section": d.metadata.get("section",""), "content": d.page_content[:500]} for d in retrieved_docs],
        "full_prompt": full_prompt,
        "response_time": response_time,
        "sources": sources,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    })
    save_json(QA_LOG_FILE, logs)

# ========== 加载知识库 ==========
with st.spinner("正在加载知识库..."):
    db, bm25, all_splits, doc_info = get_knowledge_base()
if db is None:
    st.error(f"❌ {doc_info}")
else:
    if not st.session_state.doc_list:
        st.session_state.doc_list = [doc_info]

# ========== 侧边栏 ==========
with st.sidebar:
    st.header("📂 文档管理")
    uploaded_file = st.file_uploader("上传文档", type=["pdf","docx","txt"], help="支持 PDF、Word、txt")
    if uploaded_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name
        try:
            if uploaded_file.name.endswith(".pdf"): pages = PyPDFLoader(tmp_path).load()
            elif uploaded_file.name.endswith(".docx"): pages = Docx2txtLoader(tmp_path).load()
            else:
                with open(tmp_path,"r",encoding="utf-8") as f: text = f.read()
                pages = [Document(page_content=text, metadata={"source": uploaded_file.name})]
            splits = smart_split_documents(pages, source_name=uploaded_file.name)
            embeddings, _ = load_models()
            if db: db.add_documents(splits)
            # 保存文档版本
            versions = load_json(DOC_VERSIONS_FILE, [])
            version_id = str(uuid.uuid4())[:8]
            version_path = os.path.join(DOC_VERSIONS_DIR, f"{version_id}_{uploaded_file.name}")
            shutil.copy(tmp_path, version_path)
            versions.append({"id": version_id, "name": uploaded_file.name, "path": version_path,
                "chunks": len(splits), "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            save_json(DOC_VERSIONS_FILE, versions)
            st.session_state.doc_list.append({"name": uploaded_file.name, "chunks": len(splits)})
            st.success(f"✅ 已上传：{uploaded_file.name}（{len(splits)} 块）")
        except Exception as e:
            st.error(f"文档解析失败：{str(e)}")
        finally:
            os.unlink(tmp_path)
    st.divider()
    st.subheader("已加载文档")
    for doc in st.session_state.doc_list:
        st.write(f"📄 {doc['name']}（{doc['chunks']} 块）")
    st.divider()
    if st.button("🗑️ 清空对话记录"):
        st.session_state.messages = []
        st.rerun()

# ========== 主区域 Tab 切换 ==========
tab1, tab2, tab3, tab4, tab5 = st.tabs(["💬 智能问答", "📊 效果评估", "📝 评测集管理", "🔍 Bad Case分析", "📂 文档版本"])

# ==================== Tab1: 智能问答 ====================
with tab1:
    # 优雅的标题区域
    st.markdown("""
    <div style="padding: 24px 0 16px 0; border-bottom: 1px solid #e0ddd8; margin-bottom: 24px;">
        <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 12px;">
            <div style="width: 48px; height: 48px; background: linear-gradient(135deg, #a8b5c4 0%, #96a5b8 100%); border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 24px; box-shadow: 0 4px 12px rgba(168,181,196,0.3);">👔</div>
            <div>
                <h1 style="margin: 0; color: #3d3833; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">企业员工智能服务助手</h1>
                <p style="margin: 4px 0 0 0; color: #8a847c; font-size: 13px; font-weight: 300;">Enterprise Employee Intelligent Service Assistant</p>
            </div>
        </div>
        <div style="display: flex; gap: 12px; flex-wrap: wrap; margin-top: 12px;">
            <span style="background: #e8eef5; color: #6b7d8f; padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 500;">基于 RAG 技术</span>
            <span style="background: #e8f0e8; color: #6b8f6b; padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 500;">混合检索 + Rerank</span>
            <span style="background: #f5f0e8; color: #8f7d6b; padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 500;">查询改写 + 多轮对话</span>
            <span style="background: #f5e8e8; color: #8f6b6b; padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 500;">全生命周期覆盖</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    if db is None:
        st.warning("⚠️ 知识库加载失败")
    else:
        for idx, message in enumerate(st.session_state.messages):
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                if message["role"] == "assistant" and "sources" in message:
                    with st.expander("📚 引用来源", expanded=False):
                        for source in message["sources"]:
                            st.markdown(f"**{source['title']}**")
                            st.write(source["content"][:300] + "...")
                    # 反馈按钮
                    qa_id = message.get("qa_id")
                    if qa_id:
                        col1, col2, col3 = st.columns([1,1,4])
                        with col1:
                            if st.button("👍", key=f"up_{qa_id}"):
                                feedback = load_json(FEEDBACK_FILE, [])
                                feedback.append({"id": qa_id, "type": "upvote", "text": "", "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                                save_json(FEEDBACK_FILE, feedback)
                                st.success("感谢反馈！")
                        with col2:
                            if st.button("👎", key=f"down_{qa_id}"):
                                st.session_state[f"show_feedback_{qa_id}"] = True
                        if st.session_state.get(f"show_feedback_{qa_id}", False):
                            feedback_text = st.text_input("请描述问题：", key=f"fb_text_{qa_id}")
                            if st.button("提交反馈", key=f"fb_submit_{qa_id}"):
                                feedback = load_json(FEEDBACK_FILE, [])
                                feedback.append({"id": qa_id, "type": "downvote", "text": feedback_text, "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                                save_json(FEEDBACK_FILE, feedback)
                                st.session_state[f"show_feedback_{qa_id}"] = False
                                st.success("反馈已提交，我们会持续优化！")

        user_input = st.chat_input("请输入你的问题，例如：公司工作时间是几点？迟到10分钟怎么处理？年假有多少天？")
        if user_input:
            if not check_api_connection():
                error_msg = "⚠️ 无法连接到智谱AI API！请检查网络连接。"
                with st.chat_message("user"): st.markdown(user_input)
                st.session_state.messages.append({"role":"user","content":user_input})
                with st.chat_message("assistant"): st.error(error_msg)
                st.session_state.messages.append({"role":"assistant","content":error_msg})
            else:
                with st.chat_message("user"): st.markdown(user_input)
                st.session_state.messages.append({"role":"user","content":user_input})
                with st.chat_message("assistant"):
                    with st.spinner("正在思考..."):
                        start_time = time.time()
                        try:
                            embeddings, llm = load_models()
                            rewritten = rewrite_query(user_input, st.session_state.messages[:-1], llm)
                            candidate_docs = hybrid_retrieve(rewritten, db, bm25, all_splits, k=10)
                            if len(candidate_docs) < 3:
                                fallback = hybrid_retrieve(user_input, db, bm25, all_splits, k=10)
                                seen, merged = set(), []
                                for d in candidate_docs + fallback:
                                    k = d.page_content[:100]
                                    if k not in seen: seen.add(k); merged.append(d)
                                candidate_docs = merged
                            final_docs = rerank_docs(user_input, candidate_docs, embeddings, top_k=3)
                            prompt = ChatPromptTemplate.from_template("""你是企业员工服务助手，根据员工手册回答员工问题。
【对话历史】{history}
【参考资料】{context}
【员工问题】{question}
回答要求：1.基于参考资料回答 2.有部分相关内容就基于相关内容回答 3.完全没有相关内容才说无法回答 4.分点列出，简洁明了""")
                            history_text = "\n".join([f"{'用户' if m['role']=='user' else '助手'}：{m['content'][:200]}" for m in st.session_state.messages[-7:-1]]) or "（无）"
                            context_text = "\n\n".join(d.page_content for d in final_docs)
                            full_prompt = prompt.format(history=history_text, context=context_text, question=user_input)
                            full_response = ""
                            placeholder = st.empty()
                            for chunk in llm.stream(full_prompt):
                                if chunk and chunk.content:
                                    full_response += chunk.content
                                    placeholder.markdown(full_response + "▌")
                            placeholder.markdown(full_response)
                            response = full_response if full_response else "⚠️ 模型返回空回答，请换个问法重试。"
                            response_time = round(time.time() - start_time, 2)
                            sources = []
                            with st.expander("📚 引用来源", expanded=False):
                                for i, doc in enumerate(final_docs, 1):
                                    section = doc.metadata.get("section","未知章节")
                                    st.markdown(f"**[{i}] {section}**")
                                    st.write(doc.page_content[:300] + "...")
                                    sources.append({"title": f"[{i}] {section}", "content": doc.page_content})
                            with st.expander("🔍 检索过程详情", expanded=False):
                                st.markdown(f"**原始问题：** {user_input}")
                                st.markdown(f"**改写后问题：** {rewritten}")
                                st.markdown(f"**混合召回：** {len(candidate_docs)} 条")
                                st.markdown(f"**Rerank后：** {len(final_docs)} 条")
                                st.markdown(f"**响应时长：** {response_time} 秒")
                            qa_id = str(uuid.uuid4())[:8]
                            log_qa(qa_id, user_input, response, rewritten, final_docs, full_prompt, response_time, sources)
                            st.session_state.messages.append({"role":"assistant","content":response,"sources":sources,"qa_id":qa_id})
                        except Exception as e:
                            error_str = str(e)
                            if "Insufficient Balance" in error_str or "402" in error_str:
                                error_msg = "⚠️ 智谱AI账户余额不足！请去 https://open.bigmodel.cn/ 充值。"
                            elif "Invalid API Key" in error_str or "401" in error_str:
                                error_msg = "⚠️ API Key无效！请检查智谱AI API Key。"
                            else:
                                error_msg = f"生成回答时出错：{error_str}"
                            st.error(error_msg)
                            st.session_state.messages.append({"role":"assistant","content":error_msg})

# ==================== Tab2: 效果评估 ====================
with tab2:
    st.title("📊 效果评估仪表盘")
    st.caption("实时监控系统效果指标")
    feedback = load_json(FEEDBACK_FILE, [])
    qa_logs = load_json(QA_LOG_FILE, [])
    total_qa = len(qa_logs)
    upvotes = len([f for f in feedback if f["type"] == "upvote"])
    downvotes = len([f for f in feedback if f["type"] == "downvote"])
    total_feedback = upvotes + downvotes
    satisfaction = round(upvotes / total_feedback * 100, 1) if total_feedback > 0 else 0
    avg_response_time = round(sum(q["response_time"] for q in qa_logs) / total_qa, 2) if total_qa > 0 else 0
    # 幻觉率估算：点踩中提到"幻觉""编造""不对"的比例
    hallucination_count = len([f for f in feedback if f["type"]=="downvote" and any(k in f.get("text","") for k in ["幻觉","编造","不对","错误","瞎编"])])
    hallucination_rate = round(hallucination_count / downvotes * 100, 1) if downvotes > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("总问答数", total_qa)
    with col2: st.metric("用户满意度", f"{satisfaction}%", f"👍{upvotes} 👎{downvotes}")
    with col3: st.metric("平均响应时长", f"{avg_response_time}秒")
    with col4: st.metric("估算幻觉率", f"{hallucination_rate}%")

    st.divider()
    st.subheader("📈 反馈趋势")
    if feedback:
        feedback_dates = {}
        for f in feedback:
            date = f["timestamp"][:10]
            if date not in feedback_dates: feedback_dates[date] = {"up":0,"down":0}
            if f["type"] == "upvote": feedback_dates[date]["up"] += 1
            else: feedback_dates[date]["down"] += 1
        dates = sorted(feedback_dates.keys())
        up_data = [feedback_dates[d]["up"] for d in dates]
        down_data = [feedback_dates[d]["down"] for d in dates]
        chart_data = {"日期": dates, "👍 点赞": up_data, "👎 点踩": down_data}
        import pandas as pd
        st.bar_chart(pd.DataFrame(chart_data).set_index("日期"))
    else:
        st.info("暂无反馈数据，使用问答功能后会自动统计。")

    st.divider()
    st.subheader("📋 最近问答记录")
    if qa_logs:
        for q in reversed(qa_logs[-10:]):
            with st.expander(f"Q: {q['question'][:50]}... ({q['timestamp']})"):
                st.markdown(f"**回答：** {q['answer'][:300]}...")
                st.markdown(f"**改写后：** {q['rewritten']}")
                st.markdown(f"**响应时长：** {q['response_time']}秒")
                st.markdown(f"**检索文档数：** {len(q['retrieved_docs'])}")
    else:
        st.info("暂无问答记录。")

# ==================== Tab3: 评测集管理 ====================
with tab3:
    st.title("📝 评测集管理")
    st.caption("构建标准问答对，自动跑评测，版本对比")
    eval_set = load_json(EVAL_FILE, [])

    st.subheader("➕ 添加评测题")
    with st.form("add_eval"):
        eval_question = st.text_input("问题")
        eval_answer = st.text_area("标准答案")
        eval_category = st.selectbox("分类", ["入职流程","考勤制度","请假制度","报销制度","福利政策","IT与办公","培训发展","离职流程","边缘问题"])
        submitted = st.form_submit_button("添加")
        if submitted and eval_question and eval_answer:
            eval_set.append({"id": str(uuid.uuid4())[:8], "question": eval_question, "answer": eval_answer,
                "category": eval_category, "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            save_json(EVAL_FILE, eval_set)
            st.success("添加成功！")
            st.rerun()

    st.divider()
    st.subheader("📋 评测题列表")
    if eval_set:
        for i, item in enumerate(eval_set):
            with st.expander(f"[{item['category']}] {item['question'][:50]}..."):
                st.markdown(f"**问题：** {item['question']}")
                st.markdown(f"**标准答案：** {item['answer']}")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("🗑️ 删除", key=f"del_eval_{item['id']}"):
                        eval_set = [e for e in eval_set if e["id"] != item["id"]]
                        save_json(EVAL_FILE, eval_set)
                        st.rerun()
    else:
        st.info("暂无评测题，先添加一些吧。")

    st.divider()
    st.subheader("🚀 一键跑评测")
    if st.button("开始评测", type="primary") and eval_set and db:
        embeddings, llm = load_models()
        results = []
        progress = st.progress(0)
        for i, item in enumerate(eval_set):
            try:
                rewritten = rewrite_query(item["question"], [], llm)
                candidate = hybrid_retrieve(rewritten, db, bm25, all_splits, k=10)
                final = rerank_docs(item["question"], candidate, embeddings, top_k=3)
                context = "\n\n".join(d.page_content for d in final)
                eval_prompt = f"""根据参考资料回答问题。
参考资料：{context}
问题：{item['question']}
请回答："""
                answer = llm.invoke(eval_prompt).content
                # 简单相似度评分
                std_words = set(item["answer"])
                ans_words = set(answer)
                similarity = len(std_words & ans_words) / len(std_words) if std_words else 0
                score = round(similarity * 100, 1)
                results.append({"question": item["question"], "std_answer": item["answer"], "model_answer": answer, "score": score, "retrieved": len(final)})
            except Exception as e:
                results.append({"question": item["question"], "error": str(e), "score": 0})
            progress.progress((i+1)/len(eval_set))
        # 保存评测结果
        eval_results = load_json(os.path.join(DATA_DIR, "eval_results.json"), [])
        eval_results.append({"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "total": len(results),
            "avg_score": round(sum(r["score"] for r in results)/len(results),1) if results else 0, "details": results})
        save_json(os.path.join(DATA_DIR, "eval_results.json"), eval_results)
        st.success(f"评测完成！平均得分：{round(sum(r['score'] for r in results)/len(results),1) if results else 0}分")
        for r in results:
            with st.expander(f"Q: {r['question'][:40]}... 得分:{r.get('score',0)}"):
                if "error" in r: st.error(r["error"])
                else:
                    st.markdown(f"**标准答案：** {r['std_answer'][:200]}")
                    st.markdown(f"**模型回答：** {r['model_answer'][:200]}")

    st.divider()
    st.subheader("📊 历史评测版本对比")
    eval_results = load_json(os.path.join(DATA_DIR, "eval_results.json"), [])
    if eval_results:
        for er in reversed(eval_results[-5:]):
            st.write(f"**{er['timestamp']}** | 共{er['total']}题 | 平均分：{er['avg_score']}分")
    else:
        st.info("暂无评测记录。")

# ==================== Tab4: Bad Case 分析 ====================
with tab4:
    st.title("🔍 Bad Case 分析工具")
    st.caption("点踩的回答可以查看完整链路：检索了什么、Prompt是什么、模型输出是什么")
    feedback = load_json(FEEDBACK_FILE, [])
    qa_logs = load_json(QA_LOG_FILE, [])
    downvote_ids = [f["id"] for f in feedback if f["type"] == "downvote"]
    bad_cases = [q for q in qa_logs if q["id"] in downvote_ids]

    st.metric("Bad Case 总数", len(bad_cases))
    st.divider()

    if bad_cases:
        for case in reversed(bad_cases):
            fb = next((f for f in feedback if f["id"]==case["id"]), {})
            with st.expander(f"❌ Q: {case['question'][:50]}... ({case['timestamp']})"):
                st.markdown("### 📝 用户反馈")
                st.info(fb.get("text","无文本反馈"))
                st.markdown("### ❓ 原始问题")
                st.write(case["question"])
                st.markdown("### 🔄 改写后问题")
                st.write(case["rewritten"])
                st.markdown("### 🤖 模型回答")
                st.write(case["answer"])
                st.markdown("### 📚 检索到的文档")
                for i, doc in enumerate(case["retrieved_docs"], 1):
                    st.markdown(f"**文档{i} [{doc['section']}]：**")
                    st.write(doc["content"][:300] + "...")
                st.markdown("### 📋 完整 Prompt")
                st.code(case["full_prompt"], language="text")
                st.markdown("### ⏱️ 响应时长")
                st.write(f"{case['response_time']} 秒")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("✅ 标记已修复", key=f"fix_{case['id']}"):
                        fixed = load_json(os.path.join(DATA_DIR, "fixed_cases.json"), [])
                        fixed.append(case["id"])
                        save_json(os.path.join(DATA_DIR, "fixed_cases.json"), fixed)
                        st.success("已标记为已修复！")
    else:
        st.info("暂无 Bad Case，继续使用收集反馈吧。")

    st.divider()
    st.subheader("📊 问题分类统计")
    if bad_cases:
        categories = {"检索不到":0, "检索不准":0, "模型幻觉":0, "回答不完整":0,
                      "回答格式问题":0, "响应太慢":0, "上下文理解错误":0,
                      "查询改写错误":0, "切片问题":0, "引用来源错误":0}
        for f in feedback:
            if f["type"] == "downvote":
                bad_type = f.get("bad_case_type", "其他")
                if bad_type in categories:
                    categories[bad_type] += 1
                else:
                    text = f.get("text","")
                    if any(k in text for k in ["找不到","没有","检索"]): categories["检索不到"] += 1
                    elif any(k in text for k in ["不准","不对","相关"]): categories["检索不准"] += 1
                    elif any(k in text for k in ["幻觉","编造","瞎编"]): categories["模型幻觉"] += 1
                    elif any(k in text for k in ["不完整","太短","不全"]): categories["回答不完整"] += 1
                    else: categories["其他"] = categories.get("其他",0) + 1
        for cat, count in categories.items():
            if count > 0:
                st.write(f"- **{cat}**：{count} 个")
    else:
        st.info("暂无统计数据。")

# ==================== Tab5: 文档版本管理 ====================
with tab5:
    st.title("📂 文档版本管理")
    st.caption("文档更新后自动重建索引，保留历史版本，支持回滚")
    versions = load_json(DOC_VERSIONS_FILE, [])

    st.subheader("📋 文档版本列表")
    if versions:
        for v in reversed(versions):
            with st.expander(f"📄 {v['name']} ({v['timestamp']})"):
                st.markdown(f"**版本ID：** {v['id']}")
                st.markdown(f"**文档名：** {v['name']}")
                st.markdown(f"**切片数：** {v['chunks']}")
                st.markdown(f"**上传时间：** {v['timestamp']}")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("📥 下载此版本", key=f"dl_{v['id']}"):
                        if os.path.exists(v["path"]):
                            with open(v["path"], "rb") as f:
                                st.download_button("确认下载", f, file_name=v["name"])
                        else:
                            st.error("文件不存在")
                with col2:
                    if st.button("🔄 回滚到此版本", key=f"rollback_{v['id']}"):
                        st.info("回滚功能：将此版本重新加载到知识库（需手动确认）")
                        if st.button("确认回滚", key=f"confirm_rollback_{v['id']}"):
                            try:
                                if v["name"].endswith(".pdf"): pages = PyPDFLoader(v["path"]).load()
                                elif v["name"].endswith(".docx"): pages = Docx2txtLoader(v["path"]).load()
                                else:
                                    with open(v["path"],"r",encoding="utf-8") as f: text = f.read()
                                    pages = [Document(page_content=text, metadata={"source": v["name"]})]
                                splits = smart_split_documents(pages, source_name=v["name"])
                                embeddings, _ = load_models()
                                chroma_path = os.path.join(SCRIPT_DIR, "chroma_db_persist")
                                if os.path.exists(chroma_path): shutil.rmtree(chroma_path, ignore_errors=True)
                                new_db = Chroma.from_documents(splits, embeddings, persist_directory=chroma_path)
                                st.success("回滚成功！请刷新页面。")
                            except Exception as e:
                                st.error(f"回滚失败：{str(e)}")
    else:
        st.info("暂无文档版本，上传文档后会自动保存版本。")

    st.divider()
    st.subheader("🔧 索引管理")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 重建索引"):
            try:
                chroma_path = os.path.join(SCRIPT_DIR, "chroma_db_persist")
                bm25_path = os.path.join(SCRIPT_DIR, "bm25_index.pkl")
                splits_path = os.path.join(SCRIPT_DIR, "splits.pkl")
                for p in [chroma_path, bm25_path, splits_path]:
                    if os.path.exists(p):
                        if os.path.isdir(p): shutil.rmtree(p, ignore_errors=True)
                        else: os.remove(p)
                st.success("索引已清除，刷新页面后会自动重建。")
            except Exception as e:
                st.error(f"重建失败：{str(e)}")
    with col2:
        if st.button("📊 查看索引状态"):
            chroma_path = os.path.join(SCRIPT_DIR, "chroma_db_persist")
            bm25_path = os.path.join(SCRIPT_DIR, "bm25_index.pkl")
            splits_path = os.path.join(SCRIPT_DIR, "splits.pkl")
            st.write(f"向量库存在：{'✅' if os.path.exists(chroma_path) else '❌'}")
            st.write(f"BM25索引存在：{'✅' if os.path.exists(bm25_path) else '❌'}")
            st.write(f"切片数据存在：{'✅' if os.path.exists(splits_path) else '❌'}")
            if os.path.exists(splits_path):
                with open(splits_path,"rb") as f: splits = pickle.load(f)
                st.write(f"切片总数：{len(splits)}")

st.divider()
st.caption("💡 企业员工智能服务助手 | 用户反馈 | 效果评估 | 评测集管理 | Bad Case分析 | 文档版本管理 | 混合检索 | Rerank | 查询改写 | 多轮对话 | 智能切片")
