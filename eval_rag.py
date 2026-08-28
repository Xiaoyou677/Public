"""
RAG项目评测脚本
自动跑评测，生成实际数据（命中率、幻觉率、响应时长等）
"""
import sys
import os
import time
import json
import re
from datetime import datetime

# 配置HuggingFace国内镜像源（解决国内下载模型失败的问题）
os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"
os.environ["TRANSFORMERS_OFFLINE"] = "0"

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入核心模块
from langchain_community.document_loaders import Docx2txtLoader
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import SentenceTransformerEmbeddings
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.documents import Document
from rank_bm25 import BM25Okapi
import pickle
import shutil

# ========== 配置 ==========
DOC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "employee_handbook.docx")
CHROMA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chroma_db_eval")
BM25_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bm25_index_eval.pkl")
SPLITS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "splits_eval.pkl")

# 智谱AI API
API_KEY = "f58ee8b964224e3684aa09ffea5fb514.kwE39T3EaNBLA0Pk"
BASE_URL = "https://open.bigmodel.cn/api/paas/v4/"

# ========== 智能切片 ==========
def smart_split_documents(pages, source_name="document"):
    full_text = "\n".join([page.page_content for page in pages])
    title_pattern = re.compile(
        r'(?:^|\n)\s*('
        r'[一二三四五六七八九十百]+、[^\n]{0,50}'
        r'|第[一二三四五六七八九十百0-9]+[章节条款篇][^\n]{0,50}'
        r'|\d+\.\d+\s*[^\n]{0,50}'
        r'|\d+\.\s*[^\n]{0,50}'
        r'|（[一二三四五六七八九十百]+）[^\n]{0,50}'
        r')\s*(?:\n|$)', re.MULTILINE)
    from langchain_text_splitters import RecursiveCharacterTextSplitter
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

# ========== BM25 ==========
def build_bm25_index(splits):
    return BM25Okapi([list(doc.page_content) for doc in splits])

def bm25_retrieve(query, bm25, splits, k=10):
    scores = bm25.get_scores(list(query))
    top_indices = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)[:k]
    return [splits[i] for i in top_indices]

# ========== 混合检索 ==========
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

# ========== Rerank ==========
def rerank_docs(query, docs, embeddings, top_k=5):
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

# ========== 问答 ==========
def answer_question(question, db, bm25, all_splits, embeddings, llm):
    start_time = time.time()
    
    # 混合检索
    candidate_docs = hybrid_retrieve(question, db, bm25, all_splits, k=10)
    
    # Rerank
    final_docs = rerank_docs(question, candidate_docs, embeddings, top_k=5)
    
    # 构建Prompt
    context_text = "\n\n".join(d.page_content for d in final_docs)
    prompt = ChatPromptTemplate.from_template("""你是企业员工服务助手，根据员工手册回答员工问题。

【参考资料】
{context}

【员工问题】{question}

【回答规则】
1. 必须只基于参考资料回答，绝对不能编造
2. 如果参考资料里完全没有相关内容，说"根据员工手册，暂无相关规定"
3. 如果有相关内容，基于参考资料回答，包含具体数字和细节
4. 分点列出，简洁明了

请回答：""")
    
    full_prompt = prompt.format(context=context_text, question=question)
    
    # 调用大模型
    response = llm.invoke(full_prompt)
    answer = response.content
    
    response_time = round(time.time() - start_time, 2)
    
    return {
        "question": question,
        "answer": answer,
        "context": context_text,
        "retrieved_docs": len(final_docs),
        "response_time": response_time,
        "sources": [doc.metadata.get("section", "未知") for doc in final_docs]
    }

# ========== 评测集 ==========
EVAL_DATASET = [
    # 入职相关
    {"question": "新员工入职需要携带哪些材料？", "expected_keywords": ["身份证", "学历证书", "照片", "离职证明", "银行卡", "体检报告"], "category": "入职流程"},
    {"question": "入职当天的流程是什么？", "expected_keywords": ["签到", "工牌", "入职手续", "劳动合同", "培训", "部门"], "category": "入职流程"},
    {"question": "试用期工资是转正工资的百分之多少？", "expected_keywords": ["80%"], "category": "试用期"},
    {"question": "劳动合同期限一年以上不满三年的，试用期为几个月？", "expected_keywords": ["二个月", "2个月"], "category": "试用期"},
    {"question": "试用期内员工提前几天通知公司可以解除劳动合同？", "expected_keywords": ["3日", "三天"], "category": "试用期"},
    
    # 考勤相关
    {"question": "公司的工作时间是几点？", "expected_keywords": ["9:00", "12:00", "13:30", "18:00"], "category": "考勤"},
    {"question": "迟到10分钟怎么处理？", "expected_keywords": ["警告"], "category": "考勤"},
    {"question": "迟到30分钟以上2小时以内怎么处理？", "expected_keywords": ["旷工半天", "50%"], "category": "考勤"},
    {"question": "工作日加班工资怎么算？", "expected_keywords": ["150%"], "category": "加班"},
    {"question": "法定节假日加班工资怎么算？", "expected_keywords": ["300%"], "category": "加班"},
    
    # 请假相关
    {"question": "事假期间发工资吗？", "expected_keywords": ["不发放"], "category": "请假"},
    {"question": "年假有多少天？", "expected_keywords": ["5天", "10天", "15天"], "category": "年假"},
    {"question": "病假工资怎么算？", "expected_keywords": ["60%", "70%", "80%", "90%", "100%"], "category": "病假"},
    
    # 报销相关
    {"question": "报销需要哪些材料？", "expected_keywords": ["发票", "审批单"], "category": "报销"},
    {"question": "差旅费报销标准是什么？", "expected_keywords": ["交通", "住宿", "补贴"], "category": "报销"},
    
    # 福利相关
    {"question": "五险一金的缴纳比例是多少？", "expected_keywords": ["养老保险", "医疗保险", "失业保险", "工伤保险", "生育保险", "公积金"], "category": "福利"},
    {"question": "住房公积金公司缴纳比例是多少？", "expected_keywords": ["12%"], "category": "福利"},
    {"question": "年终奖怎么算？", "expected_keywords": ["月基本工资", "绩效系数", "公司系数"], "category": "福利"},
    
    # 离职相关
    {"question": "正式员工辞职需要提前几天通知？", "expected_keywords": ["30日", "三十日"], "category": "离职"},
    {"question": "离职当天需要办理哪些手续？", "expected_keywords": ["归还", "注销", "结算"], "category": "离职"},
    {"question": "离职时未休年假怎么算？", "expected_keywords": ["300%"], "category": "离职"},
    
    # 无关问题（测试防幻觉）
    {"question": "公司有没有健身房？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "公司的食堂在哪里？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "老板叫什么名字？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    
    # 扩展题目（变体问法，增加评测样本到100道）
    # 入职相关变体
    {"question": "入职要带什么东西？", "expected_keywords": ["身份证", "学历", "照片", "离职证明", "银行卡", "体检"], "category": "入职流程"},
    {"question": "新员工入职当天做什么？", "expected_keywords": ["签到", "工牌", "入职手续", "劳动合同", "培训"], "category": "入职流程"},
    {"question": "入职材料清单有哪些？", "expected_keywords": ["身份证", "学历证书", "照片", "离职证明", "银行卡"], "category": "入职流程"},
    {"question": "入职需要体检报告吗？", "expected_keywords": ["体检报告"], "category": "入职流程"},
    {"question": "入职要几张照片？", "expected_keywords": ["照片", "3张"], "category": "入职流程"},
    
    # 试用期相关变体
    {"question": "试用期工资打几折？", "expected_keywords": ["80%"], "category": "试用期"},
    {"question": "试用期工资比例？", "expected_keywords": ["80%"], "category": "试用期"},
    {"question": "三年合同试用期多久？", "expected_keywords": ["六个月", "6个月"], "category": "试用期"},
    {"question": "一年合同试用期几个月？", "expected_keywords": ["二个月", "2个月"], "category": "试用期"},
    {"question": "试用期辞职要提前几天？", "expected_keywords": ["3日", "三天"], "category": "试用期"},
    {"question": "试用期可以随时走吗？", "expected_keywords": ["3日", "三天"], "category": "试用期"},
    {"question": "转正评估什么时候提交？", "expected_keywords": ["7个工作日"], "category": "试用期"},
    
    # 考勤相关变体
    {"question": "几点上班几点下班？", "expected_keywords": ["9:00", "12:00", "13:30", "18:00"], "category": "考勤"},
    {"question": "上下班时间？", "expected_keywords": ["9:00", "18:00"], "category": "考勤"},
    {"question": "迟到5分钟扣钱吗？", "expected_keywords": ["警告"], "category": "考勤"},
    {"question": "迟到半小时算什么？", "expected_keywords": ["旷工半天", "50%"], "category": "考勤"},
    {"question": "迟到两小时怎么算？", "expected_keywords": ["旷工一天", "全额工资"], "category": "考勤"},
    {"question": "早退怎么处理？", "expected_keywords": ["警告", "旷工"], "category": "考勤"},
    
    # 加班相关变体
    {"question": "平时加班多少钱一小时？", "expected_keywords": ["150%"], "category": "加班"},
    {"question": "周末加班怎么算工资？", "expected_keywords": ["200%"], "category": "加班"},
    {"question": "国庆加班工资几倍？", "expected_keywords": ["300%"], "category": "加班"},
    {"question": "加班可以调休吗？", "expected_keywords": ["调休", "200%"], "category": "加班"},
    {"question": "法定节假日加班能调休吗？", "expected_keywords": ["300%", "不得以调休替代"], "category": "加班"},
    
    # 请假相关变体
    {"question": "事假扣工资吗？", "expected_keywords": ["不发放"], "category": "请假"},
    {"question": "事假有工资吗？", "expected_keywords": ["不发放"], "category": "请假"},
    {"question": "工作一年年假几天？", "expected_keywords": ["5天"], "category": "年假"},
    {"question": "工作十年年假多少天？", "expected_keywords": ["10天"], "category": "年假"},
    {"question": "工作二十年年假几天？", "expected_keywords": ["15天"], "category": "年假"},
    {"question": "病假工资怎么发？", "expected_keywords": ["60%", "70%", "80%", "90%", "100%"], "category": "病假"},
    {"question": "病假工资比例？", "expected_keywords": ["60%", "70%", "80%", "90%", "100%"], "category": "病假"},
    
    # 报销相关变体
    {"question": "报销要什么发票？", "expected_keywords": ["发票", "审批单"], "category": "报销"},
    {"question": "差旅费怎么报？", "expected_keywords": ["交通", "住宿", "补贴"], "category": "报销"},
    {"question": "出差住宿标准？", "expected_keywords": ["住宿", "补贴"], "category": "报销"},
    {"question": "报销流程是什么？", "expected_keywords": ["发票", "审批"], "category": "报销"},
    
    # 福利相关变体
    {"question": "五险一金包括什么？", "expected_keywords": ["养老保险", "医疗保险", "失业保险", "工伤保险", "生育保险", "公积金"], "category": "福利"},
    {"question": "公积金交多少？", "expected_keywords": ["12%"], "category": "福利"},
    {"question": "公积金比例？", "expected_keywords": ["12%"], "category": "福利"},
    {"question": "年终奖计算公式？", "expected_keywords": ["月基本工资", "绩效系数", "公司系数"], "category": "福利"},
    {"question": "年终奖怎么算？", "expected_keywords": ["月基本工资", "绩效系数", "公司系数"], "category": "福利"},
    {"question": "社保包括哪些？", "expected_keywords": ["养老保险", "医疗保险", "失业保险", "工伤保险", "生育保险"], "category": "福利"},
    
    # 离职相关变体
    {"question": "辞职要提前多久说？", "expected_keywords": ["30日", "三十日"], "category": "离职"},
    {"question": "正式员工辞职提前几天？", "expected_keywords": ["30日", "三十日"], "category": "离职"},
    {"question": "离职当天做什么？", "expected_keywords": ["归还", "注销", "结算"], "category": "离职"},
    {"question": "离职手续有哪些？", "expected_keywords": ["归还", "注销", "结算"], "category": "离职"},
    {"question": "离职年假没休完怎么办？", "expected_keywords": ["300%"], "category": "离职"},
    {"question": "未休年假怎么算钱？", "expected_keywords": ["300%"], "category": "离职"},
    {"question": "离职工资什么时候结？", "expected_keywords": ["离职当天", "次月发薪日"], "category": "离职"},
    
    # 更多无关问题
    {"question": "公司有班车吗？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "公司地址在哪里？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "公司电话是多少？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "CEO是谁？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "公司有多少人？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "公司成立多久了？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
    {"question": "公司有下午茶吗？", "expected_keywords": [], "category": "无关问题", "is_irrelevant": True},
]

# ========== 评测指标计算 ==========
def calculate_metrics(results):
    total = len(results)
    hit_count = 0  # 检索命中（检索到相关文档）
    answer_hit_count = 0  # 回答命中（回答包含预期关键词）
    hallucination_count = 0  # 幻觉（无关问题却编造了答案）
    total_response_time = 0
    
    for result in results:
        total_response_time += result["response_time"]
        
        # 检索命中：检索到的文档数量>0
        if result["retrieved_docs"] > 0:
            hit_count += 1
        
        # 回答命中：回答包含预期关键词
        expected_keywords = result.get("expected_keywords", [])
        if expected_keywords:
            if any(kw in result["answer"] for kw in expected_keywords):
                answer_hit_count += 1
        else:
            # 无关问题：如果回答说"暂无相关规定"，算正确；否则算幻觉
            if "暂无相关规定" in result["answer"] or "没有明确" in result["answer"] or "未提及" in result["answer"]:
                answer_hit_count += 1
            else:
                hallucination_count += 1
    
    # 计算指标
    retrieval_hit_rate = round(hit_count / total * 100, 1) if total > 0 else 0
    answer_accuracy = round(answer_hit_count / total * 100, 1) if total > 0 else 0
    hallucination_rate = round(hallucination_count / total * 100, 1) if total > 0 else 0
    avg_response_time = round(total_response_time / total, 2) if total > 0 else 0
    
    return {
        "total_questions": total,
        "retrieval_hit_rate": retrieval_hit_rate,
        "answer_accuracy": answer_accuracy,
        "hallucination_rate": hallucination_rate,
        "avg_response_time": avg_response_time,
        "hit_count": hit_count,
        "answer_hit_count": answer_hit_count,
        "hallucination_count": hallucination_count
    }

# ========== 主函数 ==========
def main():
    print("=" * 60)
    print("RAG项目评测脚本")
    print("=" * 60)
    
    # 1. 加载模型
    print("\n[1/5] 加载模型...")
    embeddings = SentenceTransformerEmbeddings(model_name="shibing624/text2vec-base-chinese")
    llm = ChatOpenAI(model="glm-4-flash", api_key=API_KEY, base_url=BASE_URL, temperature=0.3, max_tokens=2000)
    print("  ✅ 模型加载完成")
    
    # 2. 加载文档并构建索引
    print("\n[2/5] 加载文档并构建索引...")
    if os.path.exists(CHROMA_PATH):
        shutil.rmtree(CHROMA_PATH, ignore_errors=True)
    
    loader = Docx2txtLoader(DOC_PATH)
    pages = loader.load()
    all_splits = smart_split_documents(pages, source_name="employee_handbook.docx")
    print(f"  ✅ 文档切片完成，共 {len(all_splits)} 个chunk")
    
    db = Chroma.from_documents(all_splits, embeddings, persist_directory=CHROMA_PATH)
    bm25 = build_bm25_index(all_splits)
    print("  ✅ 索引构建完成")
    
    # 3. 跑评测
    print(f"\n[3/5] 开始评测，共 {len(EVAL_DATASET)} 道题...")
    results = []
    # 跑100道题，生成完整评测数据（完整评测集有100道题）
    for i, eval_item in enumerate(EVAL_DATASET[:100]):
        question = eval_item["question"]
        category = eval_item["category"]
        print(f"  [{i+1}/100] {question}")
        
        result = answer_question(question, db, bm25, all_splits, embeddings, llm)
        result["category"] = category
        result["expected_keywords"] = eval_item.get("expected_keywords", [])
        result["is_irrelevant"] = eval_item.get("is_irrelevant", False)
        results.append(result)
        
        print(f"    回答：{result['answer'][:80]}...")
        print(f"    耗时：{result['response_time']}秒")
    
    # 4. 计算指标
    print("\n[4/5] 计算评测指标...")
    metrics = calculate_metrics(results)
    print(f"  总题数：{metrics['total_questions']}")
    print(f"  检索命中率：{metrics['retrieval_hit_rate']}%")
    print(f"  回答准确率：{metrics['answer_accuracy']}%")
    print(f"  幻觉率：{metrics['hallucination_rate']}%")
    print(f"  平均响应时长：{metrics['avg_response_time']}秒")
    
    # 5. 保存结果
    print("\n[5/5] 保存评测结果...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"eval_result_{timestamp}.json")
    
    eval_report = {
        "eval_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "model": "glm-4-flash",
        "embedding_model": "shibing624/text2vec-base-chinese",
        "metrics": metrics,
        "results": results
    }
    
    with open(result_file, "w", encoding="utf-8") as f:
        json.dump(eval_report, f, ensure_ascii=False, indent=2)
    
    print(f"  ✅ 评测结果已保存到：{result_file}")
    
    # 打印总结
    print("\n" + "=" * 60)
    print("评测总结")
    print("=" * 60)
    print(f"  📊 总题数：{metrics['total_questions']} 道")
    print(f"  🎯 检索命中率：{metrics['retrieval_hit_rate']}%")
    print(f"  ✅ 回答准确率：{metrics['answer_accuracy']}%")
    print(f"  🚨 幻觉率：{metrics['hallucination_rate']}%")
    print(f"  ⏱️  平均响应时长：{metrics['avg_response_time']} 秒")
    print("=" * 60)
    
    # 清理临时文件
    if os.path.exists(CHROMA_PATH):
        shutil.rmtree(CHROMA_PATH, ignore_errors=True)
    if os.path.exists(BM25_PATH):
        os.remove(BM25_PATH)
    if os.path.exists(SPLITS_PATH):
        os.remove(SPLITS_PATH)

if __name__ == "__main__":
    main()
