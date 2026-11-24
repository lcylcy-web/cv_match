import streamlit as st
import pandas as pd
import os
import json
import re
from openai import OpenAI

st.set_page_config(page_title="简历岗位匹配工具", page_icon="📄", layout="wide")
st.title("📄 简历与岗位智能匹配 ")

# ================= 侧边栏配置 =================
st.sidebar.header("🔑 配置大模型")

# 0. 左侧导航：文字版报告 / 表格版一览表
view_mode = st.sidebar.radio(
    "展示版本：",
    ["📄 文字版报告", "📋 表格版一览表"],
    help="文字版查看 AI 分析结论；表格版把匹配结果整理成结构化表格。"
)

# 1. 选择模型
model_choice = st.sidebar.selectbox(
    "请选择大模型：",
    ["通义千问 (Qwen)", "DeepSeek", "豆包 (Doubao)"]
)

# 2. 根据选择自动配置默认参数
if model_choice == "通义千问 (Qwen)":
    st.sidebar.markdown("[获取阿里云百炼 API Key](https://bailian.console.aliyun.com/)")
    default_base_url = "https://dashscope.aliyuncs.com/compatible-mode/v1"
    default_model = "qwen-plus"
elif model_choice == "DeepSeek":
    st.sidebar.markdown("[获取阿里云百炼 API Key](https://bailian.console.aliyuncs.com/)")
    # 使用你专属工作空间的 Base URL
    default_base_url = "https://ws-bfasmdpgnnoew27c.cn-beijing.maas.aliyuncs.com/compatible-mode/v1"
    # 更换为当前有效的模型名称
    default_model = "deepseek-v4-flash-0731"
elif model_choice == "豆包 (Doubao)":
    st.sidebar.markdown("[获取火山方舟 API Key](https://console.volcengine.com/ark/)")
    default_base_url = "https://ark.cn-beijing.volces.com/api/v3"
    default_model = "doubao-seed-2.0-lite-260428"

# 自动读取 Secrets 中的 API_KEY
api_key = st.sidebar.text_input("请输入 API Key：", value=st.secrets.get("API_KEY", ""), type="password")
model_name = st.sidebar.text_input("模型名称 / 接入点ID：", value=default_model)

# ================= 读取 Excel =================
def load_data():
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jobs.xlsx")
    if os.path.exists(excel_path):
        return pd.read_excel(excel_path)
    return None

df = load_data()

# ---------- 解析 AI 返回的结构化 JSON（失败返回 None） ----------
def parse_match_json(content: str):
    if not content:
        return None
    # 抽取 ```json ... ``` 代码块，否则取整段
    m = re.search(r"```(?:json)?\s*(.*?)```", content, re.S)
    raw = m.group(1).strip() if m else content.strip()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return None
    # 兼容两种结构：直接是列表，或包在 {"results": [...]} 里
    if isinstance(data, dict):
        for key in ("results", "jobs", "data", "列表"):
            if isinstance(data.get(key), list):
                data = data[key]
                break
        else:
            return None
    if not isinstance(data, list) or not data:
        return None
    return data

def normalize_jobs(items: list) -> pd.DataFrame:
    rows = []
    for i, it in enumerate(items, start=1):
        if not isinstance(it, dict):
            continue
        # 得分兼容多种字段名，提取数字
        score_raw = it.get("匹配得分", it.get("得分", it.get("score", "")))
        score_num = None
        if score_raw is not None:
            mm = re.search(r"\d+", str(score_raw))
            if mm:
                score_num = int(mm.group())
        rows.append({
            "排名": i,
            "企业名称": it.get("企业名称", it.get("公司", "")),
            "招聘岗位": it.get("招聘岗位", it.get("岗位", "")),
            "匹配得分": score_num if score_num is not None else "",
            "匹配理由": it.get("匹配理由", it.get("理由", it.get("匹配分析", ""))),
            "工作地点": it.get("工作地点", it.get("地点", "")),
            "薪资福利": it.get("薪资福利", it.get("薪资", "")),
            "更新日期": it.get("更新日期", it.get("日期", "")),
        })
    return pd.DataFrame(rows)

if df is None:
    st.error("❌ 未找到 jobs.xlsx，请确保文件在项目目录下！")
else:
    st.success(f"✅ 已成功加载 {len(df)} 个本地岗位数据！")
    resume_text = st.text_area("✍️ 请将简历内容（纯文本）粘贴到这里：", height=250)

    # 触发匹配：两个版本共用同一次 AI 调用，结果存入 session_state
    if st.button("🚀 开始 AI 智能匹配", type="primary"):
        if not api_key:
            st.warning("⚠️ 请先在左侧输入 API Key！")
        elif not resume_text.strip():
            st.warning("⚠️ 请先粘贴简历内容！")
        elif not model_name:
            st.warning("⚠️ 请填写模型名称或接入点ID！")
        else:
            with st.spinner(f"「{model_choice}」正在深度分析你的简历和所有岗位，请稍候..."):
                try:
                    # 动态初始化客户端
                    client = OpenAI(
                        api_key=api_key,
                        base_url=default_base_url
                    )

                    # 拼接岗位数据（精准适配你的9个表头）
                    jobs_text = ""
                    for index, row in df.iterrows():
                        jobs_text += f"岗位{index+1}：\n"
                        jobs_text += f"企业名称：{row.get('企业名称', '')}\n"
                        jobs_text += f"企业性质：{row.get('企业性质', '')}\n"
                        jobs_text += f"招聘岗位：{row.get('招聘岗位', '')}\n"
                        jobs_text += f"工作地点：{row.get('工作地点', '')}\n"
                        jobs_text += f"学历要求：{row.get('学历要求', '')}\n"
                        jobs_text += f"专业要求：{row.get('专业要求', '')}\n"
                        jobs_text += f"经验要求：{row.get('经验要求', '')}\n"
                        jobs_text += f"薪资福利：{row.get('薪资福利', '')}\n"
                        jobs_text += f"更新日期：{row.get('更新日期', '')}\n"
                        jobs_text += "---\n"

                    # 构建给 AI 的指令：要求同时给出文字结论和结构化 JSON，便于下方生成表格
                    prompt = f"""
                    你是一个专业的 HR 招聘助手。请根据以下候选人的简历内容，对比给出的岗位列表，评估匹配度（0-100分），并按得分从高到低排序。

                    【候选人简历】
                    {resume_text}

                    【岗位列表】
                    {jobs_text}

                    请先简要说明整体匹配结论，然后输出匹配度最高的前 5 个岗位。
                    除文字说明外，必须在最后用 ```json 代码块输出一个 JSON 数组，供系统生成表格。数组中每个对象包含以下字段：
                    - 企业名称
                    - 招聘岗位
                    - 匹配得分（0-100 的数字）
                    - 匹配理由（结合学历、专业、经验说明为什么匹配，控制在 80 字以内）
                    - 工作地点
                    - 薪资福利
                    - 更新日期

                    JSON 格式示例：
                    ```json
                    [
                      {{
                        "企业名称": "某某公司",
                        "招聘岗位": "行政岗",
                        "匹配得分": 88,
                        "匹配理由": "学历专业对口，具备行政与活动统筹经验，地点在天津",
                        "工作地点": "天津",
                        "薪资福利": "6000-7000，五险一金",
                        "更新日期": "2026.10.19"
                      }}
                    ]
                    """

                    # 调用大模型
                    response = client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": "你是一个资深的招聘专家，擅长人岗匹配分析。"},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.3,
                        max_tokens=2000
                    )

                    ai_content = response.choices[0].message.content
                    # 文字报告：去掉 JSON 代码块部分，避免重复
                    text_report = re.sub(r"```(?:json)?\s*.*?```", "", ai_content, flags=re.S).strip()
                    items = parse_match_json(ai_content)
                    result_df = normalize_jobs(items) if items else pd.DataFrame()

                    # 存入会话状态，供两个版本页面分别读取
                    st.session_state["text_report"] = text_report
                    st.session_state["result_df"] = result_df

                except Exception as e:
                    st.error(f"❌ 调用 AI 失败，错误信息：{e}")
                    st.info("💡 常见原因：1. API Key 无效；2. 模型名称/接入点ID 填错；3. 账号余额不足。")

    # ================= 左侧导航：两个独立页面 =================
    text_report = st.session_state.get("text_report", "")
    result_df = st.session_state.get("result_df", pd.DataFrame())

    if view_mode.startswith("📄"):
        # ---------- 页面一：文字版报告 ----------
        st.subheader(f"🎯 {model_choice} 智能匹配报告（文字版）")
        if text_report:
            st.markdown(text_report)
        else:
            st.info("💡 请先粘贴简历并点击「🚀 开始 AI 智能匹配」，生成后在此查看文字报告。")

        # 文字版同样列出全部匹配岗位结果（逐条卡片式呈现）
        if result_df is not None and not result_df.empty:
            st.divider()
            st.subheader("📌 匹配岗位明细")
            for _, row in result_df.iterrows():
                score = row.get("匹配得分", "")
                score_text = f"{score} 分" if score != "" else "—"
                with st.container(border=True):
                    st.markdown(
                        f"**{row.get('排名', '')}. {row.get('企业名称', '')} — {row.get('招聘岗位', '')}**　"
                        f"🔖 匹配得分：**{score_text}**"
                    )
                    reason = row.get("匹配理由", "")
                    if reason:
                        st.markdown(f"　匹配理由：{reason}")
                    detail_parts = []
                    if row.get("工作地点", ""):
                        detail_parts.append(f"📍 工作地点：{row.get('工作地点', '')}")
                    if row.get("薪资福利", ""):
                        detail_parts.append(f"💰 薪资福利：{row.get('薪资福利', '')}")
                    if row.get("更新日期", ""):
                        detail_parts.append(f"🗓 更新日期：{row.get('更新日期', '')}")
                    if detail_parts:
                        st.markdown("　" + "　｜　".join(detail_parts))

    else:
        # ---------- 页面二：表格版一览表（单独一页） ----------
        st.subheader("📋 匹配结果一览表（表格版）")
        if result_df is not None and not result_df.empty:
            st.dataframe(
                result_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "匹配得分": st.column_config.ProgressColumn(
                        "匹配得分", min_value=0, max_value=100, format="%d 分"
                    ),
                },
            )
            st.download_button(
                "⬇️ 下载匹配结果 CSV",
                data=result_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="AI匹配结果.csv",
                mime="text/csv",
            )
        else:
            st.info("💡 暂无表格数据。请先粘贴简历并点击「🚀 开始 AI 智能匹配」；若模型未返回结构化数据，可在文字版查看结论。")
