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

def _esc(text) -> str:
    """HTML 转义，防止岗位文本里的特殊字符破坏卡片结构。"""
    if text is None:
        return ""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )

# ---------- 卡片版样式（参照活动市集模板：绿色主色、圆角白卡、进度条） ----------
CARD_CSS = (
    "<style>"
    ".mk-wrap{max-width:520px;margin:0 auto}"
    ".mk-hero{position:relative;border-radius:18px;overflow:hidden;aspect-ratio:16/9;background:linear-gradient(135deg,#dfead4,#cfe0c0);margin-bottom:16px}"
    ".mk-hero-mask{position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,0) 35%,rgba(0,0,0,.45))}"
    ".mk-hero-title{position:absolute;left:16px;bottom:14px;color:#fff;font-size:22px;font-weight:700}"
    ".mk-hero-tag{position:absolute;right:14px;top:14px;background:rgba(255,255,255,.92);color:#4a8230;font-size:13px;font-weight:600;padding:6px 12px;border-radius:999px}"
    ".mk-head{padding:2px 4px 14px}"
    ".mk-head h1{font-size:20px;font-weight:700;margin-bottom:6px;color:#1f2a1c}"
    ".mk-head p{font-size:14px;color:#5b6657}"
    ".mk-card{background:#fff;border-radius:18px;padding:16px;margin-bottom:16px;box-shadow:0 1px 3px rgba(31,42,28,.06)}"
    ".mk-card-title{font-size:16px;font-weight:700;margin-bottom:14px;color:#1f2a1c}"
    ".mk-info-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}"
    ".mk-info-cell{background:#eef5e8;border-radius:14px;padding:12px 8px;text-align:center}"
    ".mk-info-cell .ico{width:22px;height:22px;color:#5f9e3f;margin:0 auto 6px;display:block}"
    ".mk-info-cell .label{font-size:12px;color:#5b6657;margin-bottom:4px}"
    ".mk-info-cell .value{font-size:15px;font-weight:700;color:#1f2a1c}"
    ".mk-info-cell .sub{font-size:11px;color:#5b6657;margin-top:2px}"
    ".mk-list{display:flex;flex-direction:column;gap:12px}"
    ".mk-job{border:1px solid #eceee8;border-radius:16px;padding:14px}"
    ".mk-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:8px}"
    ".mk-name{font-size:15px;font-weight:700;color:#1f2a1c;flex:1;min-width:0;line-height:1.4}"
    ".mk-rank{flex:0 0 auto;font-size:12px;font-weight:700;color:#4a8230;background:#eef5e8;border-radius:8px;padding:3px 9px;white-space:nowrap}"
    ".mk-reason{font-size:13px;color:#5b6657;margin-bottom:10px;line-height:1.6}"
    ".mk-meta{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12px;color:#5b6657;margin-bottom:10px}"
    ".mk-meta span{display:inline-flex;align-items:center;gap:5px;min-width:0}"
    ".mk-meta svg{width:13px;height:13px;color:#5f9e3f;flex:0 0 auto}"
    ".mk-score{display:flex;align-items:center;gap:10px}"
    ".mk-bar{flex:1;height:8px;border-radius:999px;background:#eceee8;overflow:hidden}"
    ".mk-bar i{display:block;height:100%;background:#5f9e3f;border-radius:999px}"
    ".mk-num{font-size:13px;font-weight:700;color:#4a8230;min-width:46px;text-align:right}"
    ".mk-banner{display:flex;align-items:center;gap:12px;background:#eef5e8;border-radius:16px;padding:14px;margin-bottom:16px}"
    ".mk-badge{flex:0 0 38px;height:38px;border-radius:50%;background:#5f9e3f;color:#fff;display:flex;align-items:center;justify-content:center}"
    ".mk-badge svg{width:20px;height:20px}"
    ".mk-banner .b-txt{font-size:14px;font-weight:600;color:#1f2a1c}"
    ".mk-banner .b-sub{font-size:12px;color:#5b6657}"
    ".mk-btn{display:block;width:100%;border:none;cursor:pointer;background:#5f9e3f;color:#fff;font-size:16px;font-weight:700;letter-spacing:1px;padding:15px;border-radius:16px;text-align:center}"
    "</style>"
)

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
        # ---------- 页面二：卡片版一览表（活动市集模板风 + AI 结果自动填充） ----------
        if result_df is not None and not result_df.empty:
            # 汇总统计：岗位数、最高得分、平均得分
            scores = [s for s in result_df["匹配得分"].tolist() if s != "" and s is not None]
            total = len(result_df)
            top_score = max(scores) if scores else 0
            avg_score = round(sum(scores) / len(scores)) if scores else 0

            st.markdown(CARD_CSS, unsafe_allow_html=True)

            # 顶部图卡 + 标题副标题
            hero_html = (
                '<div class="mk-wrap">'
                '<div class="mk-hero"><div class="mk-hero-mask"></div>'
                '<div class="mk-hero-title">简历岗位智能匹配</div>'
                '<div class="mk-hero-tag">匹配完成</div></div>'
                '<div class="mk-head"><h1>匹配岗位推荐</h1>'
                f'<p>共筛选出 {total} 个高匹配岗位，按匹配得分从高到低排列</p></div>'
            )

            # 三栏汇总信息小卡（对应图一 时间/地点/报名）
            info_html = (
                '<div class="mk-card"><div class="mk-card-title">匹配概览</div>'
                '<div class="mk-info-grid">'
                '<div class="mk-info-cell">'
                '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M19 8v6M22 11h-6"/></svg>'
                f'<div class="label">推荐岗位</div><div class="value">{total} 个</div><div class="sub">AI 精选 Top</div></div>'
                '<div class="mk-info-cell">'
                '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 2l2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 15.3 6.8 18.1l1-5.8L3.5 8.2l5.9-.9z"/></svg>'
                f'<div class="label">最高得分</div><div class="value">{top_score} 分</div><div class="sub">榜首岗位</div></div>'
                '<div class="mk-info-cell">'
                '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>'
                f'<div class="label">平均得分</div><div class="value">{avg_score} 分</div><div class="sub">整体匹配度</div></div>'
                '</div></div>'
            )

            # 岗位卡片流
            cards_html = []
            for _, row in result_df.iterrows():
                rank = row.get("排名", "")
                company = _esc(row.get("企业名称", ""))
                job = _esc(row.get("招聘岗位", ""))
                reason = _esc(row.get("匹配理由", ""))
                location = _esc(row.get("工作地点", ""))
                salary = _esc(row.get("薪资福利", ""))
                date = _esc(row.get("更新日期", ""))

                score = row.get("匹配得分", "")
                if score == "" or score is None:
                    score_num = 0
                    score_text = "—"
                else:
                    score_num = int(score)
                    score_text = f"{score_num} 分"

                meta_parts = []
                if location:
                    meta_parts.append(
                        f'<span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">'
                        f'<path d="M12 21s-7-6.2-7-11a7 7 0 1 1 14 0c0 4.8-7 11-7 11z"/><circle cx="12" cy="10" r="2.5"/></svg>{location}</span>'
                    )
                if salary:
                    meta_parts.append(
                        f'<span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">'
                        f'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>{salary}</span>'
                    )
                if date:
                    meta_parts.append(
                        f'<span><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">'
                        f'<rect x="3" y="4" width="18" height="17" rx="3"/><path d="M8 2v4M16 2v4M3 9h18"/></svg>{date}</span>'
                    )
                meta_html = "".join(meta_parts)

                # 卡片 HTML 压成单行、无缩进，避免被 Markdown 当成代码块原样显示
                card = (
                    f'<article class="mk-job">'
                    f'<div class="mk-top"><div class="mk-name">{company} — {job}</div>'
                    f'<div class="mk-rank">No.{rank}</div></div>'
                    f'<div class="mk-reason">{reason}</div>'
                    f'<div class="mk-meta">{meta_html}</div>'
                    f'<div class="mk-score"><div class="mk-bar"><i style="width:{score_num}%"></i></div>'
                    f'<div class="mk-num">{score_text}</div></div>'
                    f'</article>'
                )
                cards_html.append(card)

            jobs_html = (
                '<div class="mk-card"><div class="mk-card-title">推荐岗位明细</div>'
                '<div class="mk-list">' + "".join(cards_html) + "</div></div>"
            )

            # 底部横幅 + 主按钮
            banner_html = (
                '<div class="mk-banner">'
                '<span class="mk-badge"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 15.3 6.8 18.1l1-5.8L3.5 8.2l5.9-.9z"/></svg></span>'
                '<div><div class="b-txt">建议优先投递排名靠前的岗位</div>'
                '<div class="b-sub">投递前请结合岗位地点、薪资与经验要求综合判断</div></div></div>'
                '<button class="mk-btn" type="button" onclick="alert(\'演示按钮：接入投递动作后在此绑定\')">查看完整岗位列表</button>'
                "</div>"
            )

            st.markdown(hero_html + info_html + jobs_html + banner_html, unsafe_allow_html=True)

            st.download_button(
                "⬇️ 下载匹配结果 CSV",
                data=result_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="AI匹配结果.csv",
                mime="text/csv",
            )
        else:
            st.info("💡 暂无卡片数据。请先粘贴简历并点击「🚀 开始 AI 智能匹配」；若模型未返回结构化数据，可在文字版查看结论。")
