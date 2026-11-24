import streamlit as st
import pandas as pd
import os
import json
import re
import random
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

def _chart_bundle(result_df: pd.DataFrame) -> dict:
    """从结果表组装图表所需数据：条形图（岗位-得分）与环形图（得分段占比）。"""
    labels, values = [], []
    high = mid = low = 0
    for _, row in result_df.iterrows():
        name = f"{row.get('企业名称', '')}·{row.get('招聘岗位', '')}"
        name = str(name)[:14]
        score = row.get("匹配得分", "")
        score_num = int(score) if score != "" and score is not None else 0
        labels.append(name)
        values.append(score_num)
        if score_num >= 80:
            high += 1
        elif score_num >= 60:
            mid += 1
        else:
            low += 1
    return {
        "labels": labels,
        "values": values,
        "segments": [
            {"name": "高匹配(≥80)", "value": high, "color": "#5f9e3f"},
            {"name": "中匹配(60-79)", "value": mid, "color": "#f0a93b"},
            {"name": "待提升(<60)", "value": low, "color": "#e57373"},
        ],
    }

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

                    # 本次随机匹配岗位数量：2~5 个
                    top_n = random.randint(2, 5)

                    # 构建给 AI 的指令：要求同时给出文字结论和结构化 JSON，便于下方生成表格
                    prompt = f"""
                    你是一个专业的 HR 招聘助手。请根据以下候选人的简历内容，对比给出的岗位列表，评估匹配度（0-100分），并按得分从高到低排序。

                    【候选人简历】
                    {resume_text}

                    【岗位列表】
                    {jobs_text}

                    请先简要说明整体匹配结论，然后输出匹配度最高的前 {top_n} 个岗位（数量必须为 {top_n} 个，不要多给也不要少给）。
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
                    if items:
                        items = items[:top_n]  # 兜底：按本次随机数量截断
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
        # ---------- 页面二：表格版一览表（Streamlit 原生组件 + 图表） ----------
        if result_df is not None and not result_df.empty:
            scores = [s for s in result_df["匹配得分"].tolist() if s != "" and s is not None]
            total = len(result_df)
            top_score = max(scores) if scores else 0
            avg_score = round(sum(scores) / len(scores)) if scores else 0
            high_cnt = next((s["value"] for s in _chart_bundle(result_df)["segments"] if s["name"].startswith("高")), 0)
            chart = _chart_bundle(result_df)

            # 原生指标卡一行四列
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("推荐岗位", f"{total} 个")
            m2.metric("最高得分", f"{top_score} 分")
            m3.metric("平均得分", f"{avg_score} 分")
            m4.metric("高匹配岗位", f"{high_cnt} 个")

            st.divider()

            # 图表区：左条形图、右环形图（ECharts 嵌入）
            chart_json = json.dumps(chart, ensure_ascii=False)
            charts_html = f"""
            <div style="display:grid;grid-template-columns:3fr 2fr;gap:16px;">
              <div style="background:#fff;border:1px solid #eee;border-radius:14px;padding:14px;">
                <div style="font-size:15px;font-weight:600;margin-bottom:8px;">各岗位匹配得分对比</div>
                <div id="barChart" style="width:100%;height:300px;"></div>
              </div>
              <div style="background:#fff;border:1px solid #eee;border-radius:14px;padding:14px;">
                <div style="font-size:15px;font-weight:600;margin-bottom:8px;">匹配程度构成</div>
                <div id="pieChart" style="width:100%;height:300px;"></div>
              </div>
            </div>
            <script src="https://fastly.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
            <script>
              var data = {chart_json};
              var bar = echarts.init(document.getElementById('barChart'));
              bar.setOption({{
                grid: {{left: 10, right: 34, top: 10, bottom: 10, containLabel: true}},
                tooltip: {{trigger: 'axis', axisPointer: {{type: 'shadow'}}}},
                xAxis: {{type: 'value', max: 100, axisLabel: {{color: '#6b7280'}}, splitLine: {{lineStyle: {{color: '#eee'}}}}}},
                yAxis: {{type: 'category', data: data.labels.slice().reverse(), axisLabel: {{color: '#374151', fontSize: 11}}, axisLine: {{show: false}}}},
                series: [{{type: 'bar', data: data.values.slice().reverse(), barWidth: 14, itemStyle: {{color: '#5f9e3f', borderRadius: [0, 7, 7, 0]}}, label: {{show: true, position: 'right', color: '#4a8230', fontWeight: 600}}}}]
              }});
              var pie = echarts.init(document.getElementById('pieChart'));
              pie.setOption({{
                tooltip: {{trigger: 'item'}},
                legend: {{bottom: 0, textStyle: {{color: '#374151', fontSize: 12}}}},
                series: [{{type: 'pie', radius: ['42%', '68%'], center: ['50%', '44%'],
                  itemStyle: {{borderColor: '#fff', borderWidth: 2}},
                  label: {{formatter: '{{b}}\\n{{c}} 个', color: '#374151', fontSize: 12}},
                  data: data.segments.map(function(s){{return {{name: s.name, value: s.value, itemStyle: {{color: s.color}}}};}})
                }}]
              }});
              window.addEventListener('resize', function(){{bar.resize(); pie.resize();}});
            </script>
            """
            from streamlit.components.v1 import html as st_html
            st_html(charts_html, height=360)

            st.divider()
            st.subheader("📌 匹配岗位明细")

            # 原生列布局岗位卡片（避免宽表文字截断、列宽死板）
            for _, row in result_df.iterrows():
                rank = row.get("排名", "")
                company = str(row.get("企业名称", "") or "")
                job = str(row.get("招聘岗位", "") or "")
                reason = str(row.get("匹配理由", "") or "")
                location = str(row.get("工作地点", "") or "")
                salary = str(row.get("薪资福利", "") or "")
                date = str(row.get("更新日期", "") or "")

                score = row.get("匹配得分", "")
                score_num = int(score) if score != "" and score is not None else 0
                score_text = f"{score_num} 分" if score != "" and score is not None else "—"

                with st.container(border=True):
                    c1, c2 = st.columns([5, 1])
                    with c1:
                        st.markdown(f"**{rank}. {company} — {job}**")
                        if reason:
                            st.caption(reason)
                        meta = "　｜　".join(
                            part for part in [
                                f"📍 {location}" if location else "",
                                f"💰 {salary}" if salary else "",
                                f"🗓 {date}" if date else "",
                            ] if part
                        )
                        if meta:
                            st.caption(meta)
                    with c2:
                        st.metric("匹配得分", score_text)

            st.download_button(
                "⬇️ 下载匹配结果 CSV",
                data=result_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="AI匹配结果.csv",
                mime="text/csv",
            )
        else:
            st.info("💡 暂无表格数据。请先粘贴简历并点击「🚀 开始 AI 智能匹配」；若模型未返回结构化数据，可在文字版查看结论。")
