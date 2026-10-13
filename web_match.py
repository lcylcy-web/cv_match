import streamlit as st
import pandas as pd
import os
from openai import OpenAI

st.set_page_config(page_title="简历岗位匹配工具", page_icon="📄", layout="wide")
st.title("📄 简历与岗位智能匹配 ")

# ================= 侧边栏配置 =================
st.sidebar.header("🔑 配置大模型")

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
    st.sidebar.markdown("[获取 DeepSeek API Key](https://platform.deepseek.com/)")
    default_base_url = "https://api.deepseek.com"
    default_model = "deepseek-chat"
elif model_choice == "豆包 (Doubao)":
    st.sidebar.markdown("[获取火山方舟 API Key](https://console.volcengine.com/ark/)")
    default_base_url = "https://ark.cn-beijing.volces.com/api/v3"
    default_model = "doubao-seed-2.0-lite-260428"

# 自动读取 Secrets 中的 API_KEY
api_key = st.sidebar.text_input("请输入 API Key：", value=st.secrets.get("API_KEY", ""), type="password")
model_name = st.sidebar.text_input("模型名称 / 接入点ID：", value=default_model)

# ================= 主页面 =================
# 读取 Excel
def load_data():
    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jobs.xlsx")
    if os.path.exists(excel_path):
        return pd.read_excel(excel_path)
    return None

df = load_data()

if df is None:
    st.error("❌ 未找到 jobs.xlsx，请确保文件在项目目录下！")
else:
    st.success(f"✅ 已成功加载 {len(df)} 个本地岗位数据！")
    
    resume_text = st.text_area("✍️ 请将简历内容（纯文本）粘贴到这里：", height=250)
    
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

                    # 构建给 AI 的指令
                    prompt = f"""
                    你是一个专业的 HR 招聘助手。请根据以下候选人的简历内容，对比给出的岗位列表，评估匹配度（0-100分），并按得分从高到低排序。

                    【候选人简历】
                    {resume_text}

                    【岗位列表】
                    {jobs_text}

                    请直接输出匹配度最高的前 5 个岗位，格式要求：
                    1. 【企业名称 - 招聘岗位】匹配得分：XX分
                    2. 匹配理由：结合候选人的学历、专业、经验，说明为什么匹配该岗位。
                    3. 岗位详情：包括工作地点、薪资福利、更新日期。
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

                    # 展示结果
                    st.subheader(f"🎯 {model_choice} 智能匹配报告")
                    st.markdown(response.choices[0].message.content)
                    
                except Exception as e:
                    st.error(f"❌ 调用 AI 失败，错误信息：{e}")
                    st.info("💡 常见原因：1. API Key 无效；2. 模型名称/接入点ID 填错；3. 账号余额不足。")
