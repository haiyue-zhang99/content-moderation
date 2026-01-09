# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import jieba
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import math
from streamlit_option_menu import option_menu

# 页面设置
st.set_page_config(
    page_title="🧑‍💻 内容审核数据统计工具",
    layout="wide"
)

# 顶部标题与美化
st.markdown("""
    <div style="background: linear-gradient(90deg, #F0F8FF 0%, #E6E6FA 100%);
                padding: 18px 0; border-radius: 12px; margin-bottom: 18px;">
        <h1 style="color: #725e82; text-align: center; font-size: 40px; letter-spacing: 2px;">
            内容审核数据统计工具
        </h1>
    </div>
""", unsafe_allow_html=True)

# 初始化 session_state
for key in ["uploaded_file1", "df1", "uploaded_file2", "df2"]:
    if key not in st.session_state:
        st.session_state[key] = None


# =========================
# 列名规范化与文件加载（修复 BOM）
# =========================
def normalize_col(c: object) -> str:
    """
    规范化单个列名用于校验（不影响原列名显示）：
    - 转为字符串并 strip 前后空白
    - 去掉 UTF-8 BOM 前缀 \uFEFF
    - 统一去掉全角空格（\u3000）
    - 若整列被一对中/英文括号包裹，去掉外层括号
    - 返回规范化后的列名
    """
    s = str(c)
    s = s.strip()
    # 去 BOM
    if s.startswith("\ufeff"):
        s = s.lstrip("\ufeff")
    # 去全角空格
    s = s.replace("\u3000", " ")
    # 去外层括号（中/英）
    if (s.startswith("（") and s.endswith("）")) or (s.startswith("(") and s.endswith(")")):
        s = s[1:-1].strip()
    return s

def load_dataframe(uploaded_file, required_cols=None, sheet_engine="openpyxl"):
    """
    根据扩展名读取 DataFrame，并进行必需列宽松校验：
    - CSV：优先 utf-8；失败自动回退 gbk；sep=None, engine="python" 自适应分隔符
    - XLSX：openpyxl 引擎
    - 列名校验：对 required_cols 与 df.columns 做 normalize + lower 后比对
      （大小写不敏感、去 BOM、去空白）
    - 成功返回 DataFrame；失败抛出异常（ValueError / RuntimeError）
    """
    if uploaded_file is None:
        return None

    name = uploaded_file.name.lower().strip()
    df = None

    try:
        if name.endswith(".csv"):
            # 先尝试 utf-8
            try:
                df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="utf-8")
            except Exception:
                # 回退 gbk
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=None, engine="python", encoding="gbk")
        elif name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine=sheet_engine)
        else:
            raise ValueError("不支持的文件类型，请上传 .xlsx 或 .csv。")

        # 保留原始列名用于展示
        original_cols = list(df.columns)

        # 生成“规范化小写版”列集合用于校验
        normalized_set = {normalize_col(c).lower() for c in df.columns}

        # 校验必需列
        if required_cols:
            missing = []
            for rc in required_cols:
                key = normalize_col(rc).lower()
                if key not in normalized_set:
                    missing.append(rc)

            if missing:
                # 提示里展示实际原始列名，帮助定位
                raise ValueError(
                    f"上传文件不符合要求，缺少以下字段：{', '.join(missing)}；"
                    f"文件实际列名为：{', '.join(map(str, original_cols))}"
                )

        # 额外：把列名进行“去空白+去BOM”的轻度修正，以便后续代码按原逻辑访问
        # 注意：仅去空白和BOM，不改大小写，不改原始语义
        df.columns = [normalize_col(c) for c in df.columns]

        return df

    except ValueError as ve:
        raise ve
    except Exception as e:
        raise RuntimeError(f"读取文件失败：{e}")

# 滑动标签导航
selected = option_menu(
    None,
    ["审核数据统计", "编辑加分统计"],
    icons=["bar-chart", "award"],
    orientation="horizontal",
    styles={
        "container": {"padding": "0!important", "background-color": "#fafafa"},
        "icon": {"color": "#9999CC", "font-size": "18px"},
        "nav-link": {
            "font-size": "18px",
            "text-align": "center",
            "margin": "0px",
            "--hover-color": "#eee",
        },
        "nav-link-selected": {"background-color": "#9999CC", "color": "white"},
    }
)

if selected == "审核数据统计":
    # 内容区美化卡片
    st.markdown("""
        <div style="
            background: #fff;
            border-radius: 18px;
            box-shadow: 0 4px 16px rgba(79,139,249,0.10);
            padding: 42px 38px 34px 38px;
            margin-bottom: 32px;
            border: 1px solid #f0f4fa;
            ">
            <div style="background: linear-gradient(90deg, #99CCCC 0%, #9999CC 100%);
                        border-radius: 12px; padding: 12px 0; margin-bottom: 18px;">
                <h2 style="color: white; text-align: center; font-size: 1.6rem; letter-spacing: 2px; margin:0;">
                    📊 审核数据统计
                </h2>
            </div>
            <p style="font-size: 1.15rem; color: #555; text-align:center; margin-bottom:18px;">
                上传审核数据，平台即刻自动执行多维度统计分析，并生成可视化成果。
            </p>
        </div>
    """, unsafe_allow_html=True)
    uploaded_file1 = st.file_uploader("📂 Upload file（.xlsx 或 .csv）", type=["xlsx", "csv"], key="file1")

    # 每次上传新文件时，先清空旧状态
    if uploaded_file1 is not None:
        try:
            required_cols_1 = ['Requester', 'RankList', 'Action', 'Author', 'ProviderName', 'Reason', 'Comment']
            df1 = load_dataframe(uploaded_file1, required_cols=required_cols_1)
            st.session_state["df1"] = df1
            st.session_state["uploaded_file1"] = uploaded_file1
            st.success("✅ 数据已保存，切换页面不会丢失！")
        except ValueError as ve:
            st.session_state["df1"] = None
            st.error(f"❌ {ve}")
        except Exception as e:
            st.session_state["df1"] = None
            st.error(f"❌ 上传文件不符合要求，请重新上传。错误详情：{e}")
    else:
        st.info("请上传脱敏审核数据以开始分析。")

    # 只有 df1 有效时才继续后续统计与展示
    if st.session_state["df1"] is not None:
        df1 = st.session_state["df1"]


        # 1. 审核列表统计
        st.markdown("""
            <div style="
                background: #f7faff;
                border-radius: 12px;
                padding: 18px 20px;
                margin-bottom: 18px;
                border-left: 4px solid #6699CC;
                ">
                <h4 style="color: #6699CC; margin:0 0 8px 0;">📋 第一项：审核列表统计</h4>
                <p style="color: #666; font-size: 1rem; margin:0;">
                    快速呈现各审核人员的列表审核数量。
                </p>
            </div>
        """, unsafe_allow_html=True)
        email_order = [
            "v-qingqinghe@microsoft.com", "v-yangyang5@microsoft.com", "v-qiangwei@microsoft.com", "v-cwen@microsoft.com",
            "v-yuehan@microsoft.com", "v-shuagao@microsoft.com", "v-xiyuan1@microsoft.com", "v-xuelyang@microsoft.com",
            "v-wenlchen@microsoft.com", "v-huiwwang@microsoft.com", "v-dandanli@microsoft.com", "v-yuanjunli@microsoft.com",
            "v-yuqincheng@microsoft.com", "v-jiaozhang@microsoft.com", "v-wangjua@microsoft.com", "v-qingkzhang@microsoft.com",
            "v-minshi1@microsoft.com", "v-hengma@microsoft.com", "v-yingxli@microsoft.com", "v-peilirao@microsoft.com",
            "v-jiangqia@microsoft.com", "v-chaozhao@microsoft.com", "v-junlanli@microsoft.com", "v-tengpan@microsoft.com",
            "v-qinjiang@microsoft.com", "v-yancche@microsoft.com", "v-jiangli@microsoft.com", "v-yuleiwu@microsoft.com",
            "v-liuyang@microsoft.com", "v-weihu1@microsoft.com", "v-haizhang@microsoft.com", "v-xiaoqingwu@microsoft.com",
            "v-lixren@microsoft.com",
        ]
        stat1 = []
        for email in email_order:
            user_data = df1[df1['Requester'] == email]
            simple_count = user_data[user_data['RankList'].str.contains('简单列表', na=False)].shape[0]
            general_quality_count = user_data[user_data['RankList'].str.contains('一般列表|优质列表|视频列表', na=False)].shape[0]
            stat1.append({
                '审核人员': email,
                '简单列表数量': simple_count,
                '其他列表数量': general_quality_count
            })
        stat1_df1 = pd.DataFrame(stat1)
        st.markdown('<hr style="border:1px solid #e3eaf2;">', unsafe_allow_html=True)
        st.dataframe(stat1_df1.style.set_properties(**{'font-size': '12px'}), use_container_width=True)
        st.markdown('<hr style="border:1px solid #e3eaf2;">', unsafe_allow_html=True)

        # 2. 审核人员表现分析
        st.markdown("""
            <div style="
                background: #f7faff;
                border-radius: 12px;
                padding: 18px 20px;
                margin-bottom: 18px;
                border-left: 4px solid #669966;
                ">
                <h4 style="color: #669966; margin:0 0 8px 0;">📈 第二项：审核人员表现分析总结</h4>
                <p style="color: #666; font-size: 1rem; margin:0;">
                    梳理表现突出或异常的审核人员。
                </p>
            </div>
        """, unsafe_allow_html=True)
        efficiency_data = []
        for email in email_order:
            user_data = df1[df1['Requester'] == email]
            total = user_data.shape[0]
            rejected = user_data[user_data['Action'] == 'Rejected'].shape[0]
            approved = user_data[user_data['Action'] == 'Approved'].shape[0]
            approval_rate = round(approved / total * 100, 2) if total > 0 else 0
            reject_rate = round(rejected / total * 100, 2) if total > 0 else 0
            efficiency_data.append({
                'email': email,
                'total': total,
                'approval_rate': approval_rate,
                'reject_rate': reject_rate
            })
        sorted_by_total = sorted(efficiency_data, key=lambda x: x['total'], reverse=True)
        top_reviewers = sorted_by_total[:3]
        high_rejectors = [x for x in efficiency_data if x['reject_rate'] > 20 and x['reject_rate'] < 100]
        low_rejectors = [x for x in efficiency_data if x['reject_rate'] < 5 and x['reject_rate'] > 0]
        st.success("🏆 审核量最高的前三名审核人员：")
        for person in top_reviewers:
            st.markdown(f"- **{person['email']}**：共审核了 **{person['total']}** 条内容，拒绝率为 **{person['reject_rate']}%**")
        if high_rejectors:
            st.warning("⚠️ 以下审核人员的拒绝率偏高（超过 20%）：")
            for person in high_rejectors:
                st.markdown(f"- **{person['email']}**：拒绝率为 **{person['reject_rate']}%**")
        if low_rejectors:
            st.info("⚠️ 以下审核人员的拒绝率偏低（低于 5%）：")
            for person in low_rejectors:
                st.markdown(f"- **{person['email']}**：拒绝率为 **{person['reject_rate']}%**")
        if not high_rejectors and not low_rejectors:
            st.markdown("👍 所有审核人员的拒绝率均在合理范围内。")

        # 3. 被拒次数最多的作者 Top 20
        st.markdown("""
            <div style="
                background: #f7faff;
                border-radius: 12px;
                padding: 18px 20px;
                margin-bottom: 18px;
                border-left: 4px solid #9999CC;
                ">
                <h4 style="color: #9999CC; margin:0 0 8px 0;">📌 第三项：被拒次数最多的作者 Top 20</h4>
                <p style="color: #666; font-size: 1rem; margin:0;">
                    统计被拒次数最多的作者排名。
                </p>
            </div>
        """, unsafe_allow_html=True)
        rejected_authors = df1[df1['Action'] == 'Rejected']['Author'].value_counts().head(20).reset_index()
        rejected_authors.columns = ['作者', '被拒次数']
        st.dataframe(rejected_authors.style.set_properties(**{'font-size': '12px'}), use_container_width=True)

        # 4. 品牌审核结果统计
        st.markdown("""
            <div style="
                background: #f7faff;
                border-radius: 12px;
                padding: 18px 20px;
                margin-bottom: 18px;
                border-left: 4px solid #6699CC;
                ">
                <h4 style="color: #6699CC; margin:0 0 8px 0;">🏷️ 第四项：品牌审核结果统计（审核量超过500）</h4>
                <p style="color: #666; font-size: 1rem; margin:0;">
                    统计重要品牌的下线原因及其分布情况。
                </p>
            </div>
        """, unsafe_allow_html=True)
        provider_counts = df1['ProviderName'].value_counts()
        popular_providers = provider_counts[provider_counts > 500].index.tolist()
        summary = {}
        for provider in popular_providers:
            provider_data = df1[df1['ProviderName'] == provider]
            total = provider_data.shape[0]
            reason_counts = provider_data['Reason'].fillna('通过').value_counts()
            reason_percentages = (reason_counts / total * 100).round(2)
            reason_summary = {}
            other_percentage = 0.0
            for reason, percentage in reason_percentages.items():
                if percentage <= 1:
                    other_percentage += percentage
                else:
                    reason_summary[reason] = percentage
            if other_percentage > 0:
                reason_summary['另外'] = round(other_percentage, 2)
            summary[provider] = reason_summary
        all_reasons = sorted(set(r for reasons in summary.values() for r in reasons))
        table_data = []
        for provider, reasons in summary.items():
            row = {'ProviderName': provider}
            for reason in all_reasons:
                row[reason] = reasons.get(reason, 0.0)
            table_data.append(row)
        result_df1 = pd.DataFrame(table_data).set_index('ProviderName')
        st.dataframe(result_df1.style.set_properties(**{'font-size': '12px'}).format("{:.2f}%"), use_container_width=True)

        # 5. 品牌模糊查询
        st.markdown("""
            <div style="
                background: #f7faff;
                border-radius: 12px;
                padding: 18px 20px;
                margin-bottom: 18px;
                border-left: 4px solid #669966;
                ">
                <h4 style="color: #669966; margin:0 0 8px 0;">🔍 第五项：品牌查询</h4>
                <p style="color: #666; font-size: 1rem; margin:0;">
                    输入品牌关键词进行定向查询。
                </p>
            </div>
        """, unsafe_allow_html=True)
        all_providers = df1['ProviderName'].dropna().unique()
        input_text = st.text_input("请至少输入两个字符")
        if input_text and len(input_text) >= 2:
            matching_brands = [brand for brand in all_providers if input_text.lower() in brand.lower()]
            if matching_brands:
                for selected_brand in matching_brands:
                    provider_data = df1[df1['ProviderName'] == selected_brand]
                    total = provider_data.shape[0]
                    reason_counts = provider_data['Reason'].fillna('通过').value_counts()
                    reason_percentages = (reason_counts / total * 100).round(2)
                    reason_summary = {}
                    other_percentage = 0.0
                    for reason, percentage in reason_percentages.items():
                        if percentage <= 1:
                            other_percentage += percentage
                        else:
                            reason_summary[reason] = percentage
                    if other_percentage > 0:
                        reason_summary['另外'] = round(other_percentage, 2)
                    st.markdown(f"### 品牌：{selected_brand} 的审核结果占比")
                    for reason, percentage in reason_summary.items():
                        st.markdown(f"- {reason}：{percentage}%")
            else:
                st.markdown("❌ 未找到匹配的品牌名称，请尝试其他输入。")

        # 6. 编辑备注关键词词云图
        st.markdown("""
            <div style="
                background: #f7faff;
                border-radius: 12px;
                padding: 18px 20px;
                margin-bottom: 18px;
                border-left: 4px solid #9999CC;
                ">
                <h4 style="color: #9999CC; margin:0 0 8px 0;">💡 第六项：编辑备注词云图</h4>
                <p style="color: #666; font-size: 1rem; margin:0;">
                    展示审核备注的关键词。
                </p>
            </div>
        """, unsafe_allow_html=True)
        stopwords = set([
            "的", "了", "和", "是", "我", "也", "就", "都", "而", "及", "与", "着", "或", "一个", "没有", "我们", "你", "他", "她", "它",
            "在", "上", "下", "中", "为", "对", "不", "这", "那", "：", "，", "。", "！", "？", "（", "）", "、", "；", "“", "”", "‘", "’",
            "【", "】", "《", "》", " ", "\n", "\r", "\t"
        ])
        comments = df1['Comment'].dropna().astype(str).tolist()
        words = []
        for comment in comments:
            seg_list = jieba.lcut(comment)
            filtered_words = [word for word in seg_list if word not in stopwords and len(word.strip()) > 1]
            words.extend(filtered_words)
        word_freq = Counter(words)
        if word_freq:
            wc = WordCloud(font_path="MSYH.TTC", width=800, height=400, background_color="white")
            wc.generate_from_frequencies(word_freq)
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            st.pyplot(fig)
        else:
            st.info("暂无可生成词云的备注内容。")

elif selected == "编辑加分统计":
    # 内容区美化卡片
    st.markdown("""
        <div style="
            background: #fff;
            border-radius: 18px;
            box-shadow: 0 4px 16px rgba(79,139,249,0.10);
            padding: 42px 38px 34px 38px;
            margin-bottom: 32px;
            border: 1px solid #f0f4fa;
            ">
            <div style="background: linear-gradient(90deg, #9999CC 0%, #99CC99 100%);
                        border-radius: 12px; padding: 12px 0; margin-bottom: 18px;">
                <h2 style="color: white; text-align: center; font-size: 1.6rem; letter-spacing: 2px; margin:0;">
                    📈 编辑加分统计
                </h2>
            </div>
            <p style="font-size: 1.15rem; color: #555; text-align: center; margin-bottom:18px;">
上传编辑审核数据，加分结果即刻自动计算呈现。
            </p>
        </div>
    """, unsafe_allow_html=True)

    uploaded_file2 = st.file_uploader("📂 Upload file（.xlsx）", type=["xlsx"], key="file2")
    if uploaded_file2 is not None:
        try:
            df2 = pd.read_excel(uploaded_file2, engine='openpyxl')
            required_cols = ['审核人员']
            # 动态判断所有天的“简单列表数量”等列是否存在
            day_cols = [col for col in df2.columns if "简单列表数量" in col or "一般+优质列表数量" in col]
            if not day_cols:
                st.session_state["df2"] = None
                st.error("❌ 上传文件不符合要求，缺少“简单列表数量”或“一般+优质列表数量”相关字段，请重新上传。")
            elif any(col not in df2.columns for col in required_cols):
                st.session_state["df2"] = None
                st.error("❌ 上传文件不符合要求，缺少“审核人员”字段，请重新上传。")
            else:
                st.session_state["uploaded_file2"] = uploaded_file2
                st.session_state["df2"] = df2
                st.success("✅ 数据已保存，切换页面不会丢失！")
        except Exception as e:
            st.session_state["df2"] = None
            st.error("❌ 上传文件不符合要求，请重新上传。")
    else:
        st.info("请上传编辑审核数据以开始分析。")

    if st.session_state["df2"] is not None:
        df2 = st.session_state["df2"]
        days = []
        for col in df2.columns:
            if "简单列表数量" in col:
                day = col.replace("简单列表数量", "")
                days.append(day)
        days = list(dict.fromkeys(days))  # 去重保序
        results = []

        def is_valid_number(x):
            try:
                val = float(x)
                if math.isnan(val) or val == 0:
                    return False
                return True
            except (ValueError, TypeError):
                return False

        def safe_float(x):
            try:
                return float(x)
            except (ValueError, TypeError):
                return None

        for _, row in df2.iterrows():
            name = row['审核人员']
            row_result = {'审核人员': name}
            total_score = 0
            for day in days:
                simple_qty = safe_float(row.get(f"{day}简单列表数量", None))
                simple_time = safe_float(row.get(f"{day}简单列表时长", None))
                complex_qty = safe_float(row.get(f"{day}一般+优质列表数量", None))
                complex_time = safe_float(row.get(f"{day}一般+优质列表时长", None))
                simple_valid = is_valid_number(simple_qty) and is_valid_number(simple_time)
                complex_valid = is_valid_number(complex_qty) and is_valid_number(complex_time)
                simple_avg = simple_qty / simple_time if simple_valid else None
                complex_avg = complex_qty / complex_time if complex_valid else None
                score = 0
                if complex_valid and not simple_valid:
                    if complex_avg >= 160:
                        score = 1
                    elif complex_avg < 153.75:
                        score = -1
                elif simple_valid and not complex_valid:
                    if simple_avg >= 348:
                        score = 1
                    elif simple_avg < 318:
                        score = -1
                elif simple_valid and complex_valid:
                    if simple_avg >= 348 and complex_avg >= 160:
                        score = 1
                    elif simple_avg < 318 or complex_avg < 153.75:
                        score = -1
                total_score += score
                row_result[f'{day}简单列表时均'] = round(simple_avg, 2) if simple_avg is not None else ''
                row_result[f'{day}一般+优质列表时均'] = round(complex_avg, 2) if complex_avg is not None else ''
                row_result[f'{day}加扣分'] = score
            row_result['总分'] = total_score
            results.append(row_result)
        wide_df = pd.DataFrame(results)

        def color_score(val):
            if val == 1:
                return 'color: #228B22; font-weight: bold;'  # 草绿色
            elif val == -1:
                return 'color: #CC0000; font-weight: bold;'  # 深红色
            return ''

        score_cols = [col for col in wide_df.columns if '加扣分' in col]
        st.dataframe(wide_df.style.applymap(color_score, subset=score_cols).set_properties(**{'font-size': '13px'}), use_container_width=True)

        # 下载按钮
        st.markdown("#### ⬇️ 下载周度横向宽表加分结果")
        wide_df.to_excel("周度横向宽表加分结果.xlsx", index=False)
        with open("周度横向宽表加分结果.xlsx", "rb") as file:
            st.download_button("⬇️ 点此下载", data=file, file_name="周度横向宽表加分结果.xlsx")

# 页面底部
st.markdown("---")
st.markdown("""
    <div style="text-align:center; color: #888; font-size: 0.95rem;">
        © 2025 内容审核数据统计工具 <br>
        Powered by Streamlit
    </div>
""", unsafe_allow_html=True)

