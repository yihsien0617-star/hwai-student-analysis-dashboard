"""
🎓 大學招生分析儀表板
讓各系主任上傳資料即可自動產生五大招生分析報告
"""
import streamlit as st
import pandas as pd
from io import BytesIO

from utils import (
    DataLoader,
    GeoAnalyzer,
    FeederSchoolAnalyzer,
    ChannelAnalyzer,
    ProfileAnalyzer,
    TrendAnalyzer,
)

# ══════════════════════════════════════════════
# 頁面設定
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="招生分析儀表板",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 自訂 CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: bold;
        color: #1E3A5F;
        text-align: center;
        padding: 1rem 0;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .insight-box {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #667eea;
        margin-bottom: 0.5rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        font-size: 1rem;
    }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# 工具函數
# ══════════════════════════════════════════════
def create_download_button(df: pd.DataFrame, filename: str, label: str):
    """產生 Excel 下載按鈕"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=True, sheet_name="分析結果")
    output.seek(0)

    st.download_button(
        label=f"📥 {label}",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def display_insights(insights: list, title: str = "💡 關鍵發現"):
    """顯示洞察列表"""
    st.subheader(title)
    for insight in insights:
        st.markdown(
            f'<div class="insight-box">{insight}</div>',
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════
# 側邊欄
# ══════════════════════════════════════════════
with st.sidebar:
    st.image(
        "https://img.icons8.com/color/96/000000/graduation-cap.png",
        width=80,
    )
    st.title("招生分析儀表板")
    st.markdown("---")

    # 檔案上傳
    st.subheader("📂 上傳資料")
    uploaded_file = st.file_uploader(
        "請上傳新生資料 Excel 檔",
        type=["xlsx", "xls"],
        help="支援 .xlsx 及 .xls 格式，可包含多個工作表",
    )

    st.markdown("---")
    st.subheader("📋 必要欄位")
    st.markdown("""
    | 欄位 | 說明 |
    |------|------|
    | 入學學年 | 如：111、112、113 |
    | 班級名稱 | 系所名稱 |
    | 學號 | 唯一識別碼 |
    | 區域 | 北/中/南/東/離島 |
    | 縣市 | 如：台北市、高雄市 |
    | 入學前學歷 | 高中/高職/綜合高中 |
    | 畢業學校 | 來源學校名稱 |
    | 畢業科系 | 如：普通科、護理科 |
    | 入學方式 | 如：申請入學、分發 |
    """)

    st.markdown("---")
    st.caption("© 2024 招生分析系統 v1.0")


# ══════════════════════════════════════════════
# 主頁面
# ══════════════════════════════════════════════
st.markdown(
    '<div class="main-header">🎓 招生資料分析儀表板</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-header">'
    "上傳新生資料，自動產生五大分析報告，"
    "為招生策略提供數據支持"
    "</div>",
    unsafe_allow_html=True,
)

if uploaded_file is None:
    # ── 未上傳檔案：顯示說明頁 ──
    st.info("👈 請從左側上傳 Excel 檔案以開始分析")

    with st.expander("📖 使用說明", expanded=True):
        st.markdown("""
        ### 系統功能

        本系統提供 **五大分析面向**：

        | # | 分析面向 | 內容說明 |
        |---|---------|---------|
        | 1 | 🗺️ 生源地理分析 | 學生來自哪些區域/縣市？趨勢如何？ |
        | 2 | 🏫 餵校分析 | 哪些高中職是穩定的學生來源校？ |
        | 3 | 🎯 入學管道分析 | 各招生管道的成效與風險評估 |
        | 4 | 🎓 學生背景分析 | 高中 vs 高職？普通科 vs 專業科？ |
        | 5 | 📈 跨年度趨勢 | 綜合趨勢與多元性指標追蹤 |

        ### 操作步驟

        1. 準備包含必要欄位的 Excel 檔案
        2. 從左側上傳檔案
        3. 選擇要分析的系所與學年範圍
        4. 瀏覽各頁籤的分析結果
        5. 可下載分析報表
        """)

    st.stop()


# ══════════════════════════════════════════════
# 資料載入與驗證
# ══════════════════════════════════════════════
loader = DataLoader()

with st.spinner("📊 正在載入資料..."):
    raw_df = loader.load_excel(uploaded_file)

if raw_df.empty:
    st.error("無法讀取檔案，請確認檔案格式正確")
    st.stop()

# 驗證欄位
validation = loader.validate_columns(raw_df)

if not validation["is_valid"]:
    st.error("❌ 資料欄位不完整！")
    st.warning(f"缺少以下欄位：{', '.join(validation['missing_columns'])}")
    st.info("已找到的欄位：" + ", ".join(validation["found_columns"]))
    st.dataframe(raw_df.head())
    st.stop()

# 清洗資料
df = loader.clean_data_pipeline(raw_df)
stats = loader.get_summary_stats(df)

# ── 篩選器 ──
st.markdown("---")
col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    departments = loader.get_department_list(df)
    selected_depts = st.multiselect(
        "🏛️ 選擇系所（可多選，空白 = 全部）",
        options=departments,
        default=[],
    )

with col_filter2:
    year_min, year_max = loader.get_year_range(df)
    selected_years = st.slider(
        "📅 選擇學年範圍",
        min_value=year_min,
        max_value=year_max,
        value=(year_min, year_max),
    )

# 套用篩選
filtered_df = loader.filter_data(
    df,
    departments=selected_depts if selected_depts else None,
    years=selected_years,
)

# ── KPI 總覽 ──
st.markdown("---")
st.subheader("📊 資料總覽")

kpi_cols = st.columns(6)
kpi_items = [
    ("👥 總學生數", stats["總學生數"]),
    ("📅 涵蓋學年", stats["涵蓋學年數"]),
    ("🏛️ 系所數", stats["系所數"]),
    ("🏙️ 來源縣市", stats["來源縣市數"]),
    ("🏫 來源學校", stats["來源學校數"]),
    ("🎯 入學管道", stats["入學管道數"]),
]

for col, (label, value) in zip(kpi_cols, kpi_items):
    col.metric(label=label, value=value)

st.caption(
    f"目前篩選後資料：**{len(filtered_df)}** 筆 "
    f"（{selected_years[0]}~{selected_years[1]} 學年度）"
)

# ══════════════════════════════════════════════
# 五大分析頁籤
# ══════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🗺️ 生源地理分析",
    "🏫 餵校分析",
    "🎯 入學管道分析",
    "🎓 學生背景分析",
    "📈 跨年度趨勢",
    "📋 原始資料",
])

# ──────────────────────────────────────────────
# Tab 1: 生源地理分析
# ──────────────────────────────────────────────
with tab1:
    st.header("🗺️ 分析一：生源地理分布分析")
    st.markdown("> 掌握學生從哪裡來，找出招生強勢區與待開發區")

    geo = GeoAnalyzer(filtered_df)

    # 洞察
    display_insights(geo.get_geo_insights())

    st.markdown("---")

    # 圖表
    col1, col2 = st.columns(2)
    with col1:
        year_options = [None] + sorted(
            filtered_df["入學學年"].unique().tolist()
        )
        selected_year_geo = st.selectbox(
            "選擇學年度", year_options,
            format_func=lambda x: "全部學年度" if x is None else f"{x} 學年度",
            key="geo_year",
        )
        st.plotly_chart(
            geo.plot_region_pie(selected_year_geo), use_container_width=True
        )

    with col2:
        st.plotly_chart(geo.plot_region_trend(), use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        top_n_county = st.slider("顯示前 N 大縣市", 5, 25, 15, key="geo_topn")
        st.plotly_chart(
            geo.plot_county_bar(selected_year_geo, top_n_county),
            use_container_width=True,
        )

    with col4:
        st.plotly_chart(
            geo.plot_county_trend_line(), use_container_width=True
        )

    # 明細表
    with st.expander("📋 查看完整縣市數據"):
        county_data = geo.get_county_distribution()
        st.dataframe(county_data, use_container_width=True)
        create_download_button(
            county_data, "生源地理分析.xlsx", "下載地理分析數據"
        )


# ──────────────────────────────────────────────
# Tab 2: 餵校分析
# ──────────────────────────────────────────────
with tab2:
    st.header("🏫 分析二：餵校（Feeder School）分析")
    st.markdown("> 辨識穩定的學生來源校，優化招生資源配置")

    feeder = FeederSchoolAnalyzer(filtered_df)

    # 洞察
    display_insights(feeder.get_feeder_insights())

    st.markdown("---")

    top_n_school = st.slider("顯示前 N 大來源學校", 10, 30, 20, key="feeder_topn")

    col1, col2 = st.columns([3, 2])
    with col1:
        st.plotly_chart(
            feeder.plot_school_ranking(top_n_school),
            use_container_width=True,
        )

    with col2:
        st.plotly_chart(
            feeder.plot_school_concentration(),
            use_container_width=True,
        )

    st.plotly_chart(
        feeder.plot_school_heatmap(min(top_n_school, 15)),
        use_container_width=True,
    )

    # 學校分類
    st.subheader("🏷️ 學校分類")
    classifications = feeder.classify_schools()

    class_tabs = st.tabs(list(classifications.keys()))
    for tab, (category, school_df) in zip(
        class_tabs, classifications.items()
    ):
        with tab:
            if not school_df.empty:
                st.dataframe(school_df, use_container_width=True)
                st.caption(f"共 {len(school_df)} 所學校")
            else:
                st.info("此類別無資料")

    # 逐年明細
    with st.expander("📋 查看學校逐年送生明細"):
        yearly_detail = feeder.get_school_yearly_detail(top_n_school)
        st.dataframe(yearly_detail, use_container_width=True)
        create_download_button(
            yearly_detail, "餵校分析.xlsx", "下載餵校分析數據"
        )


# ──────────────────────────────────────────────
# Tab 3: 入學管道分析
# ──────────────────────────────────────────────
with tab3:
    st.header("🎯 分析三：入學管道成效分析")
    st.markdown("> 評估各招生管道效益，優化資源配置")

    channel = ChannelAnalyzer(filtered_df)

    # 洞察
    display_insights(channel.get_channel_insights())

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        year_options_ch = [None] + sorted(
            filtered_df["入學學年"].unique().tolist()
        )
        selected_year_ch = st.selectbox(
            "選擇學年度", year_options_ch,
            format_func=lambda x: "全部學年度" if x is None else f"{x} 學年度",
            key="channel_year",
        )
        st.plotly_chart(
            channel.plot_channel_pie(selected_year_ch),
            use_container_width=True,
        )

    with col2:
        st.plotly_chart(
            channel.plot_channel_trend(), use_container_width=True
        )

    st.plotly_chart(
        channel.plot_channel_region_heatmap(), use_container_width=True
    )

    # 交叉分析表
    col3, col4 = st.columns(2)
    with col3:
        st.subheader("管道 × 區域")
        st.dataframe(
            channel.get_channel_region_cross(), use_container_width=True
        )
    with col4:
        st.subheader("管道 × 學歷")
        st.dataframe(
            channel.get_channel_school_type_cross(),
            use_container_width=True,
        )


# ──────────────────────────────────────────────
# Tab 4: 學生背景分析
# ──────────────────────────────────────────────
with tab4:
    st.header("🎓 分析四：學生背景輪廓分析")
    st.markdown("> 了解什麼背景的學生會選擇本系")

    profile = ProfileAnalyzer(filtered_df)

    # 洞察
    display_insights(profile.get_profile_insights())

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        year_options_pf = [None] + sorted(
            filtered_df["入學學年"].unique().tolist()
        )
        selected_year_pf = st.selectbox(
            "選擇學年度", year_options_pf,
            format_func=lambda x: "全部學年度" if x is None else f"{x} 學年度",
            key="profile_year",
        )
        st.plotly_chart(
            profile.plot_education_pie(selected_year_pf),
            use_container_width=True,
        )

    with col2:
        st.plotly_chart(
            profile.plot_education_trend(), use_container_width=True
        )

    col3, col4 = st.columns(2)
    with col3:
        st.plotly_chart(
            profile.plot_major_bar(10), use_container_width=True
        )

    with col4:
        st.plotly_chart(
            profile.plot_major_treemap(selected_year_pf),
            use_container_width=True,
        )

    with st.expander("📋 學歷 × 科系 交叉分析表"):
        st.dataframe(
            profile.get_edu_major_cross(), use_container_width=True
        )


# ──────────────────────────────────────────────
# Tab 5: 跨年度趨勢
# ──────────────────────────────────────────────
with tab5:
    st.header("📈 分析五：跨年度綜合趨勢分析")
    st.markdown("> 追蹤長期趨勢，預測未來招生走向")

    trend = TrendAnalyzer(filtered_df)

    # KPI
    kpi = trend.plot_dashboard_kpi()
    kpi_cols = st.columns(4)
    kpi_cols[0].metric(
        f"📅 {kpi['最新學年']} 學年度學生數",
        kpi["最新學生數"],
        delta=f"{kpi['學生數變化']:+d} ({kpi['成長率']:+.1f}%)",
    )
    kpi_cols[1].metric(
        "🏫 來源學校數",
        kpi["來源學校數"],
        delta=f"{kpi['學校數變化']:+d}",
    )
    kpi_cols[2].metric(
        "🏙️ 來源縣市數",
        kpi["來源縣市數"],
    )
    kpi_cols[3].metric(
        "📈 成長率",
        f"{kpi['成長率']:+.1f}%",
    )

    # 洞察
    display_insights(trend.get_trend_insights())

    st.markdown("---")

    st.plotly_chart(
        trend.plot_enrollment_trend(), use_container_width=True
    )

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(
            trend.plot_department_trend(), use_container_width=True
        )
    with col2:
        st.plotly_chart(
            trend.plot_diversity_trend(), use_container_width=True
        )

    # 綜合摘要表
    st.subheader("📋 綜合摘要表")
    comprehensive = trend.get_comprehensive_summary()
    st.dataframe(comprehensive, use_container_width=True)
    create_download_button(
        comprehensive, "綜合趨勢分析.xlsx", "下載綜合趨勢數據"
    )


# ──────────────────────────────────────────────
# Tab 6: 原始資料
# ──────────────────────────────────────────────
with tab6:
    st.header("📋 原始資料檢視")

    st.dataframe(filtered_df, use_container_width=True)

    st.markdown(f"""
    **資料摘要：**
    - 總筆數：{len(filtered_df)}
    - 欄位數：{len(filtered_df.columns)}
    - 學年範圍：{filtered_df['入學學年'].min()} ~ {filtered_df['入學學年'].max()}
    """)

    create_download_button(
        filtered_df, "篩選後原始資料.xlsx", "下載篩選後資料"
    )
