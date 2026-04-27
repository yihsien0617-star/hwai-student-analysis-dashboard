"""
🎓 招生資料分析儀表板
版本: 1.1.0 (修正相容性)
Python: 3.8+
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO

# ════════════════════════════════════════════════════
# 全域設定
# ════════════════════════════════════════════════════

REQUIRED_COLUMNS = [
    "入學學年", "班級名稱", "學號", "縣市",
    "入學前學歷", "畢業學校", "畢業科系", "入學方式"
]

COUNTY_TO_REGION = {
    "台北市": "北部", "臺北市": "北部", "新北市": "北部",
    "基隆市": "北部", "桃園市": "北部", "新竹市": "北部",
    "新竹縣": "北部", "宜蘭縣": "北部",
    "台中市": "中部", "臺中市": "中部", "苗栗縣": "中部",
    "彰化縣": "中部", "南投縣": "中部", "雲林縣": "中部",
    "嘉義市": "南部", "嘉義縣": "南部",
    "台南市": "南部", "臺南市": "南部",
    "高雄市": "南部", "屏東縣": "南部",
    "花蓮縣": "東部", "台東縣": "東部", "臺東縣": "東部",
    "澎湖縣": "離島", "金門縣": "離島", "連江縣": "離島",
}

COLOR_REGION = {
    "北部": "#636EFA", "中部": "#EF553B", "南部": "#00CC96",
    "東部": "#AB63FA", "離島": "#FFA15A", "其他": "#999999",
}

CHART_TEMPLATE = "plotly_white"
CHART_FONT = dict(family="Microsoft JhengHei, Arial, sans-serif", size=13)


# ════════════════════════════════════════════════════
# 資料處理函數
# ════════════════════════════════════════════════════

def load_and_validate(uploaded_file):
    """載入 Excel 並驗證"""
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names

        if len(sheets) == 1:
            df = pd.read_excel(uploaded_file, sheet_name=0)
        else:
            frames = []
            for s in sheets:
                tmp = pd.read_excel(uploaded_file, sheet_name=s)
                tmp["來源工作表"] = s
                frames.append(tmp)
            df = pd.concat(frames, ignore_index=True)

        return df, None
    except Exception as e:
        return pd.DataFrame(), str(e)


def check_columns(df):
    """檢查必要欄位"""
    cols = set(df.columns.tolist())
    required = set(REQUIRED_COLUMNS)
    missing = required - cols
    found = required & cols
    return {
        "valid": len(missing) == 0,
        "missing": sorted(list(missing)),
        "found": sorted(list(found)),
        "total_rows": len(df),
    }


def clean_data(df):
    """資料清洗"""
    out = df.copy()

    # 去空白
    for c in out.select_dtypes(include=["object"]).columns:
        out[c] = out[c].astype(str).str.strip()

    # 統一台/臺
    if "縣市" in out.columns:
        out["縣市"] = out["縣市"].replace({
            "臺北市": "台北市", "臺中市": "台中市",
            "臺南市": "台南市", "臺東縣": "台東縣",
        })

    # 自動產生區域欄位
    if "區域" not in out.columns and "縣市" in out.columns:
        out["區域"] = out["縣市"].map(COUNTY_TO_REGION).fillna("其他")
    elif "區域" not in out.columns:
        out["區域"] = "未知"

    # 入學學年轉數值
    if "入學學年" in out.columns:
        out["入學學年"] = pd.to_numeric(out["入學學年"], errors="coerce")
        out = out.dropna(subset=["入學學年"])
        out["入學學年"] = out["入學學年"].astype(int)

    # 去重複
    before = len(out)
    out = out.drop_duplicates()
    removed = before - len(out)
    if removed > 0:
        st.info(f"ℹ️ 已移除 {removed} 筆重複資料")

    return out


def filter_data(df, departments=None, years=None):
    """篩選"""
    out = df.copy()
    if departments:
        out = out[out["班級名稱"].isin(departments)]
    if years:
        out = out[(out["入學學年"] >= years[0]) & (out["入學學年"] <= years[1])]
    return out


# ════════════════════════════════════════════════════
# 分析一：生源地理分析
# ════════════════════════════════════════════════════

def geo_region_pie(df, year=None):
    data = df if year is None else df[df["入學學年"] == year]
    label = f"{year} 學年" if year else "全部學年"
    counts = data["區域"].value_counts().reset_index()
    counts.columns = ["區域", "人數"]
    fig = px.pie(counts, values="人數", names="區域",
                 title=f"📍 {label} 生源區域分布",
                 color="區域", color_discrete_map=COLOR_REGION, hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label+value")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def geo_region_trend(df):
    data = df.groupby(["入學學年", "區域"]).size().reset_index(name="人數")
    fig = px.bar(data, x="入學學年", y="人數", color="區域",
                 title="📈 各區域生源逐年趨勢",
                 color_discrete_map=COLOR_REGION, text="人數", barmode="stack")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def geo_county_bar(df, year=None, top_n=15):
    data = df if year is None else df[df["入學學年"] == year]
    label = f"{year} 學年" if year else "全部學年"
    counts = data["縣市"].value_counts().head(top_n).reset_index()
    counts.columns = ["縣市", "人數"]
    counts = counts.sort_values("人數", ascending=True)
    fig = px.bar(counts, x="人數", y="縣市", orientation="h",
                 title=f"🏙️ {label} Top {top_n} 生源縣市",
                 text="人數", color="人數", color_continuous_scale="Tealgrn")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT, showlegend=False)
    return fig


def geo_county_trend(df, top_n=8):
    tops = df["縣市"].value_counts().head(top_n).index.tolist()
    data = df[df["縣市"].isin(tops)]
    trend = data.groupby(["入學學年", "縣市"]).size().reset_index(name="人數")
    fig = px.line(trend, x="入學學年", y="人數", color="縣市",
                  title=f"📊 Top {top_n} 縣市逐年趨勢", markers=True)
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def geo_insights(df):
    ins = []
    top_region = df["區域"].value_counts().index[0]
    top_pct = df["區域"].value_counts(normalize=True).iloc[0] * 100
    ins.append(f"🔵 最大生源區域：**{top_region}**，佔 **{top_pct:.1f}%**")

    top_city = df["縣市"].value_counts().index[0]
    top_city_n = df["縣市"].value_counts().iloc[0]
    ins.append(f"🏆 最大生源縣市：**{top_city}**，共 **{top_city_n}** 人")

    top3 = df["縣市"].value_counts(normalize=True).head(3).sum() * 100
    if top3 > 60:
        ins.append(f"⚠️ 前 3 大縣市佔 **{top3:.1f}%**，生源高度集中，建議分散")
    return ins


# ════════════════════════════════════════════════════
# 分析二：餵校分析
# ════════════════════════════════════════════════════

def feeder_ranking(df, top_n=20):
    rank = (df.groupby("畢業學校")
            .agg(總人數=("學號", "count"),
                 涵蓋學年數=("入學學年", "nunique"),
                 來源區域=("區域", "first"),
                 來源縣市=("縣市", "first"))
            .reset_index()
            .sort_values("總人數", ascending=False)
            .head(top_n))
    rank["平均每年送生"] = (rank["總人數"] / rank["涵蓋學年數"]).round(1)
    rank["排名"] = range(1, len(rank) + 1)
    return rank


def feeder_ranking_chart(df, top_n=20):
    r = feeder_ranking(df, top_n).sort_values("總人數", ascending=True)
    fig = px.bar(r, x="總人數", y="畢業學校", orientation="h",
                 title=f"🏫 Top {top_n} 來源學校排名", text="總人數",
                 color="來源區域",
                 hover_data=["來源縣市", "涵蓋學年數", "平均每年送生"])
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(500, top_n * 28))
    return fig


def feeder_heatmap(df, top_n=15):
    tops = df["畢業學校"].value_counts().head(top_n).index.tolist()
    sub = df[df["畢業學校"].isin(tops)]
    pivot = (sub.groupby(["畢業學校", "入學學年"]).size()
             .reset_index(name="人數")
             .pivot_table(index="畢業學校", columns="入學學年",
                          values="人數", fill_value=0))
    pivot["合計"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("合計", ascending=False)
    display = pivot.drop(columns=["合計"], errors="ignore")

    fig = go.Figure(data=go.Heatmap(
        z=display.values,
        x=[str(c) for c in display.columns],
        y=display.index.tolist(),
        colorscale="YlOrRd",
        text=display.values,
        texttemplate="%{text}",
        textfont=dict(size=12),
    ))
    fig.update_layout(title="🔥 來源學校 × 學年 熱力圖",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(400, top_n * 32))
    return fig


def feeder_concentration(df):
    counts = df["畢業學校"].value_counts()
    cum_pct = counts.cumsum() / counts.sum() * 100

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=list(range(1, len(cum_pct) + 1)), y=cum_pct.values,
        mode="lines+markers", name="累積佔比", marker=dict(size=4)))
    fig.add_hline(y=80, line_dash="dash", line_color="red",
                  annotation_text="80%")
    fig.update_layout(title="📊 來源學校集中度（累積佔比）",
                      xaxis_title="學校數", yaxis_title="累積佔比(%)",
                      template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def feeder_classify(df):
    years = sorted(df["入學學年"].unique())
    total_years = len(years)
    school_yr = df.groupby(["畢業學校", "入學學年"]).size().reset_index(name="人數")

    cats = {"穩定餵校": [], "成長中": [], "流失中": [], "新進學校": [], "偶發學校": []}

    for school in df["畢業學校"].unique():
        sd = school_yr[school_yr["畢業學校"] == school].sort_values("入學學年")
        yp = len(sd)
        total = sd["人數"].sum()

        if total_years <= 1:
            if total >= 3:
                cats["穩定餵校"].append({"學校": school, "總人數": total})
            else:
                cats["偶發學校"].append({"學校": school, "總人數": total})
            continue

        if yp == total_years and total / total_years >= 2:
            cats["穩定餵校"].append({"學校": school, "總人數": total, "年數": yp})
        elif yp >= 2:
            r, p = sd.iloc[-1]["人數"], sd.iloc[-2]["人數"]
            if r > p:
                cats["成長中"].append({"學校": school, "總人數": total,
                                     "趨勢": f"{p}→{r}"})
            elif r < p:
                cats["流失中"].append({"學校": school, "總人數": total,
                                     "趨勢": f"{p}→{r}"})
            else:
                cats["穩定餵校"].append({"學校": school, "總人數": total})
        elif yp == 1 and sd.iloc[0]["入學學年"] == years[-1]:
            cats["新進學校"].append({"學校": school, "總人數": total})
        else:
            cats["偶發學校"].append({"學校": school, "總人數": total})

    result = {}
    for k, v in cats.items():
        if v:
            result[k] = pd.DataFrame(v).sort_values("總人數", ascending=False)
        else:
            result[k] = pd.DataFrame()
    return result


def feeder_insights(df):
    ins = []
    total_schools = df["畢業學校"].nunique()
    total_students = len(df)
    ins.append(f"🏫 共 **{total_schools}** 所來源學校，"
               f"平均每校 **{total_students / max(total_schools, 1):.1f}** 人")

    top10 = df["畢業學校"].value_counts().head(10).sum()
    ins.append(f"📌 前 10 大學校佔 **{top10 / total_students * 100:.1f}%**")

    single = (df["畢業學校"].value_counts() == 1).sum()
    single_pct = single / max(total_schools, 1) * 100
    if single_pct > 50:
        ins.append(f"⚠️ 僅送 1 人的學校佔 **{single_pct:.1f}%**，生源極度分散")
    return ins


# ════════════════════════════════════════════════════
# 分析三：入學管道分析
# ════════════════════════════════════════════════════

def channel_pie(df, year=None):
    data = df if year is None else df[df["入學學年"] == year]
    label = f"{year} 學年" if year else "全部學年"
    counts = data["入學方式"].value_counts().reset_index()
    counts.columns = ["入學方式", "人數"]
    fig = px.pie(counts, values="人數", names="入學方式",
                 title=f"🎯 {label} 入學管道分布", hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label+value")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def channel_trend(df):
    data = df.groupby(["入學學年", "入學方式"]).size().reset_index(name="人數")
    fig = px.bar(data, x="入學學年", y="人數", color="入學方式",
                 title="📈 各入學管道逐年趨勢", text="人數", barmode="stack")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def channel_region_heatmap(df):
    cross = pd.crosstab(df["入學方式"], df["區域"])
    fig = go.Figure(data=go.Heatmap(
        z=cross.values, x=cross.columns.tolist(), y=cross.index.tolist(),
        colorscale="Blues", text=cross.values,
        texttemplate="%{text}", textfont=dict(size=14)))
    fig.update_layout(title="🗺️ 入學管道 × 區域",
                      template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def channel_insights(df):
    ins = []
    top_ch = df["入學方式"].value_counts()
    pct = top_ch.iloc[0] / len(df) * 100
    ins.append(f"🎯 主力管道：**{top_ch.index[0]}**，佔 **{pct:.1f}%**")

    if pct > 90:
        ins.append("🔴 **集中度風險：高** — 超過 90% 來自單一管道")
    elif pct > 70:
        ins.append("🟡 **集中度風險：中** — 建議拓展其他管道")
    else:
        ins.append("🟢 **集中度風險：低** — 管道多元")

    ins.append(f"📊 目前使用 **{df['入學方式'].nunique()}** 種管道")
    return ins


# ════════════════════════════════════════════════════
# 分析四：學生背景分析
# ════════════════════════════════════════════════════

def profile_edu_pie(df, year=None):
    data = df if year is None else df[df["入學學年"] == year]
    label = f"{year} 學年" if year else "全部學年"
    counts = data["入學前學歷"].value_counts().reset_index()
    counts.columns = ["入學前學歷", "人數"]
    fig = px.pie(counts, values="人數", names="入學前學歷",
                 title=f"🎓 {label} 入學前學歷分布", hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label+value")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def profile_edu_trend(df):
    data = df.groupby(["入學學年", "入學前學歷"]).size().reset_index(name="人數")
    fig = px.bar(data, x="入學學年", y="人數", color="入學前學歷",
                 title="📈 入學前學歷逐年趨勢", text="人數", barmode="stack")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def profile_major_bar(df, top_n=10):
    counts = df["畢業科系"].value_counts().head(top_n).reset_index()
    counts.columns = ["畢業科系", "人數"]
    counts = counts.sort_values("人數", ascending=True)
    fig = px.bar(counts, x="人數", y="畢業科系", orientation="h",
                 title=f"📋 Top {top_n} 畢業科系", text="人數",
                 color="人數", color_continuous_scale="Purples")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT, showlegend=False)
    return fig


def profile_treemap(df, year=None):
    data = df if year is None else df[df["入學學年"] == year]
    label = f"{year} 學年" if year else "全部學年"
    counts = data.groupby(["入學前學歷", "畢業科系"]).size().reset_index(name="人數")
    fig = px.treemap(counts, path=["入學前學歷", "畢業科系"], values="人數",
                     title=f"🌳 {label} 學歷→科系 樹狀圖",
                     color="人數", color_continuous_scale="Viridis")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def profile_insights(df):
    ins = []
    top_edu = df["入學前學歷"].value_counts()
    ins.append(f"🎓 主要學歷：**{top_edu.index[0]}**，"
               f"佔 **{top_edu.iloc[0] / len(df) * 100:.1f}%**")

    top_maj = df["畢業科系"].value_counts()
    ins.append(f"📚 主要科系：**{top_maj.index[0]}**，"
               f"佔 **{top_maj.iloc[0] / len(df) * 100:.1f}%**")
    return ins


# ════════════════════════════════════════════════════
# 分析五：跨年度趨勢
# ════════════════════════════════════════════════════

def trend_enrollment(df):
    t = df.groupby("入學學年").agg(
        學生數=("學號", "count"),
        來源學校數=("畢業學校", "nunique"),
        來源縣市數=("縣市", "nunique"),
    ).reset_index()
    if len(t) >= 2:
        t["成長率(%)"] = (t["學生數"].pct_change() * 100).round(1)
    return t


def trend_enrollment_chart(df):
    t = trend_enrollment(df)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=t["入學學年"], y=t["學生數"], name="學生數",
        marker_color="#636EFA", text=t["學生數"], textposition="outside",
    ), secondary_y=False)

    if "成長率(%)" in t.columns:
        fig.add_trace(go.Scatter(
            x=t["入學學年"], y=t["成長率(%)"], name="成長率(%)",
            mode="lines+markers+text",
            text=t["成長率(%)"].apply(
                lambda x: f"{x:+.1f}%" if pd.notna(x) else ""),
            textposition="top center",
            marker=dict(color="red", size=10),
            line=dict(color="red", width=2),
        ), secondary_y=True)

    fig.update_layout(title="📊 招生人數與成長率",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    fig.update_yaxes(title_text="學生人數", secondary_y=False)
    fig.update_yaxes(title_text="成長率(%)", secondary_y=True)
    return fig


def trend_dept_chart(df):
    data = df.groupby(["入學學年", "班級名稱"]).size().reset_index(name="人數")
    fig = px.line(data, x="入學學年", y="人數", color="班級名稱",
                  title="📈 各系所招生趨勢", markers=True)
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def trend_diversity(df):
    years = sorted(df["入學學年"].unique())
    rows = []
    for y in years:
        yd = df[df["入學學年"] == y]
        cs = yd["縣市"].value_counts(normalize=True)
        ss = yd["畢業學校"].value_counts(normalize=True)
        rows.append({
            "入學學年": y,
            "縣市HHI": round((cs ** 2).sum(), 4),
            "學校HHI": round((ss ** 2).sum(), 4),
            "來源學校數": yd["畢業學校"].nunique(),
        })
    return pd.DataFrame(rows)


def trend_diversity_chart(df):
    d = trend_diversity(df)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(
        x=d["入學學年"], y=d["縣市HHI"], name="縣市集中度(HHI)",
        mode="lines+markers", line=dict(color="#636EFA"),
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=d["入學學年"], y=d["來源學校數"], name="來源學校數",
        mode="lines+markers", line=dict(color="#00CC96"),
    ), secondary_y=True)
    fig.update_layout(title="🔄 生源多元性趨勢（HHI 越低越分散）",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    fig.update_yaxes(title_text="HHI", secondary_y=False)
    fig.update_yaxes(title_text="學校數", secondary_y=True)
    return fig


def trend_summary_table(df):
    years = sorted(df["入學學年"].unique())
    rows = []
    for y in years:
        yd = df[df["入學學年"] == y]
        top_region = yd["區域"].value_counts()
        top_ch = yd["入學方式"].value_counts()
        rows.append({
            "入學學年": y,
            "學生數": len(yd),
            "來源縣市": yd["縣市"].nunique(),
            "來源學校": yd["畢業學校"].nunique(),
            "主要區域": f"{top_region.index[0]} ({top_region.iloc[0] / len(yd) * 100:.0f}%)",
            "主要管道": f"{top_ch.index[0]} ({top_ch.iloc[0] / len(yd) * 100:.0f}%)",
        })
    return pd.DataFrame(rows)


def trend_kpi(df):
    t = trend_enrollment(df)
    latest = t.iloc[-1]
    kpi = {
        "最新學年": int(latest["入學學年"]),
        "學生數": int(latest["學生數"]),
        "學校數": int(latest["來源學校數"]),
        "縣市數": int(latest["來源縣市數"]),
        "學生變化": 0, "學校變化": 0, "成長率": 0.0,
    }
    if len(t) >= 2:
        prev = t.iloc[-2]
        kpi["學生變化"] = int(latest["學生數"] - prev["學生數"])
        kpi["學校變化"] = int(latest["來源學校數"] - prev["來源學校數"])
        kpi["成長率"] = round(
            (latest["學生數"] - prev["學生數"]) / max(prev["學生數"], 1) * 100, 1)
    return kpi


def trend_insights(df):
    ins = []
    t = trend_enrollment(df)
    if len(t) >= 2:
        g = t.iloc[-1].get("成長率(%)", 0)
        if pd.notna(g):
            if g > 5:
                ins.append(f"✅ 最新學年成長 **{g:.1f}%**")
            elif g < -5:
                ins.append(f"⚠️ 最新學年下降 **{abs(g):.1f}%**，需關注")
            else:
                ins.append(f"➡️ 最新學年變化 **{g:+.1f}%**，大致持平")

        total_chg = ((t.iloc[-1]["學生數"] - t.iloc[0]["學生數"])
                     / max(t.iloc[0]["學生數"], 1) * 100)
        ins.append(
            f"📈 {int(t.iloc[0]['入學學年'])}~{int(t.iloc[-1]['入學學年'])} "
            f"整體變化 **{total_chg:+.1f}%**")
    return ins


# ════════════════════════════════════════════════════
# 通用 UI 元件
# ════════════════════════════════════════════════════

def download_excel(df, filename, label):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df.to_excel(w, index=True, sheet_name="分析結果")
    buf.seek(0)
    st.download_button(
        label=f"📥 {label}", data=buf, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def show_insights(items, title="💡 關鍵發現"):
    st.subheader(title)
    for item in items:
        st.markdown(f"""<div style="background:#f0f2f6; padding:0.8rem 1rem;
            border-radius:8px; border-left:4px solid #667eea;
            margin-bottom:0.5rem;">{item}</div>""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════
# Streamlit 主程式
# ════════════════════════════════════════════════════

st.set_page_config(page_title="招生分析儀表板", page_icon="🎓",
                   layout="wide", initial_sidebar_state="expanded")

# ── 側邊欄 ──
with st.sidebar:
    st.title("🎓 招生分析儀表板")
    st.markdown("---")
    st.subheader("📂 上傳資料")
    uploaded = st.file_uploader("上傳新生資料 Excel", type=["xlsx", "xls"],
                                help="支援 .xlsx / .xls，可含多工作表")
    st.markdown("---")
    st.subheader("📋 必要欄位")
    st.markdown("""
    - `入學學年` (111, 112, 113…)
    - `班級名稱` (系所)
    - `學號`
    - `縣市`
    - `入學前學歷`
    - `畢業學校`
    - `畢業科系`
    - `入學方式`

    > 💡 `區域` 欄位可自動產生
    """)

# ── 主頁標題 ──
st.markdown("<h1 style='text-align:center; color:#1E3A5F;'>"
            "🎓 招生資料分析儀表板</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#666;'>"
            "上傳 Excel → 自動產生五大分析報告</p>", unsafe_allow_html=True)

if uploaded is None:
    st.info("👈 請從左側上傳 Excel 檔案以開始分析")
    with st.expander("📖 使用說明", expanded=True):
        st.markdown("""
        ### 五大分析面向
        | # | 面向 | 說明 |
        |---|------|------|
        | 1 | 🗺️ 生源地理 | 學生來自哪些區域/縣市 |
        | 2 | 🏫 餵校分析 | 穩定來源校辨識 |
        | 3 | 🎯 入學管道 | 各管道成效與風險 |
        | 4 | 🎓 學生背景 | 學歷/科系分布 |
        | 5 | 📈 跨年趨勢 | KPI 與多元性追蹤 |

        ### 操作步驟
        1. 準備含必要欄位的 Excel
        2. 左側上傳 → 自動分析
        3. 選擇系所/學年 → 切換頁籤
        4. 下載分析報表
        """)
    st.stop()

# ════════════════════════════════════════════════════
# 資料載入
# ════════════════════════════════════════════════════

with st.spinner("載入中..."):
    raw_df, err = load_and_validate(uploaded)

if err:
    st.error(f"❌ 檔案讀取失敗：{err}")
    st.stop()

if raw_df.empty:
    st.error("檔案無資料")
    st.stop()

# 欄位驗證
chk = check_columns(raw_df)
if not chk["valid"]:
    st.error("❌ 缺少必要欄位！")
    st.warning(f"缺少：`{'`, `'.join(chk['missing'])}`")
    st.info(f"已找到：`{'`, `'.join(chk['found'])}`")
    with st.expander("預覽前 5 筆原始資料"):
        st.dataframe(raw_df.head())
    st.stop()

# 清洗
df = clean_data(raw_df)

# ── 篩選器 ──
st.markdown("---")
c1, c2 = st.columns(2)
with c1:
    depts = sorted(df["班級名稱"].unique().tolist())
    sel_depts = st.multiselect("🏛️ 選擇系所（空白=全部）", depts)
with c2:
    ymin, ymax = int(df["入學學年"].min()), int(df["入學學年"].max())
    if ymin == ymax:
        sel_years = (ymin, ymax)
        st.info(f"📅 資料僅含 {ymin} 學年度")
    else:
        sel_years = st.slider("📅 學年範圍", ymin, ymax, (ymin, ymax))

fdf = filter_data(df, sel_depts if sel_depts else None, sel_years)

if fdf.empty:
    st.warning("篩選後無資料，請調整條件")
    st.stop()

# ── KPI 卡片 ──
st.markdown("---")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("👥 學生數", len(fdf))
k2.metric("📅 學年數", fdf["入學學年"].nunique())
k3.metric("🏛️ 系所數", fdf["班級名稱"].nunique())
k4.metric("🏙️ 來源縣市", fdf["縣市"].nunique())
k5.metric("🏫 來源學校", fdf["畢業學校"].nunique())

# ════════════════════════════════════════════════════
# 六大頁籤
# ════════════════════════════════════════════════════

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🗺️ 生源地理", "🏫 餵校分析", "🎯 入學管道",
    "🎓 學生背景", "📈 跨年趨勢", "📋 原始資料",
])

# ── Tab 1 ──
with tab1:
    st.header("🗺️ 生源地理分布分析")
    show_insights(geo_insights(fdf))
    st.markdown("---")

    gc1, gc2 = st.columns(2)
    with gc1:
        yr_opts = [None] + sorted(fdf["入學學年"].unique().tolist())
        sel_yr = st.selectbox("學年度", yr_opts, key="geo_yr",
                              format_func=lambda x: "全部" if x is None else str(x))
        st.plotly_chart(geo_region_pie(fdf, sel_yr), use_container_width=True)
    with gc2:
        st.plotly_chart(geo_region_trend(fdf), use_container_width=True)

    gc3, gc4 = st.columns(2)
    with gc3:
        tn = st.slider("Top N 縣市", 5, 25, 15, key="geo_tn")
        st.plotly_chart(geo_county_bar(fdf, sel_yr, tn), use_container_width=True)
    with gc4:
        st.plotly_chart(geo_county_trend(fdf), use_container_width=True)

# ── Tab 2 ──
with tab2:
    st.header("🏫 餵校（Feeder School）分析")
    show_insights(feeder_insights(fdf))
    st.markdown("---")

    tn2 = st.slider("Top N 學校", 10, 30, 20, key="fd_tn")

    fc1, fc2 = st.columns([3, 2])
    with fc1:
        st.plotly_chart(feeder_ranking_chart(fdf, tn2), use_container_width=True)
    with fc2:
        st.plotly_chart(feeder_concentration(fdf), use_container_width=True)

    st.plotly_chart(feeder_heatmap(fdf, min(tn2, 15)), use_container_width=True)

    st.subheader("🏷️ 學校分類")
    cats = feeder_classify(fdf)
    cat_tabs = st.tabs(list(cats.keys()))
    for ctab, (cat_name, cat_df) in zip(cat_tabs, cats.items()):
        with ctab:
            if not cat_df.empty:
                st.dataframe(cat_df, use_container_width=True)
            else:
                st.info("此類別無資料")

# ── Tab 3 ──
with tab3:
    st.header("🎯 入學管道成效分析")
    show_insights(channel_insights(fdf))
    st.markdown("---")

    cc1, cc2 = st.columns(2)
    with cc1:
        yr_opts2 = [None] + sorted(fdf["入學學年"].unique().tolist())
        sel_yr2 = st.selectbox("學年度", yr_opts2, key="ch_yr",
                               format_func=lambda x: "全部" if x is None else str(x))
        st.plotly_chart(channel_pie(fdf, sel_yr2), use_container_width=True)
    with cc2:
        st.plotly_chart(channel_trend(fdf), use_container_width=True)

    st.plotly_chart(channel_region_heatmap(fdf), use_container_width=True)

    cc3, cc4 = st.columns(2)
    with cc3:
        st.subheader("管道 × 區域")
        st.dataframe(pd.crosstab(fdf["入學方式"], fdf["區域"],
                                 margins=True, margins_name="合計"),
                     use_container_width=True)
    with cc4:
        st.subheader("管道 × 學歷")
        st.dataframe(pd.crosstab(fdf["入學方式"], fdf["入學前學歷"],
                                 margins=True, margins_name="合計"),
                     use_container_width=True)

# ── Tab 4 ──
with tab4:
    st.header("🎓 學生背景輪廓分析")
    show_insights(profile_insights(fdf))
    st.markdown("---")

    pc1, pc2 = st.columns(2)
    with pc1:
        yr_opts3 = [None] + sorted(fdf["入學學年"].unique().tolist())
        sel_yr3 = st.selectbox("學年度", yr_opts3, key="pf_yr",
                               format_func=lambda x: "全部" if x is None else str(x))
        st.plotly_chart(profile_edu_pie(fdf, sel_yr3), use_container_width=True)
    with pc2:
        st.plotly_chart(profile_edu_trend(fdf), use_container_width=True)

    pc3, pc4 = st.columns(2)
    with pc3:
        st.plotly_chart(profile_major_bar(fdf, 10), use_container_width=True)
    with pc4:
        st.plotly_chart(profile_treemap(fdf, sel_yr3), use_container_width=True)

# ── Tab 5 ──
with tab5:
    st.header("📈 跨年度綜合趨勢分析")

    kpi = trend_kpi(fdf)
    tk1, tk2, tk3, tk4 = st.columns(4)
    tk1.metric(f"{kpi['最新學年']} 學年學生數", kpi["學生數"],
               delta=f"{kpi['學生變化']:+d} ({kpi['成長率']:+.1f}%)")
    tk2.metric("來源學校數", kpi["學校數"], delta=f"{kpi['學校變化']:+d}")
    tk3.metric("來源縣市數", kpi["縣市數"])
    tk4.metric("成長率", f"{kpi['成長率']:+.1f}%")

    show_insights(trend_insights(fdf))
    st.markdown("---")

    st.plotly_chart(trend_enrollment_chart(fdf), use_container_width=True)

    tc1, tc2 = st.columns(2)
    with tc1:
        st.plotly_chart(trend_dept_chart(fdf), use_container_width=True)
    with tc2:
        st.plotly_chart(trend_diversity_chart(fdf), use_container_width=True)

    st.subheader("📋 綜合摘要表")
    summary = trend_summary_table(fdf)
    st.dataframe(summary, use_container_width=True)
    download_excel(summary, "綜合趨勢.xlsx", "下載趨勢數據")

# ── Tab 6 ──
with tab6:
    st.header("📋 原始資料檢視")
    st.dataframe(fdf, use_container_width=True)
    st.caption(f"共 {len(fdf)} 筆，{len(fdf.columns)} 個欄位")
    download_excel(fdf, "篩選後資料.xlsx", "下載資料")
