"""
🎓 招生資料分析儀表板
版本: 1.3.1
- 餵校 → 生源學校
- 入學管道新增生源學校交叉分析
- 新增「系所分析」：以系為單位的完整剖析
- 多班合併：資管一甲+資管一乙 → 資管（自動辨識）
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import re

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
# 班級名稱 → 系所名稱 合併
# ════════════════════════════════════════════════════

def extract_department(class_name: str) -> str:
    """
    將班級名稱去除年級＋班別後綴，合併為系所。
    例：
      資管一甲 → 資管
      電機二乙 → 電機
      護理三   → 護理
      企管四丙 → 企管
      資工甲   → 資工
      日四技資管一甲 → 日四技資管
      夜四技電機二A → 夜四技電機
    策略：從尾部移除 [一二三四五1-5]?[甲乙丙丁戊A-Ea-e]? 的組合
    """
    s = str(class_name).strip()
    # 移除尾部的年級+班別，支援中文數字/阿拉伯數字 + 甲乙丙丁/ABC
    s = re.sub(r'[一二三四五六七1-7]?[甲乙丙丁戊A-Ea-e]?$', '', s)
    # 再移除尾部可能剩餘的獨立年級數字
    s = re.sub(r'[一二三四五六七1-7]$', '', s)
    return s.strip() if s.strip() else str(class_name).strip()


# ════════════════════════════════════════════════════
# 資料處理
# ════════════════════════════════════════════════════

def load_and_validate(uploaded_file):
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
    out = df.copy()
    for c in out.select_dtypes(include=["object"]).columns:
        out[c] = out[c].astype(str).str.strip()
    if "縣市" in out.columns:
        out["縣市"] = out["縣市"].replace({
            "臺北市": "台北市", "臺中市": "台中市",
            "臺南市": "台南市", "臺東縣": "台東縣",
        })
    if "區域" not in out.columns and "縣市" in out.columns:
        out["區域"] = out["縣市"].map(COUNTY_TO_REGION).fillna("其他")
    elif "區域" not in out.columns:
        out["區域"] = "未知"
    if "入學學年" in out.columns:
        out["入學學年"] = pd.to_numeric(out["入學學年"], errors="coerce")
        out = out.dropna(subset=["入學學年"])
        out["入學學年"] = out["入學學年"].astype(int)

    # === 新增：產生「系所」欄位 ===
    if "班級名稱" in out.columns:
        out["系所"] = out["班級名稱"].apply(extract_department)

    before = len(out)
    out = out.drop_duplicates()
    removed = before - len(out)
    if removed > 0:
        st.info(f"ℹ️ 已移除 {removed} 筆重複資料")
    return out


def filter_data(df, departments=None, years=None):
    out = df.copy()
    if departments:
        out = out[out["系所"].isin(departments)]
    if years:
        out = out[(out["入學學年"] >= years[0]) & (out["入學學年"] <= years[1])]
    return out


# ════════════════════════════════════════════════════
# 分析一：生源地理
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
# 分析二：生源學校分析
# ════════════════════════════════════════════════════

def source_school_ranking(df, top_n=20):
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


def source_school_ranking_chart(df, top_n=20):
    r = source_school_ranking(df, top_n).sort_values("總人數", ascending=True)
    fig = px.bar(r, x="總人數", y="畢業學校", orientation="h",
                 title=f"🏫 Top {top_n} 生源學校排名", text="總人數",
                 color="來源區域",
                 hover_data=["來源縣市", "涵蓋學年數", "平均每年送生"])
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(500, top_n * 28))
    return fig


def source_school_heatmap(df, top_n=15):
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
    fig.update_layout(title="🔥 生源學校 × 學年 熱力圖",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(400, top_n * 32))
    return fig


def source_school_concentration(df):
    counts = df["畢業學校"].value_counts()
    cum_pct = counts.cumsum() / counts.sum() * 100
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=list(range(1, len(cum_pct) + 1)), y=cum_pct.values,
        mode="lines+markers", name="累積佔比", marker=dict(size=4)))
    fig.add_hline(y=80, line_dash="dash", line_color="red",
                  annotation_text="80%")
    fig.update_layout(title="📊 生源學校集中度（累積佔比）",
                      xaxis_title="學校數", yaxis_title="累積佔比(%)",
                      template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def source_school_classify(df):
    years = sorted(df["入學學年"].unique())
    total_years = len(years)
    school_yr = df.groupby(["畢業學校", "入學學年"]).size().reset_index(name="人數")
    cats = {"穩定生源校": [], "成長中": [], "流失中": [], "新進學校": [], "偶發學校": []}

    for school in df["畢業學校"].unique():
        sd = school_yr[school_yr["畢業學校"] == school].sort_values("入學學年")
        yp = len(sd)
        total = sd["人數"].sum()

        if total_years <= 1:
            if total >= 3:
                cats["穩定生源校"].append({"學校": school, "總人數": total})
            else:
                cats["偶發學校"].append({"學校": school, "總人數": total})
            continue

        if yp == total_years and total / total_years >= 2:
            cats["穩定生源校"].append({"學校": school, "總人數": total, "年數": yp})
        elif yp >= 2:
            r, p = sd.iloc[-1]["人數"], sd.iloc[-2]["人數"]
            if r > p:
                cats["成長中"].append({"學校": school, "總人數": total,
                                     "趨勢": f"{p}→{r}"})
            elif r < p:
                cats["流失中"].append({"學校": school, "總人數": total,
                                     "趨勢": f"{p}→{r}"})
            else:
                cats["穩定生源校"].append({"學校": school, "總人數": total})
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


def source_school_insights(df):
    ins = []
    total_schools = df["畢業學校"].nunique()
    total_students = len(df)
    ins.append(f"🏫 共 **{total_schools}** 所生源學校，"
               f"平均每校 **{total_students / max(total_schools, 1):.1f}** 人")
    top10 = df["畢業學校"].value_counts().head(10).sum()
    ins.append(f"📌 前 10 大學校佔 **{top10 / total_students * 100:.1f}%**")
    single = (df["畢業學校"].value_counts() == 1).sum()
    single_pct = single / max(total_schools, 1) * 100
    if single_pct > 50:
        ins.append(f"⚠️ 僅送 1 人的學校佔 **{single_pct:.1f}%**，長尾效應明顯")
    return ins


# ════════════════════════════════════════════════════
# 分析三：入學管道
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


def channel_school_top(df, channel, top_n=10):
    sub = df[df["入學方式"] == channel]
    if sub.empty:
        return pd.DataFrame()
    rank = (sub.groupby("畢業學校")
            .agg(人數=("學號", "count"),
                 來源縣市=("縣市", "first"),
                 來源區域=("區域", "first"))
            .reset_index()
            .sort_values("人數", ascending=False)
            .head(top_n))
    rank["佔該管道比例"] = (rank["人數"] / len(sub) * 100).round(1)
    rank["排名"] = range(1, len(rank) + 1)
    return rank


def channel_school_top_chart(df, channel, top_n=10):
    r = channel_school_top(df, channel, top_n)
    if r.empty:
        return go.Figure().update_layout(title=f"{channel}：無資料")
    r = r.sort_values("人數", ascending=True)
    fig = px.bar(r, x="人數", y="畢業學校", orientation="h",
                 title=f"🏫【{channel}】Top {top_n} 生源學校",
                 text="人數", color="來源區域",
                 hover_data=["來源縣市", "佔該管道比例"])
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(400, top_n * 30))
    return fig


def channel_school_summary(df):
    rows = []
    for ch in df["入學方式"].unique():
        sub = df[df["入學方式"] == ch]
        n_student = len(sub)
        n_school = sub["畢業學校"].nunique()
        top1 = sub["畢業學校"].value_counts()
        top1_name = top1.index[0] if len(top1) > 0 else "—"
        top1_n = top1.iloc[0] if len(top1) > 0 else 0
        top1_pct = top1_n / max(n_student, 1) * 100
        top5_pct = top1.head(5).sum() / max(n_student, 1) * 100
        shares = top1 / n_student
        hhi = round((shares ** 2).sum(), 4)
        rows.append({
            "入學管道": ch, "學生數": n_student, "生源學校數": n_school,
            "平均每校人數": round(n_student / max(n_school, 1), 1),
            "第一大校": f"{top1_name} ({top1_n}人)",
            "第一大校佔比": f"{top1_pct:.1f}%",
            "前5大校佔比": f"{top5_pct:.1f}%",
            "HHI集中度": hhi,
        })
    return pd.DataFrame(rows).sort_values("學生數", ascending=False)


def channel_school_heatmap(df, top_n=10):
    tops = df["畢業學校"].value_counts().head(top_n).index.tolist()
    sub = df[df["畢業學校"].isin(tops)]
    cross = pd.crosstab(sub["入學方式"], sub["畢業學校"])
    col_order = df["畢業學校"].value_counts().head(top_n).index.tolist()
    cross = cross.reindex(columns=[c for c in col_order if c in cross.columns],
                          fill_value=0)
    fig = go.Figure(data=go.Heatmap(
        z=cross.values, x=cross.columns.tolist(), y=cross.index.tolist(),
        colorscale="Oranges", text=cross.values,
        texttemplate="%{text}", textfont=dict(size=12),
    ))
    fig.update_layout(title=f"🔥 入學管道 × Top {top_n} 生源學校",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(350, len(cross) * 50),
                      xaxis=dict(tickangle=45))
    return fig


def channel_school_trend(df, channel, top_n=5):
    sub = df[df["入學方式"] == channel]
    if sub.empty:
        return go.Figure().update_layout(title=f"{channel}：無資料")
    tops = sub["畢業學校"].value_counts().head(top_n).index.tolist()
    data = sub[sub["畢業學校"].isin(tops)]
    trend = data.groupby(["入學學年", "畢業學校"]).size().reset_index(name="人數")
    fig = px.line(trend, x="入學學年", y="人數", color="畢業學校",
                  title=f"📈【{channel}】Top {top_n} 生源學校逐年趨勢",
                  markers=True)
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
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
    summary = channel_school_summary(df)
    if not summary.empty:
        max_hhi_row = summary.loc[summary["HHI集中度"].idxmax()]
        ins.append(f"🏫 學校來源最集中的管道：**{max_hhi_row['入學管道']}**"
                   f"（HHI={max_hhi_row['HHI集中度']:.4f}，"
                   f"前5大校佔 {max_hhi_row['前5大校佔比']}）")
    return ins


# ════════════════════════════════════════════════════
# 分析四：學生背景
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
# 分析五：跨年趨勢
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
    data = df.groupby(["入學學年", "系所"]).size().reset_index(name="人數")
    fig = px.line(data, x="入學學年", y="人數", color="系所",
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
            "入學學年": y, "學生數": len(yd),
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
# 分析六：系所分析（使用合併後「系所」欄位）
# ════════════════════════════════════════════════════

def dept_overview_table(df):
    rows = []
    for dept in sorted(df["系所"].unique()):
        dd = df[df["系所"] == dept]
        n = len(dd)
        n_class = dd["班級名稱"].nunique()
        n_school = dd["畢業學校"].nunique()
        n_city = dd["縣市"].nunique()
        n_channel = dd["入學方式"].nunique()
        top_region = dd["區域"].value_counts()
        top_school = dd["畢業學校"].value_counts()
        top_channel = dd["入學方式"].value_counts()
        # 班級明細
        class_list = "、".join(sorted(dd["班級名稱"].unique()))
        # 成長率
        yr_counts = dd.groupby("入學學年").size()
        growth = ""
        if len(yr_counts) >= 2:
            last = yr_counts.iloc[-1]
            prev = yr_counts.iloc[-2]
            g = (last - prev) / max(prev, 1) * 100
            growth = f"{g:+.1f}%"
        # HHI
        shares = dd["畢業學校"].value_counts(normalize=True)
        hhi = round((shares ** 2).sum(), 4)
        rows.append({
            "系所": dept,
            "包含班級": class_list,
            "班級數": n_class,
            "學生數": n,
            "涵蓋學年": f"{dd['入學學年'].min()}~{dd['入學學年'].max()}",
            "來源學校數": n_school,
            "來源縣市數": n_city,
            "管道數": n_channel,
            "最大區域": f"{top_region.index[0]} ({top_region.iloc[0] / n * 100:.0f}%)",
            "最大生源校": f"{top_school.index[0]} ({top_school.iloc[0]}人)",
            "主要管道": f"{top_channel.index[0]} ({top_channel.iloc[0] / n * 100:.0f}%)",
            "學校HHI": hhi,
            "最近成長率": growth,
        })
    return pd.DataFrame(rows)


def dept_compare_bar(df, metric="學生數"):
    overview = dept_overview_table(df)
    if metric not in overview.columns:
        metric = "學生數"
    data = overview[["系所", metric]].sort_values(metric, ascending=True)
    fig = px.bar(data, x=metric, y="系所", orientation="h",
                 title=f"🏛️ 各系所 {metric} 比較", text=metric,
                 color=metric, color_continuous_scale="Viridis")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT, showlegend=False,
                      height=max(400, len(data) * 35))
    return fig


def dept_trend_lines(df):
    data = df.groupby(["入學學年", "系所"]).size().reset_index(name="人數")
    fig = px.line(data, x="入學學年", y="人數", color="系所",
                  title="📈 各系所逐年招生趨勢", markers=True)
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def dept_geo_pie(df, dept):
    dd = df[df["系所"] == dept]
    counts = dd["區域"].value_counts().reset_index()
    counts.columns = ["區域", "人數"]
    fig = px.pie(counts, values="人數", names="區域",
                 title=f"📍【{dept}】生源區域",
                 color="區域", color_discrete_map=COLOR_REGION, hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label+value")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def dept_county_bar(df, dept, top_n=10):
    dd = df[df["系所"] == dept]
    counts = dd["縣市"].value_counts().head(top_n).reset_index()
    counts.columns = ["縣市", "人數"]
    counts = counts.sort_values("人數", ascending=True)
    fig = px.bar(counts, x="人數", y="縣市", orientation="h",
                 title=f"🏙️【{dept}】Top {top_n} 生源縣市",
                 text="人數", color="人數", color_continuous_scale="Tealgrn")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT, showlegend=False)
    return fig


def dept_school_bar(df, dept, top_n=10):
    dd = df[df["系所"] == dept]
    counts = dd["畢業學校"].value_counts().head(top_n).reset_index()
    counts.columns = ["畢業學校", "人數"]
    counts = counts.sort_values("人數", ascending=True)
    fig = px.bar(counts, x="人數", y="畢業學校", orientation="h",
                 title=f"🏫【{dept}】Top {top_n} 生源學校",
                 text="人數", color="人數", color_continuous_scale="YlOrRd")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT, showlegend=False,
                      height=max(350, top_n * 30))
    return fig


def dept_school_heatmap(df, dept, top_n=10):
    dd = df[df["系所"] == dept]
    tops = dd["畢業學校"].value_counts().head(top_n).index.tolist()
    sub = dd[dd["畢業學校"].isin(tops)]
    if sub.empty:
        return go.Figure().update_layout(title=f"【{dept}】無足夠資料")
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
        colorscale="YlOrRd", text=display.values,
        texttemplate="%{text}", textfont=dict(size=12)))
    fig.update_layout(title=f"🔥【{dept}】生源學校 × 學年",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(300, top_n * 30))
    return fig


def dept_channel_pie(df, dept):
    dd = df[df["系所"] == dept]
    counts = dd["入學方式"].value_counts().reset_index()
    counts.columns = ["入學方式", "人數"]
    fig = px.pie(counts, values="人數", names="入學方式",
                 title=f"🎯【{dept}】入學管道分布", hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label+value")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def dept_channel_trend(df, dept):
    dd = df[df["系所"] == dept]
    data = dd.groupby(["入學學年", "入學方式"]).size().reset_index(name="人數")
    fig = px.bar(data, x="入學學年", y="人數", color="入學方式",
                 title=f"📈【{dept}】入學管道逐年趨勢",
                 text="人數", barmode="stack")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def dept_channel_school(df, dept, top_n=8):
    dd = df[df["系所"] == dept]
    tops = dd["畢業學校"].value_counts().head(top_n).index.tolist()
    sub = dd[dd["畢業學校"].isin(tops)]
    if sub.empty:
        return go.Figure().update_layout(title=f"【{dept}】無資料")
    cross = pd.crosstab(sub["入學方式"], sub["畢業學校"])
    col_order = dd["畢業學校"].value_counts().head(top_n).index.tolist()
    cross = cross.reindex(columns=[c for c in col_order if c in cross.columns],
                          fill_value=0)
    fig = go.Figure(data=go.Heatmap(
        z=cross.values, x=cross.columns.tolist(), y=cross.index.tolist(),
        colorscale="Oranges", text=cross.values,
        texttemplate="%{text}", textfont=dict(size=12)))
    fig.update_layout(title=f"🔥【{dept}】管道 × Top {top_n} 生源學校",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      height=max(300, len(cross) * 45),
                      xaxis=dict(tickangle=45))
    return fig


def dept_edu_pie(df, dept):
    dd = df[df["系所"] == dept]
    counts = dd["入學前學歷"].value_counts().reset_index()
    counts.columns = ["入學前學歷", "人數"]
    fig = px.pie(counts, values="人數", names="入學前學歷",
                 title=f"🎓【{dept}】入學前學歷", hole=0.35)
    fig.update_traces(textposition="inside", textinfo="percent+label+value")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT)
    return fig


def dept_major_bar(df, dept, top_n=10):
    dd = df[df["系所"] == dept]
    counts = dd["畢業科系"].value_counts().head(top_n).reset_index()
    counts.columns = ["畢業科系", "人數"]
    counts = counts.sort_values("人數", ascending=True)
    fig = px.bar(counts, x="人數", y="畢業科系", orientation="h",
                 title=f"📋【{dept}】Top {top_n} 畢業科系",
                 text="人數", color="人數", color_continuous_scale="Purples")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT, showlegend=False)
    return fig


def dept_enrollment_chart(df, dept):
    dd = df[df["系所"] == dept]
    t = dd.groupby("入學學年").size().reset_index(name="學生數")
    if len(t) >= 2:
        t["成長率(%)"] = (t["學生數"].pct_change() * 100).round(1)
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
    fig.update_layout(title=f"📊【{dept}】招生人數與成長率",
                      template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    fig.update_yaxes(title_text="學生人數", secondary_y=False)
    fig.update_yaxes(title_text="成長率(%)", secondary_y=True)
    return fig


def dept_class_breakdown(df, dept):
    """顯示同一系所下各班級的人數明細"""
    dd = df[df["系所"] == dept]
    data = dd.groupby(["入學學年", "班級名稱"]).size().reset_index(name="人數")
    fig = px.bar(data, x="入學學年", y="人數", color="班級名稱",
                 title=f"📊【{dept}】各班級逐年人數",
                 text="人數", barmode="group")
    fig.update_layout(template=CHART_TEMPLATE, font=CHART_FONT,
                      xaxis=dict(dtick=1))
    return fig


def dept_insights(df, dept):
    dd = df[df["系所"] == dept]
    ins = []
    n = len(dd)
    total = len(df)
    n_class = dd["班級名稱"].nunique()
    class_list = "、".join(sorted(dd["班級名稱"].unique()))

    if n_class > 1:
        ins.append(f"📋 **{dept}** 包含 **{n_class}** 個班級：{class_list}")
    ins.append(f"📊 共 **{n}** 人，佔全校 **{n / max(total, 1) * 100:.1f}%**")

    # 區域
    top_r = dd["區域"].value_counts()
    ins.append(f"📍 主要區域：**{top_r.index[0]}** ({top_r.iloc[0] / n * 100:.0f}%)")

    # 學校
    n_school = dd["畢業學校"].nunique()
    top_s = dd["畢業學校"].value_counts()
    ins.append(f"🏫 來自 **{n_school}** 所學校，"
               f"第一大生源校：**{top_s.index[0]}** ({top_s.iloc[0]}人)")

    # 管道
    top_c = dd["入學方式"].value_counts()
    ins.append(f"🎯 主要管道：**{top_c.index[0]}** ({top_c.iloc[0] / n * 100:.0f}%)")

    # 成長
    yr_counts = dd.groupby("入學學年").size()
    if len(yr_counts) >= 2:
        g = (yr_counts.iloc[-1] - yr_counts.iloc[-2]) / max(yr_counts.iloc[-2], 1) * 100
        if g > 5:
            ins.append(f"✅ 最新學年成長 **{g:.1f}%**")
        elif g < -5:
            ins.append(f"⚠️ 最新學年下降 **{abs(g):.1f}%**，需關注")
        else:
            ins.append(f"➡️ 最新學年變化 **{g:+.1f}%**，大致持平")

    # 集中度
    top5_pct = top_s.head(5).sum() / n * 100
    if top5_pct > 50:
        ins.append(f"⚠️ 前 5 大生源校佔 **{top5_pct:.1f}%**，集中度偏高")

    return ins


# ════════════════════════════════════════════════════
# 通用 UI
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
    - `班級名稱` (如：資管一甲)
    - `學號`
    - `縣市`
    - `入學前學歷`
    - `畢業學校`
    - `畢業科系`
    - `入學方式`
    > 💡 系統自動將「資管一甲+資管一乙」合併為「資管」系
    """)

st.markdown("<h1 style='text-align:center; color:#1E3A5F;'>"
            "🎓 招生資料分析儀表板</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#666;'>"
            "上傳 Excel → 自動產生六大分析報告（多班自動合併為系所）</p>",
            unsafe_allow_html=True)

if uploaded is None:
    st.info("👈 請從左側上傳 Excel 檔案以開始分析")
    with st.expander("📖 使用說明", expanded=True):
        st.markdown("""
        ### 六大分析面向
        | # | 面向 | 說明 |
        |---|------|------|
        | 1 | 🗺️ 生源地理 | 學生來自哪些區域/縣市 |
        | 2 | 🏫 生源學校 | 穩定來源校辨識與分類 |
        | 3 | 🎯 入學管道 | 各管道成效、生源學校交叉分析 |
        | 4 | 🎓 學生背景 | 學歷/科系分布 |
        | 5 | 📈 跨年趨勢 | KPI 與多元性追蹤 |
        | 6 | 🏛️ 系所分析 | 以系為單位完整剖析（多班自動合併） |

        ### 🔄 多班合併邏輯
        班級名稱尾部的年級＋班別會自動去除：
        - `資管一甲` + `資管一乙` → **資管**
        - `電機二A` + `電機二B` → **電機**
        - `護理三` → **護理**
        """)
    st.stop()

# 載入
with st.spinner("載入中..."):
    raw_df, err = load_and_validate(uploaded)

if err:
    st.error(f"❌ 讀取失敗：{err}")
    st.stop()
if raw_df.empty:
    st.error("檔案無資料")
    st.stop()

chk = check_columns(raw_df)
if not chk["valid"]:
    st.error("❌ 缺少必要欄位！")
    st.warning(f"缺少：`{'`, `'.join(chk['missing'])}`")
    st.info(f"已找到：`{'`, `'.join(chk['found'])}`")
    with st.expander("預覽前 5 筆"):
        st.dataframe(raw_df.head())
    st.stop()

df = clean_data(raw_df)

# === 顯示合併對照 ===
with st.sidebar:
    st.markdown("---")
    st.subheader("🔄 班級→系所 對照")
    mapping = (df[["班級名稱", "系所"]].drop_duplicates()
               .sort_values("系所").reset_index(drop=True))
    st.dataframe(mapping, use_container_width=True, hide_index=True, height=200)

# 篩選器（改用「系所」）
st.markdown("---")
c1, c2 = st.columns(2)
with c1:
    dept_list = sorted(df["系所"].unique().tolist())
    sel_depts = st.multiselect("🏛️ 選擇系所（空白=全部）", dept_list)
with c2:
    ymin, ymax = int(df["入學學年"].min()), int(df["入學學年"].max())
    if ymin == ymax:
        sel_years = (ymin, ymax)
        st.info(f"📅 僅含 {ymin} 學年")
    else:
        sel_years = st.slider("📅 學年範圍", ymin, ymax, (ymin, ymax))

fdf = filter_data(df, sel_depts if sel_depts else None, sel_years)

if fdf.empty:
    st.warning("篩選後無資料")
    st.stop()

# KPI
st.markdown("---")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("👥 學生數", len(fdf))
k2.metric("📅 學年數", fdf["入學學年"].nunique())
k3.metric("🏛️ 系所數", fdf["系所"].nunique())
k4.metric("🏙️ 來源縣市", fdf["縣市"].nunique())
k5.metric("🏫 生源學校", fdf["畢業學校"].nunique())

# ════════════════════════════════════════════════════
# 七大頁籤
# ════════════════════════════════════════════════════

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🗺️ 生源地理", "🏫 生源學校", "🎯 入學管道",
    "🎓 學生背景", "📈 跨年趨勢", "🏛️ 系所分析", "📋 原始資料",
])

# ── Tab 1：生源地理 ──
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

# ── Tab 2：生源學校 ──
with tab2:
    st.header("🏫 生源學校分析")
    show_insights(source_school_insights(fdf))
    st.markdown("---")
    tn2 = st.slider("Top N 學校", 10, 30, 20, key="ss_tn")
    fc1, fc2 = st.columns([3, 2])
    with fc1:
        st.plotly_chart(source_school_ranking_chart(fdf, tn2),
                        use_container_width=True)
    with fc2:
        st.plotly_chart(source_school_concentration(fdf),
                        use_container_width=True)
    st.plotly_chart(source_school_heatmap(fdf, min(tn2, 15)),
                    use_container_width=True)
    st.subheader("🏷️ 學校分類")
    cats = source_school_classify(fdf)
    cat_tabs = st.tabs(list(cats.keys()))
    for ctab, (cat_name, cat_df) in zip(cat_tabs, cats.items()):
        with ctab:
            if not cat_df.empty:
                st.dataframe(cat_df, use_container_width=True)
            else:
                st.info("此類別無資料")

# ── Tab 3：入學管道 ──
with tab3:
    st.header("🎯 入學管道成效分析")
    show_insights(channel_insights(fdf))
    st.markdown("---")

    st.subheader("📊 管道分布與趨勢")
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
        st.markdown("**管道 × 區域**")
        st.dataframe(pd.crosstab(fdf["入學方式"], fdf["區域"],
                                 margins=True, margins_name="合計"),
                     use_container_width=True)
    with cc4:
        st.markdown("**管道 × 學歷**")
        st.dataframe(pd.crosstab(fdf["入學方式"], fdf["入學前學歷"],
                                 margins=True, margins_name="合計"),
                     use_container_width=True)

    st.markdown("---")
    st.subheader("🏫 各管道生源學校分析")

    ch_summary = channel_school_summary(fdf)
    st.markdown("**▎各管道生源學校概況**")
    st.dataframe(ch_summary, use_container_width=True, hide_index=True)
    download_excel(ch_summary, "管道生源學校摘要.xlsx", "下載管道學校摘要")

    ch_tn = st.slider("熱力圖 Top N 學校", 5, 20, 10, key="ch_school_tn")
    st.plotly_chart(channel_school_heatmap(fdf, ch_tn), use_container_width=True)

    st.markdown("---")
    st.markdown("**▎單一管道深入分析**")
    channels = sorted(fdf["入學方式"].unique().tolist())
    sel_channel = st.selectbox("選擇入學管道", channels, key="ch_sel")
    ch_top_n = st.slider("Top N 學校", 5, 20, 10, key="ch_detail_tn")

    chd1, chd2 = st.columns(2)
    with chd1:
        st.plotly_chart(channel_school_top_chart(fdf, sel_channel, ch_top_n),
                        use_container_width=True)
    with chd2:
        st.plotly_chart(channel_school_trend(fdf, sel_channel, min(ch_top_n, 5)),
                        use_container_width=True)

    detail_df = channel_school_top(fdf, sel_channel, ch_top_n)
    if not detail_df.empty:
        st.markdown(f"**▎【{sel_channel}】Top {ch_top_n} 生源學校明細**")
        st.dataframe(detail_df, use_container_width=True, hide_index=True)
        download_excel(detail_df, f"{sel_channel}_生源學校.xlsx",
                       f"下載 {sel_channel} 學校明細")

# ── Tab 4：學生背景 ──
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

# ── Tab 5：跨年趨勢 ──
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

# ── Tab 6：系所分析（合併後） ──
with tab6:
    st.header("🏛️ 系所分析")
    st.markdown("> 多班自動合併為系所，以系為單位進行完整剖析")

    # 顯示合併說明
    n_class = df["班級名稱"].nunique()
    n_dept = df["系所"].nunique()
    if n_class != n_dept:
        st.success(f"✅ 已將 **{n_class}** 個班級合併為 **{n_dept}** 個系所")
        with st.expander("🔄 查看班級→系所對照表"):
            mapping_display = (fdf[["班級名稱", "系所"]].drop_duplicates()
                               .sort_values(["系所", "班級名稱"])
                               .reset_index(drop=True))
            st.dataframe(mapping_display, use_container_width=True, hide_index=True)

    # ---- 6A：系所總覽 ----
    st.subheader("📊 各系所 KPI 總覽")
    overview_df = dept_overview_table(fdf)
    st.dataframe(overview_df, use_container_width=True, hide_index=True)
    download_excel(overview_df, "系所總覽.xlsx", "下載系所總覽")

    st.markdown("---")

    # ---- 6B：系所比較 ----
    st.subheader("📊 系所間比較")
    compare_cols = st.columns(2)
    with compare_cols[0]:
        compare_metric = st.selectbox("比較指標",
                                      ["學生數", "來源學校數", "來源縣市數",
                                       "管道數", "學校HHI"],
                                      key="dept_metric")
        st.plotly_chart(dept_compare_bar(fdf, compare_metric),
                        use_container_width=True)
    with compare_cols[1]:
        st.plotly_chart(dept_trend_lines(fdf), use_container_width=True)

    st.markdown("---")

    # ---- 6C：單一系所深入分析 ----
    st.subheader("🔍 單一系所深入分析")
    all_depts = sorted(fdf["系所"].unique().tolist())
    sel_dept = st.selectbox("選擇系所", all_depts, key="dept_select")

    # 該系 KPI
    dd = fdf[fdf["系所"] == sel_dept]
    dk1, dk2, dk3, dk4, dk5, dk6 = st.columns(6)
    dk1.metric("👥 學生數", len(dd))
    dk2.metric("📚 班級數", dd["班級名稱"].nunique())
    dk3.metric("🏫 生源學校", dd["畢業學校"].nunique())
    dk4.metric("🏙️ 來源縣市", dd["縣市"].nunique())
    dk5.metric("🎯 管道數", dd["入學方式"].nunique())
    dk6.metric("📅 學年數", dd["入學學年"].nunique())

    # 關鍵發現
    show_insights(dept_insights(fdf, sel_dept),
                  title=f"💡 【{sel_dept}】關鍵發現")

    st.markdown("---")

    # ---- 班級明細（多班時顯示）----
    if dd["班級名稱"].nunique() > 1:
        st.markdown(f"#### 📚 【{sel_dept}】各班級明細")
        st.plotly_chart(dept_class_breakdown(fdf, sel_dept),
                        use_container_width=True)
        class_summary = (dd.groupby(["入學學年", "班級名稱"]).size()
                         .reset_index(name="人數")
                         .pivot_table(index="班級名稱", columns="入學學年",
                                      values="人數", fill_value=0,
                                      margins=True, margins_name="合計"))
        st.dataframe(class_summary, use_container_width=True)
        st.markdown("---")

    # ---- 6C-1：生源地理 ----
    st.markdown(f"#### 📍 【{sel_dept}】生源地理")
    dg1, dg2 = st.columns(2)
    with dg1:
        st.plotly_chart(dept_geo_pie(fdf, sel_dept), use_container_width=True)
    with dg2:
        st.plotly_chart(dept_county_bar(fdf, sel_dept, 10),
                        use_container_width=True)

    st.markdown("---")

    # ---- 6C-2：生源學校 ----
    st.markdown(f"#### 🏫 【{sel_dept}】生源學校")
    dept_tn = st.slider("Top N 學校", 5, 20, 10, key="dept_school_tn")
    ds1, ds2 = st.columns(2)
    with ds1:
        st.plotly_chart(dept_school_bar(fdf, sel_dept, dept_tn),
                        use_container_width=True)
    with ds2:
        st.plotly_chart(dept_school_heatmap(fdf, sel_dept, dept_tn),
                        use_container_width=True)

    st.markdown("---")

    # ---- 6C-3：入學管道 ----
    st.markdown(f"#### 🎯 【{sel_dept}】入學管道")
    dc1, dc2 = st.columns(2)
    with dc1:
        st.plotly_chart(dept_channel_pie(fdf, sel_dept),
                        use_container_width=True)
    with dc2:
        st.plotly_chart(dept_channel_trend(fdf, sel_dept),
                        use_container_width=True)

    st.plotly_chart(dept_channel_school(fdf, sel_dept, min(dept_tn, 8)),
                    use_container_width=True)

    st.markdown("---")

    # ---- 6C-4：學生背景 ----
    st.markdown(f"#### 🎓 【{sel_dept}】學生背景")
    dp1, dp2 = st.columns(2)
    with dp1:
        st.plotly_chart(dept_edu_pie(fdf, sel_dept), use_container_width=True)
    with dp2:
        st.plotly_chart(dept_major_bar(fdf, sel_dept, 10),
                        use_container_width=True)

    st.markdown("---")

    # ---- 6C-5：招生趨勢 ----
    st.markdown(f"#### 📈 【{sel_dept}】招生趨勢")
    st.plotly_chart(dept_enrollment_chart(fdf, sel_dept),
                    use_container_width=True)

    # ---- 6C-6：該系原始資料 ----
    with st.expander(f"📋 【{sel_dept}】原始資料（{len(dd)} 筆）"):
        st.dataframe(dd, use_container_width=True)
        download_excel(dd, f"{sel_dept}_資料.xlsx", f"下載 {sel_dept} 資料")

# ── Tab 7：原始資料 ──
with tab7:
    st.header("📋 原始資料檢視")
    st.dataframe(fdf, use_container_width=True)
    st.caption(f"共 {len(fdf)} 筆，{len(fdf.columns)} 個欄位")
    download_excel(fdf, "篩選後資料.xlsx", "下載資料")
