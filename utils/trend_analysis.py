"""
分析五：跨年度綜合趨勢分析
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from config.settings import CHART_CONFIG


class TrendAnalyzer:
    """跨年度綜合趨勢分析器"""

    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.years = sorted(df["入學學年"].unique())

    def get_enrollment_trend(self) -> pd.DataFrame:
        """整體招生人數趨勢"""
        trend = self.df.groupby("入學學年").agg(
            學生數=("學號", "count"),
            來源學校數=("畢業學校", "nunique"),
            來源縣市數=("縣市", "nunique"),
        ).reset_index()

        if len(trend) >= 2:
            trend["年成長率(%)"] = (
                trend["學生數"].pct_change() * 100
            ).round(1)
            trend["累積變化(%)"] = (
                (trend["學生數"] / trend["學生數"].iloc[0] - 1) * 100
            ).round(1)

        return trend

    def get_department_trend(self) -> pd.DataFrame:
        """各系所招生趨勢"""
        dept_trend = self.df.groupby(
            ["入學學年", "班級名稱"]
        ).size().reset_index(name="人數")

        return dept_trend

    def get_diversity_index(self) -> pd.DataFrame:
        """多元性指標（衡量生源是否越來越集中或分散）"""
        results = []
        for year in self.years:
            year_data = self.df[self.df["入學學年"] == year]

            # 縣市 HHI (Herfindahl-Hirschman Index)
            county_shares = (
                year_data["縣市"].value_counts(normalize=True)
            )
            county_hhi = (county_shares ** 2).sum()

            # 學校 HHI
            school_shares = (
                year_data["畢業學校"].value_counts(normalize=True)
            )
            school_hhi = (school_shares ** 2).sum()

            results.append({
                "入學學年": year,
                "縣市集中度(HHI)": round(county_hhi, 4),
                "學校集中度(HHI)": round(school_hhi, 4),
                "來源縣市數": year_data["縣市"].nunique(),
                "來源學校數": year_data["畢業學校"].nunique(),
            })

        return pd.DataFrame(results)

    def get_comprehensive_summary(self) -> pd.DataFrame:
        """綜合摘要表"""
        summary_rows = []
        for year in self.years:
            year_data = self.df[self.df["入學學年"] == year]

            summary_rows.append({
                "入學學年": year,
                "總學生數": len(year_data),
                "來源縣市數": year_data["縣市"].nunique(),
                "來源學校數": year_data["畢業學校"].nunique(),
                "入學管道數": year_data["入學方式"].nunique(),
                "主要區域": year_data["區域"].value_counts().index[0],
                "主要區域佔比": f"{year_data['區域'].value_counts(normalize=True).iloc[0]*100:.1f}%",
                "主要管道": year_data["入學方式"].value_counts().index[0],
                "主要管道佔比": f"{year_data['入學方式'].value_counts(normalize=True).iloc[0]*100:.1f}%",
            })

        return pd.DataFrame(summary_rows)

    # ── 圖表 ──────────────────────────────

    def plot_enrollment_trend(self) -> go.Figure:
        """招生人數趨勢圖"""
        trend = self.get_enrollment_trend()

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        fig.add_trace(
            go.Bar(
                x=trend["入學學年"], y=trend["學生數"],
                name="學生數", marker_color="#636EFA",
                text=trend["學生數"], textposition="outside",
            ),
            secondary_y=False,
        )

        if "年成長率(%)" in trend.columns:
            fig.add_trace(
                go.Scatter(
                    x=trend["入學學年"], y=trend["年成長率(%)"],
                    name="年成長率(%)", mode="lines+markers+text",
                    text=trend["年成長率(%)"].apply(
                        lambda x: f"{x:+.1f}%" if pd.notna(x) else ""
                    ),
                    textposition="top center",
                    marker=dict(color="red", size=10),
                    line=dict(color="red", width=2),
                ),
                secondary_y=True,
            )

        fig.update_layout(
            title="📊 招生人數與成長率趨勢",
            template=CHART_CONFIG["template"],
            xaxis=dict(dtick=1, title="入學學年"),
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        fig.update_yaxes(title_text="學生人數", secondary_y=False)
        fig.update_yaxes(title_text="成長率 (%)", secondary_y=True)

        return fig

    def plot_department_trend(self) -> go.Figure:
        """各系所趨勢圖"""
        dept_data = self.get_department_trend()

        fig = px.line(
            dept_data, x="入學學年", y="人數", color="班級名稱",
            title="📈 各系所招生趨勢",
            markers=True,
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            xaxis=dict(dtick=1),
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_diversity_trend(self) -> go.Figure:
        """多元性指標趨勢"""
        diversity = self.get_diversity_index()

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        fig.add_trace(
            go.Scatter(
                x=diversity["入學學年"], y=diversity["縣市集中度(HHI)"],
                name="縣市集中度", mode="lines+markers",
                line=dict(color="#636EFA"),
            ),
            secondary_y=False,
        )
        fig.add_trace(
            go.Scatter(
                x=diversity["入學學年"], y=diversity["來源學校數"],
                name="來源學校數", mode="lines+markers",
                line=dict(color="#00CC96"),
            ),
            secondary_y=True,
        )

        fig.update_layout(
            title="🔄 生源多元性趨勢（HHI 越低越分散）",
            template=CHART_CONFIG["template"],
            xaxis=dict(dtick=1),
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        fig.update_yaxes(title_text="HHI 指數", secondary_y=False)
        fig.update_yaxes(title_text="來源學校數", secondary_y=True)

        return fig

    def plot_dashboard_kpi(self) -> dict:
        """產生 KPI 指標數據"""
        trend = self.get_enrollment_trend()

        latest = trend.iloc[-1]
        kpi = {
            "最新學年": int(latest["入學學年"]),
            "最新學生數": int(latest["學生數"]),
            "來源學校數": int(latest["來源學校數"]),
            "來源縣市數": int(latest["來源縣市數"]),
        }

        if len(trend) >= 2:
            prev = trend.iloc[-2]
            kpi["學生數變化"] = int(latest["學生數"] - prev["學生數"])
            kpi["學校數變化"] = int(
                latest["來源學校數"] - prev["來源學校數"]
            )
            kpi["成長率"] = round(
                (latest["學生數"] - prev["學生數"]) / prev["學生數"] * 100, 1
            )
        else:
            kpi["學生數變化"] = 0
            kpi["學校數變化"] = 0
            kpi["成長率"] = 0

        return kpi

    def get_trend_insights(self) -> list:
        """自動產生趨勢分析洞察"""
        insights = []
        trend = self.get_enrollment_trend()

        if len(trend) >= 2:
            latest = trend.iloc[-1]
            growth = latest.get("年成長率(%)", 0)

            if pd.notna(growth):
                if growth > 5:
                    insights.append(
                        f"✅ 最新學年招生成長 **{growth:.1f}%**，"
                        "趨勢正向"
                    )
                elif growth < -5:
                    insights.append(
                        f"⚠️ 最新學年招生下降 **{abs(growth):.1f}%**，"
                        "需關注並採取行動"
                    )
                else:
                    insights.append(
                        f"➡️ 最新學年招生變化 **{growth:+.1f}%**，"
                        "大致持平"
                    )

            # 整體趨勢
            first_year_count = trend.iloc[0]["學生數"]
            last_year_count = trend.iloc[-1]["學生數"]
            total_change = (
                (last_year_count - first_year_count) / first_year_count * 100
            )
            insights.append(
                f"📈 {int(trend.iloc[0]['入學學年'])} 至 "
                f"{int(trend.iloc[-1]['入學學年'])} 學年，"
                f"整體變化 **{total_change:+.1f}%**"
            )

        # 多元性趨勢
        diversity = self.get_diversity_index()
        if len(diversity) >= 2:
            hhi_change = (
                diversity.iloc[-1]["縣市集中度(HHI)"]
                - diversity.iloc[0]["縣市集中度(HHI)"]
            )
            if hhi_change > 0.02:
                insights.append(
                    "🔴 生源集中度上升，建議拓展新的招生區域"
                )
            elif hhi_change < -0.02:
                insights.append(
                    "🟢 生源越來越分散多元，招生佈局良好"
                )

        return insights
