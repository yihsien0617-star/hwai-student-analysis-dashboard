"""
分析一：生源地理分布分析
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from config.settings import COLOR_PALETTE, CHART_CONFIG


class GeoAnalyzer:
    """生源地理分布分析器"""

    def __init__(self, df: pd.DataFrame):
        self.df = df

    def get_region_distribution(self) -> pd.DataFrame:
        """各區域學生人數與佔比"""
        region_counts = self.df.groupby(
            ["入學學年", "區域"]
        ).size().reset_index(name="人數")

        # 計算各年度佔比
        year_totals = region_counts.groupby("入學學年")["人數"].transform("sum")
        region_counts["佔比(%)"] = (
            region_counts["人數"] / year_totals * 100
        ).round(1)

        return region_counts

    def get_county_distribution(self) -> pd.DataFrame:
        """各縣市學生人數與佔比"""
        county_counts = self.df.groupby(
            ["入學學年", "區域", "縣市"]
        ).size().reset_index(name="人數")

        year_totals = county_counts.groupby("入學學年")["人數"].transform("sum")
        county_counts["佔比(%)"] = (
            county_counts["人數"] / year_totals * 100
        ).round(1)

        return county_counts.sort_values(
            ["入學學年", "人數"], ascending=[True, False]
        )

    def get_county_trend(self, top_n: int = 10) -> pd.DataFrame:
        """前 N 大縣市的逐年趨勢"""
        # 找出總人數前 N 的縣市
        top_counties = (
            self.df.groupby("縣市").size()
            .nlargest(top_n).index.tolist()
        )

        trend_data = self.df[self.df["縣市"].isin(top_counties)]
        trend = trend_data.groupby(
            ["入學學年", "縣市"]
        ).size().reset_index(name="人數")

        return trend

    def get_region_growth(self) -> pd.DataFrame:
        """各區域年成長率"""
        region_yearly = self.df.groupby(
            ["入學學年", "區域"]
        ).size().reset_index(name="人數")

        region_yearly = region_yearly.sort_values(["區域", "入學學年"])
        region_yearly["成長率(%)"] = (
            region_yearly.groupby("區域")["人數"]
            .pct_change() * 100
        ).round(1)

        return region_yearly

    # ── 圖表產生方法 ──────────────────────────────

    def plot_region_pie(self, year=None) -> go.Figure:
        """區域分布圓餅圖"""
        data = self.df if year is None else self.df[self.df["入學學年"] == year]
        title_year = f"{year} 學年度" if year else "全部學年度"

        region_counts = data["區域"].value_counts().reset_index()
        region_counts.columns = ["區域", "人數"]

        fig = px.pie(
            region_counts, values="人數", names="區域",
            title=f"📍 {title_year} 生源區域分布",
            color="區域", color_discrete_map=COLOR_PALETTE,
            hole=0.35,
        )
        fig.update_traces(
            textposition="inside",
            textinfo="percent+label+value",
            textfont_size=14,
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_region_trend(self) -> go.Figure:
        """區域趨勢堆疊長條圖"""
        region_data = self.get_region_distribution()

        fig = px.bar(
            region_data, x="入學學年", y="人數", color="區域",
            title="📈 各區域生源逐年趨勢",
            color_discrete_map=COLOR_PALETTE,
            text="人數",
            barmode="stack",
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            xaxis=dict(dtick=1, title="入學學年"),
            yaxis=dict(title="學生人數"),
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_county_bar(self, year=None, top_n: int = 15) -> go.Figure:
        """縣市排名橫條圖"""
        data = self.df if year is None else self.df[self.df["入學學年"] == year]
        title_year = f"{year} 學年度" if year else "全部學年度"

        county_counts = (
            data["縣市"].value_counts().head(top_n)
            .reset_index()
        )
        county_counts.columns = ["縣市", "人數"]
        county_counts = county_counts.sort_values("人數", ascending=True)

        fig = px.bar(
            county_counts, x="人數", y="縣市", orientation="h",
            title=f"🏙️ {title_year} Top {top_n} 生源縣市排名",
            text="人數",
            color="人數",
            color_continuous_scale="Tealgrn",
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            showlegend=False,
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_county_trend_line(self, top_n: int = 8) -> go.Figure:
        """前 N 大縣市趨勢折線圖"""
        trend = self.get_county_trend(top_n)

        fig = px.line(
            trend, x="入學學年", y="人數", color="縣市",
            title=f"📊 Top {top_n} 縣市逐年趨勢",
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

    def get_geo_insights(self) -> list:
        """自動產生地理分析洞察"""
        insights = []

        # 最大生源區域
        top_region = self.df["區域"].value_counts().index[0]
        top_region_pct = (
            self.df["區域"].value_counts(normalize=True).iloc[0] * 100
        )
        insights.append(
            f"🔵 最大生源區域為**{top_region}**，"
            f"佔整體 **{top_region_pct:.1f}%**"
        )

        # 最大生源縣市
        top_county = self.df["縣市"].value_counts().index[0]
        top_county_count = self.df["縣市"].value_counts().iloc[0]
        insights.append(
            f"🏆 最大生源縣市為**{top_county}**，"
            f"共 **{top_county_count}** 人"
        )

        # 生源集中度
        top3_pct = self.df["縣市"].value_counts(normalize=True).head(3).sum()
        if top3_pct > 0.6:
            insights.append(
                f"⚠️ 前 3 大縣市佔比達 **{top3_pct*100:.1f}%**，"
                f"生源高度集中，建議分散風險"
            )

        # 成長率分析（需多年度資料）
        if self.df["入學學年"].nunique() >= 2:
            growth = self.get_region_growth()
            latest_year = growth["入學學年"].max()
            latest_growth = growth[growth["入學學年"] == latest_year]

            for _, row in latest_growth.iterrows():
                if pd.notna(row["成長率(%)"]):
                    direction = "成長" if row["成長率(%)"] > 0 else "下降"
                    insights.append(
                        f"📈 {row['區域']}地區最新學年{direction} "
                        f"**{abs(row['成長率(%)']):.1f}%**"
                    )

        return insights
