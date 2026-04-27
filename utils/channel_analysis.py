"""
分析三：入學管道成效分析
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from config.settings import CHART_CONFIG


class ChannelAnalyzer:
    """入學管道分析器"""

    def __init__(self, df: pd.DataFrame):
        self.df = df

    def get_channel_distribution(self) -> pd.DataFrame:
        """各入學管道人數與佔比"""
        channel_data = self.df.groupby(
            ["入學學年", "入學方式"]
        ).size().reset_index(name="人數")

        year_totals = channel_data.groupby("入學學年")["人數"].transform("sum")
        channel_data["佔比(%)"] = (
            channel_data["人數"] / year_totals * 100
        ).round(1)

        return channel_data

    def get_channel_region_cross(self) -> pd.DataFrame:
        """入學管道 × 區域 交叉分析"""
        cross = pd.crosstab(
            self.df["入學方式"], self.df["區域"],
            margins=True, margins_name="合計",
        )
        return cross

    def get_channel_school_type_cross(self) -> pd.DataFrame:
        """入學管道 × 入學前學歷 交叉分析"""
        cross = pd.crosstab(
            self.df["入學方式"], self.df["入學前學歷"],
            margins=True, margins_name="合計",
        )
        return cross

    def get_channel_concentration(self) -> dict:
        """管道集中度風險評估"""
        channel_pct = (
            self.df["入學方式"].value_counts(normalize=True) * 100
        )

        risk_level = "低"
        if channel_pct.iloc[0] > 90:
            risk_level = "高"
        elif channel_pct.iloc[0] > 70:
            risk_level = "中"

        return {
            "主力管道": channel_pct.index[0],
            "主力管道佔比": round(channel_pct.iloc[0], 1),
            "管道數量": len(channel_pct),
            "集中度風險": risk_level,
        }

    # ── 圖表 ──────────────────────────────

    def plot_channel_pie(self, year=None) -> go.Figure:
        """入學管道圓餅圖"""
        data = self.df if year is None else self.df[self.df["入學學年"] == year]
        title_year = f"{year} 學年度" if year else "全部學年度"

        counts = data["入學方式"].value_counts().reset_index()
        counts.columns = ["入學方式", "人數"]

        fig = px.pie(
            counts, values="人數", names="入學方式",
            title=f"🎯 {title_year} 入學管道分布",
            hole=0.35,
        )
        fig.update_traces(
            textposition="inside",
            textinfo="percent+label+value",
            textfont_size=13,
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_channel_trend(self) -> go.Figure:
        """管道趨勢堆疊圖"""
        channel_data = self.get_channel_distribution()

        fig = px.bar(
            channel_data, x="入學學年", y="人數", color="入學方式",
            title="📈 各入學管道逐年趨勢",
            text="人數", barmode="stack",
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

    def plot_channel_region_heatmap(self) -> go.Figure:
        """管道 × 區域 熱力圖"""
        cross = self.get_channel_region_cross()
        display = cross.drop(index="合計", errors="ignore").drop(
            columns="合計", errors="ignore"
        )

        fig = go.Figure(data=go.Heatmap(
            z=display.values,
            x=display.columns.tolist(),
            y=display.index.tolist(),
            colorscale="Blues",
            text=display.values,
            texttemplate="%{text}",
            textfont={"size": 14},
        ))
        fig.update_layout(
            title="🗺️ 入學管道 × 區域 交叉分析",
            template=CHART_CONFIG["template"],
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def get_channel_insights(self) -> list:
        """自動產生管道分析洞察"""
        insights = []

        concentration = self.get_channel_concentration()
        insights.append(
            f"🎯 主力管道為**{concentration['主力管道']}**，"
            f"佔比 **{concentration['主力管道佔比']}%**"
        )

        if concentration["集中度風險"] == "高":
            insights.append(
                "🔴 **管道集中度風險：高** — "
                "超過 90% 學生來自單一管道，"
                "若該管道政策改變將嚴重影響招生"
            )
        elif concentration["集中度風險"] == "中":
            insights.append(
                "🟡 **管道集中度風險：中** — "
                "建議積極拓展其他入學管道"
            )
        else:
            insights.append(
                "🟢 **管道集中度風險：低** — 管道多元，風險分散良好"
            )

        insights.append(
            f"📊 目前共使用 **{concentration['管道數量']}** 種入學管道"
        )

        return insights
