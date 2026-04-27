"""
分析四：學生背景輪廓分析
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from config.settings import CHART_CONFIG


class ProfileAnalyzer:
    """學生背景輪廓分析器"""

    def __init__(self, df: pd.DataFrame):
        self.df = df

    def get_education_distribution(self) -> pd.DataFrame:
        """入學前學歷分布"""
        edu_data = self.df.groupby(
            ["入學學年", "入學前學歷"]
        ).size().reset_index(name="人數")

        year_totals = edu_data.groupby("入學學年")["人數"].transform("sum")
        edu_data["佔比(%)"] = (
            edu_data["人數"] / year_totals * 100
        ).round(1)

        return edu_data

    def get_major_distribution(self, top_n: int = 10) -> pd.DataFrame:
        """畢業科系分布"""
        major_data = self.df.groupby(
            ["入學學年", "畢業科系"]
        ).size().reset_index(name="人數")

        # 找出 Top N 科系
        top_majors = (
            self.df["畢業科系"].value_counts().head(top_n).index.tolist()
        )

        major_data["畢業科系_分組"] = major_data["畢業科系"].apply(
            lambda x: x if x in top_majors else "其他"
        )

        grouped = major_data.groupby(
            ["入學學年", "畢業科系_分組"]
        )["人數"].sum().reset_index()

        return grouped

    def get_edu_major_cross(self) -> pd.DataFrame:
        """學歷 × 科系 交叉表"""
        cross = pd.crosstab(
            self.df["入學前學歷"], self.df["畢業科系"],
            margins=True, margins_name="合計",
        )
        return cross

    def get_profile_summary(self) -> pd.DataFrame:
        """學生背景綜合摘要"""
        summary = self.df.groupby("入學學年").agg(
            學生數=("學號", "count"),
            學歷類別數=("入學前學歷", "nunique"),
            科系類別數=("畢業科系", "nunique"),
            來源學校數=("畢業學校", "nunique"),
        ).reset_index()

        return summary

    # ── 圖表 ──────────────────────────────

    def plot_education_pie(self, year=None) -> go.Figure:
        """入學前學歷圓餅圖"""
        data = self.df if year is None else self.df[self.df["入學學年"] == year]
        title_year = f"{year} 學年度" if year else "全部學年度"

        edu_counts = data["入學前學歷"].value_counts().reset_index()
        edu_counts.columns = ["入學前學歷", "人數"]

        fig = px.pie(
            edu_counts, values="人數", names="入學前學歷",
            title=f"🎓 {title_year} 入學前學歷分布",
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

    def plot_education_trend(self) -> go.Figure:
        """學歷類別趨勢圖"""
        edu_data = self.get_education_distribution()

        fig = px.bar(
            edu_data, x="入學學年", y="人數", color="入學前學歷",
            title="📈 入學前學歷逐年趨勢",
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

    def plot_major_treemap(self, year=None) -> go.Figure:
        """畢業科系樹狀圖"""
        data = self.df if year is None else self.df[self.df["入學學年"] == year]
        title_year = f"{year} 學年度" if year else "全部學年度"

        major_counts = data.groupby(
            ["入學前學歷", "畢業科系"]
        ).size().reset_index(name="人數")

        fig = px.treemap(
            major_counts,
            path=["入學前學歷", "畢業科系"],
            values="人數",
            title=f"🌳 {title_year} 學生背景樹狀圖（學歷→科系）",
            color="人數",
            color_continuous_scale="Viridis",
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_major_bar(self, top_n: int = 10) -> go.Figure:
        """畢業科系排名"""
        major_counts = (
            self.df["畢業科系"].value_counts().head(top_n).reset_index()
        )
        major_counts.columns = ["畢業科系", "人數"]
        major_counts = major_counts.sort_values("人數", ascending=True)

        fig = px.bar(
            major_counts, x="人數", y="畢業科系", orientation="h",
            title=f"📋 Top {top_n} 畢業科系排名",
            text="人數",
            color="人數",
            color_continuous_scale="Purples",
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

    def get_profile_insights(self) -> list:
        """自動產生背景分析洞察"""
        insights = []

        # 學歷分布
        top_edu = self.df["入學前學歷"].value_counts()
        insights.append(
            f"🎓 最主要學歷來源為**{top_edu.index[0]}**，"
            f"佔 **{top_edu.iloc[0]/len(self.df)*100:.1f}%**"
        )

        # 科系分布
        top_major = self.df["畢業科系"].value_counts()
        insights.append(
            f"📚 最主要畢業科系為**{top_major.index[0]}**，"
            f"佔 **{top_major.iloc[0]/len(self.df)*100:.1f}%**"
        )

        # 趨勢（高中 vs 高職）
        if self.df["入學學年"].nunique() >= 2:
            yearly_edu = self.df.groupby(
                ["入學學年", "入學前學歷"]
            ).size().unstack(fill_value=0)

            for edu_type in yearly_edu.columns:
                if len(yearly_edu) >= 2:
                    first_val = yearly_edu[edu_type].iloc[0]
                    last_val = yearly_edu[edu_type].iloc[-1]
                    if first_val > 0:
                        change_pct = (last_val - first_val) / first_val * 100
                        if abs(change_pct) > 10:
                            direction = "上升" if change_pct > 0 else "下降"
                            insights.append(
                                f"📈 {edu_type}來源{direction}趨勢明顯，"
                                f"變化幅度 **{abs(change_pct):.1f}%**"
                            )

        return insights
