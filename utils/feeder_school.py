"""
分析二：餵校（Feeder School）分析
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from config.settings import CHART_CONFIG


class FeederSchoolAnalyzer:
    """餵校分析器"""

    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.years = sorted(df["入學學年"].unique())

    def get_school_ranking(self, top_n: int = 20) -> pd.DataFrame:
        """來源學校排名"""
        ranking = (
            self.df.groupby(["畢業學校"])
            .agg(
                總人數=("學號", "count"),
                涵蓋學年數=("入學學年", "nunique"),
                來源區域=("區域", "first"),
                來源縣市=("縣市", "first"),
            )
            .reset_index()
            .sort_values("總人數", ascending=False)
            .head(top_n)
        )

        ranking["平均每年送生"] = (
            ranking["總人數"] / ranking["涵蓋學年數"]
        ).round(1)
        ranking["排名"] = range(1, len(ranking) + 1)

        return ranking

    def get_school_yearly_detail(self, top_n: int = 20) -> pd.DataFrame:
        """各校逐年送生明細"""
        top_schools = (
            self.df["畢業學校"].value_counts().head(top_n).index.tolist()
        )

        detail = (
            self.df[self.df["畢業學校"].isin(top_schools)]
            .groupby(["畢業學校", "入學學年"])
            .size()
            .reset_index(name="人數")
        )

        # 轉成寬表（學校 × 學年）
        pivot = detail.pivot_table(
            index="畢業學校", columns="入學學年",
            values="人數", fill_value=0,
        )
        pivot["合計"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("合計", ascending=False)

        return pivot

    def classify_schools(self) -> dict:
        """將學校分類：穩定校 / 成長校 / 流失校 / 新進校"""
        total_years = len(self.years)
        school_yearly = (
            self.df.groupby(["畢業學校", "入學學年"])
            .size().reset_index(name="人數")
        )

        classifications = {
            "穩定餵校": [],
            "成長中學校": [],
            "流失中學校": [],
            "新進學校": [],
            "偶發學校": [],
        }

        for school in self.df["畢業學校"].unique():
            school_data = school_yearly[
                school_yearly["畢業學校"] == school
            ].sort_values("入學學年")

            years_present = len(school_data)
            total_students = school_data["人數"].sum()

            if total_years <= 1:
                # 只有一年資料，無法判斷趨勢
                if total_students >= 3:
                    classifications["穩定餵校"].append({
                        "學校": school, "總人數": total_students,
                        "出現年數": years_present,
                    })
                continue

            # 穩定校：每年都有送生且平均 >= 2 人
            if (years_present == total_years and
                    total_students / total_years >= 2):
                classifications["穩定餵校"].append({
                    "學校": school, "總人數": total_students,
                    "出現年數": years_present,
                })

            # 成長校：最近學年人數 > 前一學年
            elif years_present >= 2:
                recent = school_data.iloc[-1]["人數"]
                previous = school_data.iloc[-2]["人數"]
                if recent > previous:
                    classifications["成長中學校"].append({
                        "學校": school, "總人數": total_students,
                        "出現年數": years_present,
                        "最近成長": f"{previous}→{recent}",
                    })
                elif recent < previous:
                    classifications["流失中學校"].append({
                        "學校": school, "總人數": total_students,
                        "出現年數": years_present,
                        "最近變化": f"{previous}→{recent}",
                    })

            # 新進校：只在最近一年出現
            elif (years_present == 1 and
                    school_data.iloc[0]["入學學年"] == self.years[-1]):
                classifications["新進學校"].append({
                    "學校": school, "總人數": total_students,
                })

            else:
                classifications["偶發學校"].append({
                    "學校": school, "總人數": total_students,
                    "出現年數": years_present,
                })

        # 轉成 DataFrame
        result = {}
        for category, schools in classifications.items():
            if schools:
                result[category] = pd.DataFrame(schools).sort_values(
                    "總人數", ascending=False
                )
            else:
                result[category] = pd.DataFrame()

        return result

    # ── 圖表 ──────────────────────────────

    def plot_school_ranking(self, top_n: int = 20) -> go.Figure:
        """來源學校排名橫條圖"""
        ranking = self.get_school_ranking(top_n)
        ranking = ranking.sort_values("總人數", ascending=True)

        fig = px.bar(
            ranking, x="總人數", y="畢業學校", orientation="h",
            title=f"🏫 Top {top_n} 來源學校排名",
            text="總人數",
            color="來源區域",
            hover_data=["來源縣市", "涵蓋學年數", "平均每年送生"],
        )
        fig.update_layout(
            template=CHART_CONFIG["template"],
            height=max(500, top_n * 30),
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_school_heatmap(self, top_n: int = 15) -> go.Figure:
        """學校 × 學年 熱力圖"""
        pivot = self.get_school_yearly_detail(top_n)
        display_data = pivot.drop(columns=["合計"], errors="ignore")

        fig = go.Figure(data=go.Heatmap(
            z=display_data.values,
            x=[str(c) for c in display_data.columns],
            y=display_data.index.tolist(),
            colorscale="YlOrRd",
            text=display_data.values,
            texttemplate="%{text}",
            textfont={"size": 12},
        ))

        fig.update_layout(
            title="🔥 來源學校 × 學年度 送生熱力圖",
            xaxis_title="入學學年",
            yaxis_title="畢業學校",
            template=CHART_CONFIG["template"],
            height=max(500, top_n * 35),
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def plot_school_concentration(self) -> go.Figure:
        """學校集中度分析（累積佔比曲線）"""
        school_counts = self.df["畢業學校"].value_counts()
        cumulative_pct = school_counts.cumsum() / school_counts.sum() * 100

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=list(range(1, len(cumulative_pct) + 1)),
            y=cumulative_pct.values,
            mode="lines+markers",
            name="累積佔比",
            marker=dict(size=4),
        ))
        fig.add_hline(y=80, line_dash="dash", line_color="red",
                      annotation_text="80% 門檻")

        fig.update_layout(
            title="📊 來源學校集中度分析（累積佔比曲線）",
            xaxis_title="學校數量（依送生人數排序）",
            yaxis_title="累積佔比 (%)",
            template=CHART_CONFIG["template"],
            font=dict(
                family=CHART_CONFIG["font_family"],
                size=CHART_CONFIG["font_size"],
            ),
        )
        return fig

    def get_feeder_insights(self) -> list:
        """自動產生餵校分析洞察"""
        insights = []

        total_schools = self.df["畢業學校"].nunique()
        total_students = len(self.df)
        insights.append(
            f"🏫 共有 **{total_schools}** 所來源學校，"
            f"平均每校送生 **{total_students/total_schools:.1f}** 人"
        )

        # Top 10 學校佔比
        top10_count = self.df["畢業學校"].value_counts().head(10).sum()
        top10_pct = top10_count / total_students * 100
        insights.append(
            f"📌 前 10 大來源學校共送生 **{top10_count}** 人，"
            f"佔 **{top10_pct:.1f}%**"
        )

        # 單人學校比例
        single_schools = (
            self.df["畢業學校"].value_counts() == 1
        ).sum()
        single_pct = single_schools / total_schools * 100
        if single_pct > 50:
            insights.append(
                f"⚠️ 僅送 1 人的學校佔 **{single_pct:.1f}%**，"
                f"生源極度分散，建議聚焦經營重點餵校"
            )

        # 分類結果摘要
        classifications = self.classify_schools()
        for category, school_df in classifications.items():
            if not school_df.empty:
                insights.append(
                    f"{'🟢' if '穩定' in category else '🟡' if '成長' in category else '🔴' if '流失' in category else '🔵'} "
                    f"{category}：**{len(school_df)}** 所"
                )

        return insights
