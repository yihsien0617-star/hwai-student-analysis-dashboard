"""
資料載入與清洗模組
"""
import pandas as pd
import streamlit as st
from config.settings import REQUIRED_COLUMNS, COUNTY_TO_REGION


class DataLoader:
    """處理資料匯入、驗證、清洗"""

    def __init__(self):
        self.raw_data = None
        self.clean_data = None
        self.validation_errors = []

    def load_excel(self, uploaded_file) -> pd.DataFrame:
        """載入 Excel 檔案，支援多個工作表"""
        try:
            # 讀取所有工作表
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names

            if len(sheet_names) == 1:
                df = pd.read_excel(uploaded_file, sheet_name=0)
            else:
                # 多工作表：合併所有工作表
                dfs = []
                for sheet in sheet_names:
                    temp_df = pd.read_excel(uploaded_file, sheet_name=sheet)
                    temp_df["資料來源工作表"] = sheet
                    dfs.append(temp_df)
                df = pd.concat(dfs, ignore_index=True)

            self.raw_data = df
            return df

        except Exception as e:
            st.error(f"❌ 檔案讀取失敗：{str(e)}")
            return pd.DataFrame()

    def validate_columns(self, df: pd.DataFrame) -> dict:
        """驗證必要欄位是否存在"""
        existing_cols = set(df.columns.tolist())
        required = set(REQUIRED_COLUMNS)

        missing = required - existing_cols
        found = required & existing_cols
        extra = existing_cols - required

        result = {
            "is_valid": len(missing) == 0,
            "missing_columns": list(missing),
            "found_columns": list(found),
            "extra_columns": list(extra),
            "total_rows": len(df),
        }

        return result

    def clean_data_pipeline(self, df: pd.DataFrame) -> pd.DataFrame:
        """資料清洗流程"""
        cleaned = df.copy()

        # 1. 去除前後空白
        for col in cleaned.select_dtypes(include=["object"]).columns:
            cleaned[col] = cleaned[col].astype(str).str.strip()

        # 2. 統一縣市名稱（臺/台統一）
        if "縣市" in cleaned.columns:
            cleaned["縣市"] = cleaned["縣市"].replace({
                "臺北市": "台北市", "臺中市": "台中市",
                "臺南市": "台南市", "臺東縣": "台東縣",
            })

        # 3. 若無「區域」欄位，自動根據縣市對照產生
        if "區域" not in cleaned.columns and "縣市" in cleaned.columns:
            cleaned["區域"] = cleaned["縣市"].map(COUNTY_TO_REGION).fillna("其他")

        # 4. 確保入學學年為數值
        if "入學學年" in cleaned.columns:
            cleaned["入學學年"] = pd.to_numeric(
                cleaned["入學學年"], errors="coerce"
            )
            cleaned = cleaned.dropna(subset=["入學學年"])
            cleaned["入學學年"] = cleaned["入學學年"].astype(int)

        # 5. 去除完全重複的列
        before_dedup = len(cleaned)
        cleaned = cleaned.drop_duplicates()
        after_dedup = len(cleaned)

        if before_dedup != after_dedup:
            st.info(f"ℹ️ 已移除 {before_dedup - after_dedup} 筆重複資料")

        self.clean_data = cleaned
        return cleaned

    def get_department_list(self, df: pd.DataFrame) -> list:
        """取得所有系所清單"""
        if "班級名稱" in df.columns:
            departments = df["班級名稱"].unique().tolist()
            return sorted(departments)
        return []

    def get_year_range(self, df: pd.DataFrame) -> tuple:
        """取得資料涵蓋的學年範圍"""
        if "入學學年" in df.columns:
            return int(df["入學學年"].min()), int(df["入學學年"].max())
        return (0, 0)

    def filter_data(self, df: pd.DataFrame, departments=None,
                    years=None) -> pd.DataFrame:
        """依條件篩選資料"""
        filtered = df.copy()

        if departments:
            filtered = filtered[filtered["班級名稱"].isin(departments)]
        if years:
            filtered = filtered[
                (filtered["入學學年"] >= years[0]) &
                (filtered["入學學年"] <= years[1])
            ]

        return filtered

    def get_summary_stats(self, df: pd.DataFrame) -> dict:
        """取得資料摘要統計"""
        stats = {
            "總學生數": len(df),
            "涵蓋學年數": df["入學學年"].nunique() if "入學學年" in df.columns else 0,
            "系所數": df["班級名稱"].nunique() if "班級名稱" in df.columns else 0,
            "來源縣市數": df["縣市"].nunique() if "縣市" in df.columns else 0,
            "來源學校數": df["畢業學校"].nunique() if "畢業學校" in df.columns else 0,
            "入學管道數": df["入學方式"].nunique() if "入學方式" in df.columns else 0,
        }
        return stats
