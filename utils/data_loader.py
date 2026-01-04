import pandas as pd
import streamlit as st


@st.cache_data
def load_노임단가():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSlIUPyOxmtCRrXFqQKZ7Ge3um3xi5VCaua1OvyC27Y7vw5jqJhzbFpnTeb-fcxGS3_wNxuhnBddRl4/pub?output=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    return df

@st.cache_data
def load_손해보험요율():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRzdleYSG38-1FpxjoIkQbhWHbwY4himRBCO7LR8wWkCg0bnplhSTecIHNInZ-5NCjcjfuwmGotRFd_/pub?output=csv"
    df = pd.read_csv(url)
    return df

@st.cache_data
def load_조경_기준(설계유형: str) -> pd.DataFrame:
    """조경 기준표를 Google Sheets에서 불러옵니다."""
    직급리스트 = ["기술사", "특급기술자", "고급기술자", "중급기술자", "초급기술자"]
    urls = {
        "기본설계": "https://docs.google.com/spreadsheets/d/e/2PACX-1vSffous-aCPOAcKkizEiELMpZVECskizIxlP2Vn_eHTfLnviFFCn0S1fAZPy0OkFLE508TspBu9VuuV/pub?output=csv",
        "실시설계": "https://docs.google.com/spreadsheets/d/e/2PACX-1vSRBhcxu6BMlio-obyGAj44PhEP07BPAFC9l53gad1TqZPgQyAkj289qqshKNFQ1jHYYtIrWlO9wKOm/pub?output=csv",
        "기본 및 실시설계": "https://docs.google.com/spreadsheets/d/e/2PACX-1vTcmEUxkny-pnOAPFvb67DH-MpINOZY6PqCGz9m6U3DUzFcTeqgd7Mvm7Ss1_m7i0RYE4locXoE1HuK/pub?output=csv"
    }
    df = pd.read_csv(urls.get(설계유형, urls["기본설계"]))
    df.columns = [c.strip() for c in df.columns]
    for 직급 in 직급리스트:
        if 직급 in df.columns:
            df[직급] = pd.to_numeric(df[직급], errors="coerce").fillna(0.0)
    return df

@st.cache_data
def load_환경_기준(설계유형: str) -> pd.DataFrame:
    """환경영향평가 기준표를 Google Sheets에서 불러옵니다."""
    직급리스트 = ["기술사", "특급기술자", "고급기술자", "중급기술자", "초급기술자"]
    urls = {
        "소규모 환경영향평가 대행": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQK25xZ-K2mo1mpr_Sz4cUEv8lHW6gY0-Ps0BVW-GSFRk7_53WBVHXBIBOQChTrtdYaJIP8T1p1U0sX/pub?output=csv",
    }
    url = urls.get(설계유형, urls["소규모 환경영향평가 대행"])
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    for 직급 in 직급리스트:
        if 직급 in df.columns:
            df[직급] = pd.to_numeric(df[직급], errors="coerce").fillna(0.0)
    return df