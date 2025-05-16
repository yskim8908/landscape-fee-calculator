import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import shutil
from openpyxl import load_workbook
import datetime as dt

import sqlite3
from pathlib import Path

DB_PATH = Path("visits.db")
conn = sqlite3.connect(DB_PATH, check_same_thread=False)
cur  = conn.cursor()
cur.execute("""
CREATE TABLE IF NOT EXISTS visits (
    timestamp TEXT
)
""")
conn.commit()

def record_visit():
    now = dt.datetime.now().isoformat()
    cur.execute("INSERT INTO visits (timestamp) VALUES (?)", (now,))
    conn.commit()

def get_today_count():
    today = dt.date.today().isoformat()
    cur.execute(
        "SELECT COUNT(*) FROM visits WHERE substr(timestamp,1,10)=?",
        (today,)
    )
    return cur.fetchone()[0]

def get_total_count():
    cur.execute("SELECT COUNT(*) FROM visits")
    return cur.fetchone()[0]

if not st.session_state.get("visit_recorded", False):
    record_visit()
    st.session_state["visit_recorded"] = True

st.sidebar.markdown(f"👀 오늘 누적 방문 수: **{get_today_count()}** 회")
st.sidebar.markdown(f"🕒 전체 누적 방문 수: **{get_total_count()}** 회")

난이도_map = {
    "도시공원": [
        "단순 (소공원·묘지공원·보행자 전용도로·광장·도시공원 내 시설 교체사업)",
        "보통 (국가도시공원·근린공원·체육공원·수변공원·도시농업공원·유원지·공공공지·광장(재생사업))",
        "복잡1 (어린이공원·문화공원·역사공원·방재공원)",
        "복잡2 (도시공원(재생사업))",
    ],
    "공동주택 및 대지의 조경": [
        "보통 (공동주택 조경)",
        "복잡1 (주택정원·건축물 조경·옥상조경(옥상정원))",
        "복잡2 (실내조경(실내정원))",
    ],
    "녹지 및 도시숲": [
        "단순 (완충 녹지·가로변 녹지·가로수·경관숲)",
        "보통 (연결 녹지·경관 녹지·유휴지 녹화·마을숲·유아숲체험원)",
        "복잡 (가로변 녹지(정원형)·학교숲·도시숲)",
    ],
    "주제형 사업": [
        "단순 (야영장·둘레길·하천 경관 개선, 생태통로, 숲길조성)",
        "보통 (테마시설 조성·관광지·관광지 활성화 사업·가로 환경개선 등)",
        "복잡 (관광단지·동물원·골프장·스키장·2종 이상 복합 사업)",
    ],
}

성격_coeffs = {
    "도시공원":             1.0,
    "공동주택 및 대지의 조경": 1.1,
    "녹지 및 도시숲":        0.8,
    "주제형 사업":           1.2,
}

# ── 계산용: 대상지 성격별 난이도계수 α₃ ──
난이도_coeffs = {
    "도시공원": {
        "단순": 0.9,
        "보통": 1.0,
        "복잡1": 1.1,
        "복잡2": 1.2,
    },
    "공동주택 및 대지의 조경": {
        "보통": 1.0,
        "복잡1": 1.1,
        "복잡2": 1.2,
    },
    "녹지 및 도시숲": {
        "단순": 0.9,
        "보통": 1.0,
        "복잡": 1.1,
    },
    "주제형 사업": {
        "단순": 0.9,
        "보통": 1.0,
        "복잡": 1.1,
    },
}

def build_excel_overlay(template_path="template.xlsx") -> BytesIO:
    # 1) 템플릿 파일을 안전하게 복사
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    shutil.copy(template_path, tmp.name)

    # 2) openpyxl 로 '갑지' 시트만 직접 값 채우기
    wb = load_workbook(tmp.name)
    ws_cover = wb["갑지"]
    ws_cover["D10"].value = st.session_state.get("용역명", "")
    ws_cover["G22"].value = st.session_state.get("발주기관명", "")
    raw = st.session_state.get("도급예정액", 0)
    ws_cover["G20"].value = f"{int(raw//1000)*1000:,} 원"
    ws_cover["A1"].value = dt.date.today().strftime("%Y-%m-%d")
    wb.save(tmp.name)  # 여긴 openpyxl 방식으로 덮어쓴 뒤 저장

    # 3) pandas ExcelWriter 를 overlay 모드로 열어서
    #    나머지 시트에 DataFrame 값만 덮어쓰기
    with pd.ExcelWriter(tmp.name,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="overlay") as writer:

        # A) 내역서
        df_detail = st.session_state.get("df_detail", pd.DataFrame())
        if not df_detail.empty:
            df_detail.to_excel(
                writer,
                sheet_name="내역서",
                index=False,
                header=False,    # 템플릿의 1행 헤더 아래부터 덮어쓰기
                startrow=2
            )

        # B) 투입인원 및 내역
        df_person = st.session_state.get("투입인원DF", pd.DataFrame())
        if not df_person.empty:
            df_person.to_excel(
                writer,
                sheet_name="투입인원 및 내역",
                index=False,
                header=False,
                startrow=2
            )

        # C) 투입인원수 산정기준
        df_basis = st.session_state.get("기준계산결과", pd.DataFrame())
        if not df_basis.empty:
            df_basis.to_excel(
                writer,
                sheet_name="투입인원수 산정기준",
                index=False,
                header=False,
                startrow=2
            )

        # D) 노임단가
        df_wage = st.session_state.get("최종_단가", pd.DataFrame())
        if not df_wage.empty:
            df_wage.to_excel(
                writer,
                sheet_name="노임단가",
                index=False,
                header=False,
                startrow=2
            )

        # E) 손해보험요율
        df_ins = st.session_state.get("보험요율DF", pd.DataFrame())
        if not df_ins.empty:
            df_ins.to_excel(
                writer,
                sheet_name="손해보험요율",
                index=False,
                header=False,
                startrow=2
            )

    # 4) 완성된 파일을 BytesIO 로 읽어서 반환
    buf = BytesIO()
    with open(tmp.name, "rb") as f:
        buf.write(f.read())
    buf.seek(0)
    return buf

@st.cache_data
def load_기준인원수(설계유형):
    if 설계유형 == "기본설계":
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSffous-aCPOAcKkizEiELMpZVECskizIxlP2Vn_eHTfLnviFFCn0S1fAZPy0OkFLE508TspBu9VuuV/pub?output=csv"
    elif 설계유형 == "실시설계":
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSRBhcxu6BMlio-obyGAj44PhEP07BPAFC9l53gad1TqZPgQyAkj289qqshKNFQ1jHYYtIrWlO9wKOm/pub?output=csv"
    else: 
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTcmEUxkny-pnOAPFvb67DH-MpINOZY6PqCGz9m6U3DUzFcTeqgd7Mvm7Ss1_m7i0RYE4locXoE1HuK/pub?output=csv"
    return pd.read_csv(url)

직급리스트 = ["기술사", "특급기술자", "고급기술자", "중급기술자", "초급기술자"]

st.title("실시설계 용역 대가 산출 프로그램")

(
    tab_기초입력,
    tab_갑지,
    tab_내역서,
    tab_투입인원및내역,
    tab_산정기준,
    tab_노임단가,
    tab_손해보험요율
) = st.tabs([
    "기초입력",
    "갑지",
    "내역서",
    "투입인원 및 내역",
    "투입인원수 산정기준",
    "노임단가",
    "손해보험요율"
])

with tab_기초입력:

    st.header("기초입력")

    용역명 = st.text_input("용역명", value=st.session_state.get("용역명", "")) 
    st.session_state["용역명"] = 용역명

    발주기관명 = st.text_input("발주기관명", value=st.session_state.get("발주기관명", ""))
    st.session_state["발주기관명"] = 발주기관명

    공종_선택 = st.selectbox("공종을 선택하세요", ["조경"])
    st.session_state["선택공종"] = 공종_선택

    if 공종_선택 == "조경":
        options = [
            "기본설계",
            "실시설계",
            "기본 및 실시설계",
        ]
        current = st.session_state.get("설계유형", "기본설계")
        index = options.index(current) if current in options else 0

        설계유형 = st.radio(
            "설계유형을 선택하세요",
            options,
            index=index,
            key="설계유형_radio"
        )
        st.session_state["설계유형"] = 설계유형
    else:
        st.session_state["설계유형"] = None

    면적 = st.number_input("대상 면적 (㎡)",
                     min_value=100.0, step=100.0,
                     value=st.session_state.get("면적",100.0))
    st.session_state["면적"] = 면적

    성격_options = [
        "도시공원",
        "공동주택 및 대지의 조경",
        "녹지 및 도시숲",
        "주제형 사업"
    ]
    default_성격 = st.session_state.get("대상지_성격", "도시공원")
    if default_성격 not in 성격_options:                 
        default_성격 = "도시공원"

    대상지_성격 = st.selectbox(
       "대상지 성격",
       성격_options,
       index=성격_options.index(default_성격)
    )
    st.session_state["대상지_성격"] = 대상지_성격

    options_nd = 난이도_map.get(대상지_성격, ["단순","보통","복잡"])
    prev_nd    = st.session_state.get("난이도", options_nd[0])
    if prev_nd not in options_nd:
        prev_nd = options_nd[0]

    난이도 = st.selectbox(
        "업무 난이도",
        options_nd,
        index=options_nd.index(prev_nd),
        key="난이도"
    )

    전단계_활용 = st.checkbox(
        "기본계획 등 설계에 활용할 전 단계 성과물이 있습니까?", 
        value=False
    )
    st.session_state["전단계_활용"] = 전단계_활용

    if st.button("🔄  입력값 모두 초기화", help="용역명·면적 등 기초입력과 계산 결과를 지웁니다."):
        reset_keys = [
            "용역명", "발주기관명",
            "선택공종", "설계유형",
            "면적", "대상지_성격", "난이도", "전단계_활용",
            "기준계산결과", "직접인건비", "도급예정액",
        ]
        reset_keys += [k for k in st.session_state if k.startswith("기간_")]

        for k in reset_keys:
            st.session_state.pop(k, None)    

        st.rerun()   

with tab_갑지:
    import datetime
    today = datetime.date.today().strftime("%Y-%m-%d")

    st.markdown(f"##### 날짜: {today}")

    st.markdown(
        f"<h2 style='text-align:center;'>{용역명}</h2>",
        unsafe_allow_html=True
    )

    if "도급예정액" not in st.session_state:
        st.info("먼저 ‘내역서’ 탭에서 **산출 완료✅** 버튼을 눌러 금액을 확정하세요.")
    else:
        raw        = st.session_state["도급예정액"]
        용역비      = int(raw // 1000) * 1000   
        st.write(f"**용역비:** {용역비:,.0f} 원")

    발주기관 = st.session_state.get("발주기관명", "")
    st.write(f"**발주기관:** {발주기관}")

    if "도급예정액" in st.session_state and st.session_state["도급예정액"] > 0:
        excel_buf = build_excel_overlay("template.xlsx")
        st.download_button(
            label="⬇️ 갑지(Excel) 다운로드",
            data=excel_buf,
            file_name=f"{st.session_state['용역명']}_갑지.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.caption("※ 산출 완료 후 버튼이 활성화됩니다.")

with tab_내역서:
    st.header("내역서")
    st.caption("※ 각 숫자를 수정한 뒤 **Enter** 를 눌러야 계산이 반영됩니다.")

    직접인건비 = st.session_state.get("직접인건비")
    if 직접인건비 is None:
        st.warning("먼저 ‘투입인원 및 내역’ 탭에서 직접인건비를 계산해 주세요.")
    else:
        제경비율   = st.number_input("제경비율 (110~120%)",     value=110.0, step=0.1)
        직접경비   = st.number_input("직접경비 금액 (원)", value=5_000_000, step=1_000)
        기술료율   = st.number_input("기술료율 (20~40%)",     value=20.0, step=0.1)
        공제율    = st.number_input("손해공제비율 (관람집회공사,0.432적용)",   value=0.432, step=0.001)
        부가세율   = st.number_input("부가가치세율 (%)",   value=10.0, step=0.1)

        제경비     = 직접인건비 * 제경비율   / 100
        기술료     = (직접인건비 + 제경비 + 직접경비) * 기술료율 / 100
        손해공제비 = (직접인건비 + 제경비 + 직접경비 + 기술료) * 공제율   / 100
        부가세     = (직접인건비 + 제경비 + 직접경비 + 기술료 + 손해공제비) * 부가세율 / 100
        도급예정액  = 직접인건비 + 제경비 + 직접경비 + 기술료 + 손해공제비 + 부가세

        if st.button("✅ 산출 완료"):
            st.session_state["도급예정액"] = 도급예정액
            st.success(f"도급예정액 {도급예정액:,.0f}원이 저장되었습니다.")
        else:
            st.info("▶️ 값이 맞다면 ‘✅ 산출 완료’ 버튼을 눌러 주세요.")

        rows = [
            {"공종":"직접인건비", "규격":"-",         "수량":"-", "단위":"", "총액":직접인건비, "노무비":직접인건비, "경비":"",        "비고":""},
            {"공종":"제경비",     "규격":"직접인건비×율", "수량":"-", "단위":"", "총액":"",    "노무비":"",        "경비":제경비,    "비고":f"{제경비율}%"},
            {"공종":"직접경비",   "규격":"제출도서 인쇄",   "수량":1,   "단위":"식", "총액":"",    "노무비":"",        "경비":직접경비,  "비고":""},
            {"공종":"기술료",     "규격":"인건비+제경비×율","수량":"-", "단위":"", "총액":"",    "노무비":"",        "경비":기술료,    "비고":f"{기술료율}%"},
            {"공종":"손해공제비", "규격":"용역비×율",      "수량":"-", "단위":"", "총액":"",    "노무비":"",        "경비":손해공제비,"비고":f"{공제율}"},
            {"공종":"부가가치세", "규격":"합계×율",       "수량":"-", "단위":"", "총액":"",    "노무비":"",        "경비":부가세,    "비고":f"{부가세율}%"},
            {"공종":"도급예정액","규격":"",             "수량":"-", "단위":"", "총액":도급예정액,"노무비":"",        "경비":"",        "비고":""},
        ]
        df = pd.DataFrame(rows)

        for c in ["총액","노무비","경비"]:
            df[c] = df[c].apply(lambda x: f"{int(x):,}" if isinstance(x,(int,float)) else x)

        st.dataframe(df[[
            "공종","규격","수량","단위",
            "총액","노무비","경비","비고"
        ]])
        st.session_state["df_detail"] = df

with tab_투입인원및내역:
    st.header("투입인원 및 내역")

    기준결과    = st.session_state.get("기준계산결과")
    노임단가_df = st.session_state.get("최종_단가")

    if 기준결과 is None or 노임단가_df is None:
        st.warning("먼저 '투입인원수 산정기준'과 '노임단가' 탭을 완료해주세요.")
    else:
        결과표 = 기준결과.copy()
        결과표 = 결과표[결과표["단위"] != ""].reset_index(drop=True)
        for 직급 in 직급리스트:
            결과표[직급] = pd.to_numeric(결과표[직급], errors='coerce').fillna(0.0)

        횟수_키워드 = ["위원회 심의", "주민설명회", "관계기관 협의"]
        기간값 = {}
        n = len(결과표)
        half = (n + 1) // 2
        left, right = st.columns(2)

        with left:
            for idx, row in 결과표.iloc[:half].iterrows():
                업무 = row["업무구분"]
                단위 = row["단위"].strip()  # 단위 칼럼 읽어오기

                if 단위 == "식":
                    기본값 = 1
                    라벨 = f"{업무} (식)"

                elif any(kw in 업무 for kw in 횟수_키워드):
                    if "주민설명회" in 업무:
                        기본값 = 2
                    else:
                        기본값 = 1
                    라벨 = f"{업무} (회)"

                else:
                    기본값 = 2
                    라벨 = f"{업무} 기간 (일)"

                값 = st.number_input(
                   라벨,
                   min_value=0,
                   step=1,
                   value=int(st.session_state.get(f"기간_{idx}", 기본값)),
                   key=f"기간_=L_{idx}"
                )
                기간값[idx] = 값

        with right:
            for idx, row in 결과표.iloc[half:].iterrows():
                업무 = row["업무구분"]
                단위 = row["단위"].strip()  
                if 단위 == "식":
                    기본값 = 1
                    라벨 = f"{업무} (식)"
                else:
                    기본값 = 2
                    라벨 = f"{업무} 기간 (일)"

                값 = st.number_input(
                   라벨,
                   min_value=0,
                   step=1,
                   value=int(st.session_state.get(f"기간_{idx}", 기본값)),
                   key=f"기간_=L_{idx}"
                )
                기간값[idx] = 값

        결과표["기간"] = [기간값[i] for i in range(n)]

        노임단가_df.columns   = [c.strip() for c in 노임단가_df.columns]
        노임단가_df["직종명"] = 노임단가_df["직종명"].astype(str).str.strip()
        노임단가_df["건설"]   = (
            노임단가_df["건설"]
              .astype(str)
              .str.replace(",", "")
              .str.strip()
              .astype(float)
        )
        직급리스트 = ["기술사","특급기술자","고급기술자","중급기술자","초급기술자"]
        건설단가 = {}
        for 직급 in 직급리스트:
            sub = 노임단가_df[노임단가_df["직종명"] == 직급]
            건설단가[직급] = float(sub["건설"].iloc[0]) if not sub.empty else 0.0


        계산된_계 = []
        for _, row in 결과표.iterrows():
            인건비합 = sum(row[직급] * 건설단가[직급] for 직급 in 직급리스트)
            계산된_계.append(round(인건비합 * row["기간"], 2))
        결과표["계"] = 계산된_계

        표시열   = ["업무구분","계"] + 직급리스트 + ["기간"]
        sum_계   = 결과표["계"].sum()
        총계행    = {c: "" for c in 표시열}
        총계행["업무구분"] = "총계"
        총계행["계"]    = sum_계
        total_df = pd.DataFrame([총계행])
        final_df = pd.concat([total_df, 결과표[표시열]], ignore_index=True)

        for c in ["계","기간"] + 직급리스트:
            final_df[c] = final_df[c].apply(
                lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x
            )

        sum_계 = 결과표["계"].sum()

        st.session_state["직접인건비"] = sum_계

        st.session_state["투입인원DF"] = final_df

        st.subheader("📊 기술자별 투입 인원 및 총액")
        st.dataframe(final_df)
        
with tab_산정기준:
    st.header("투입인원수 산정기준")

    공종        = st.session_state.get("선택공종")
    설계유형     = st.session_state.get("설계유형")
    대상_면적    = st.session_state.get("면적", 0)
    성격        = st.session_state.get("대상지_성격")
    전단계_활용  = st.session_state.get("전단계_활용", False)

    if 공종 == "조경" and 설계유형 in ["기본설계", "실시설계", "기본 및 실시설계"]:
        기준표 = load_기준인원수(설계유형).copy()
        for 직급 in 직급리스트:
            기준표[직급] = pd.to_numeric(기준표[직급], errors="coerce").fillna(0.0)
        if 면적 <= 5000:
            환산계수 = (대상_면적 / 5000) ** 0.7
        else:
            환산계수 = (대상_면적 / 5000) ** 0.4
        full_nd_label = st.session_state.get("난이도", "")
        diff_key = full_nd_label.split()[0] if full_nd_label else ""
        a2 = 성격_coeffs.get(성격, 1.0)
        a3 = 난이도_coeffs.get(성격, {}).get(diff_key, 1.0)

        for 직급 in 직급리스트:
            계산값, 계산식 = [], []
            for _, row in 기준표.iterrows():
                base = row[직급]
                v = base
                parts = []

                if row["환산계수(α₁)"] == "적용":
                    v *= 환산계수; parts.append(f"{환산계수:.3f}")

                if row["보정계수(α₂, α₃)"] == "적용" and row["업무구분"] not in ["조사", "기술협의"]:
                    v *= a2 * a3
                    parts.append(f"{a2:.3f}")  # 성격계수
                    parts.append(f"{a3:.3f}")  # 난이도계수


                if 전단계_활용:
                    first_token = str(row["업무구분"]).strip().split()[0]
                    if first_token.startswith("2.1"):        
                        v *= 0.7
                        parts.append("0.700")

                formula = f"{base} × " + " × ".join(parts) if parts else f"{base} (고정)"
                계산값.append(round(v,2)); 계산식.append(formula)
            기준표[직급] = 계산값
            기준표[f"{직급}_계산식"] = 계산식

        기준표 = 기준표.fillna("") 

        for 직급 in 직급리스트:
            calc_col = f"{직급}_계산식"
            if calc_col in 기준표.columns:
                기준표[calc_col] = 기준표[calc_col].str.replace(r"\s*\(고정\)", "", regex=True)
        mask = 기준표['단위'] == ""
        cols = 직급리스트 + [f"{j}_계산식" for j in 직급리스트]
        기준표.loc[mask, cols] = ""

        st.subheader(f"📊 {설계유형} 계산된 투입인원 (인·일)")
        표시열 = ["업무구분", "단위"] + sum([[j, f"{j}_계산식"] for j in 직급리스트], [])
        st.dataframe(기준표[표시열])
        st.session_state["기준계산결과"] = 기준표

    else:
        st.info("‘조경’과 ‘기본설계’, ‘실시설계’, ‘기본 및 실시설계’ 중 하나를 모두 선택해야 계산이 표시됩니다.")

with tab_노임단가:
    st.header("노임단가")

    sheet_url = (
        "https://docs.google.com/spreadsheets/d/e/"
        "2PACX-1vSlIUPyOxmtCRrXFqQKZ7Ge3um3xi5VCaua1OvyC27Y7vw5jqJhzbFpnTeb-fcxGS3_wNxuhnBddRl4"
        "/pub?output=csv"
    )

    try:
        기본_단가_df = pd.read_csv(sheet_url)
    except Exception as e:
        st.error("❌ 기본 단가를 불러오지 못했습니다.")
        st.text(f"에러 메시지: {e}")
        st.stop()

    기본_단가_df.columns = [c.strip() for c in 기본_단가_df.columns]
    if "직종" in 기본_단가_df.columns and "직종명" not in 기본_단가_df.columns:
        기본_단가_df = 기본_단가_df.rename(columns={"직종": "직종명"})

    st.dataframe(기본_단가_df)

    st.session_state["최종_단가"] = 기본_단가_df

with tab_손해보험요율:
    st.header("손해보험요율")

    insurance_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRzdleYSG38-1FpxjoIkQbhWHbwY4himRBCO7LR8wWkCg0bnplhSTecIHNInZ-5NCjcjfuwmGotRFd_/pub?output=csv"

    try:
        공제요율_df = pd.read_csv(insurance_url)
        st.session_state["보험요율DF"] = 공제요율_df
        st.success("✅ 공제요율 정보를 불러왔습니다.")

        st.dataframe(공제요율_df)
    except Exception as e:
        st.error("❌ 공제요율 정보를 불러오지 못했습니다.")
        st.text(f"에러 메시지: {e}")
        st.stop()