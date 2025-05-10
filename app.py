import streamlit as st
import pandas as pd

@st.cache_data
def load_기준인원수(설계유형):
    if 설계유형 == "기본설계":
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSffous-aCPOAcKkizEiELMpZVECskizIxlP2Vn_eHTfLnviFFCn0S1fAZPy0OkFLE508TspBu9VuuV/pub?output=csv"
    elif 설계유형 == "실시설계":
        url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSRBhcxu6BMlio-obyGAj44PhEP07BPAFC9l53gad1TqZPgQyAkj289qqshKNFQ1jHYYtIrWlO9wKOm/pub?output=csv"
    else:  # "기본 및 실시설계"
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

    공사명 = st.text_input("공사명", value=st.session_state.get("공사명", "")) 
    st.session_state["공사명"] = 공사명

    발주기관명 = st.text_input("발주기관명", value=st.session_state.get("발주기관명", ""))
    st.session_state["발주기관명"] = 발주기관명

    공종_선택 = st.selectbox("공종을 선택하세요", ["조경"])
    st.session_state["선택공종"] = 공종_선택

    if 공종_선택 == "조경":
        설계유형 = st.radio("설계유형을 선택하세요",
                           ["기본설계", "실시설계", "기본 및 실시설계"],
                           index=(["기본설계","실시설계","기본 및 실시설계"]
                                  .index(st.session_state.get("설계유형","기본설계"))))
        st.session_state["설계유형"] = 설계유형
    else:
        st.session_state["설계유형"] = None

    면적 = st.number_input("대상 면적 (㎡)",
                     min_value=100.0, step=100.0,
                     value=st.session_state.get("면적",100.0))
    st.session_state["면적"] = 면적

    대상지_성격 = st.selectbox("대상지 성격",
                        ["도시공원", "공동주택 및 대지의 조경", "녹지", "주제형"],
                        index=(["도시공원","공동주택 및 대지의 조경","녹지","주제형"]
                               .index(st.session_state.get("대상지_성격","도시공원"))))
    st.session_state["대상지_성격"] = 대상지_성격

    난이도 = st.selectbox("업무 난이도",
                    ["단순", "보통", "복잡"],
                    index=(["단순","보통","복잡"]
                           .index(st.session_state.get("난이도","보통"))))
    st.session_state["난이도"] = 난이도

    전단계_활용 = st.checkbox(
        "기본계획 등 설계에 활용할 전 단계 성과물이 있습니까?", 
        value=False
    )
    st.session_state["전단계_활용"] = 전단계_활용

with tab_갑지:
    import datetime
    today = datetime.date.today().strftime("%Y-%m-%d")

    st.markdown(f"##### 날짜: {today}")

    st.markdown(
        f"<h2 style='text-align:center;'>{공사명}</h2>",
        unsafe_allow_html=True
    )

    raw = st.session_state.get("도급예정액", 0)
    truncated = int(raw // 1000) * 1000
    st.write(f"**용역비:** {truncated:,.0f} 원")

    발주기관 = st.session_state.get("발주기관명", "")
    st.write(f"**발주기관:** {발주기관}")

    

with tab_내역서:
    st.header("내역서")

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
        st.session_state["도급예정액"] = 도급예정액

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

with tab_투입인원및내역:
    st.header("투입인원 및 내역")

    기준결과    = st.session_state.get("기준계산결과")
    노임단가_df = st.session_state.get("최종_단가")

    if 기준결과 is None or 노임단가_df is None:
        st.warning("먼저 '투입인원수 산정기준'과 '노임단가' 탭을 완료해주세요.")
    else:
        결과표 = 기준결과.copy()

        횟수_키워드 = ["위원회 심의", "주민설명회", "관계기관 협의"]
        기간값 = {}
        n = len(결과표)
        half = (n + 1) // 2
        left, right = st.columns(2)

        with left:
            for idx, row in 결과표.iloc[:half].iterrows():
                업무 = row["업무구분"]
                if any(kw in 업무 for kw in 횟수_키워드):
                    if "주민설명회" in 업무:
                        기본값 = 2
                    else: 
                        기본값 = 1   
                    라벨 = f"{업무} (회)"
                else:
                    기본값 = 5
                    라벨 = f"{업무} 기간 (일)"
                값 = st.number_input(
                    라벨,
                    min_value=1, step=1,
                    value=int(st.session_state.get(f"기간_{idx}", 기본값)),
                    key=f"기간_=L_{idx}"
                )
                기간값[idx] = 값

        with right:
            for idx, row in 결과표.iloc[half:].iterrows():
                업무 = row["업무구분"]
                if any(kw in 업무 for kw in 횟수_키워드):
                    if "주민설명회" in 업무:
                        기본값 = 2
                    else:    
                        기본값 = 1
                    라벨 = f"{업무} (회)"
                else:
                    기본값 = 5
                    라벨 = f"{업무} 기간 (일)"
                값 = st.number_input(
                    라벨,
                    min_value=1, step=1,
                    value=int(st.session_state.get(f"기간_{idx}", 기본값)),
                    key=f"기간_R_{idx}"
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

        st.subheader("📊 기술자별 투입 인원 및 총액")
        st.dataframe(final_df)
        
with tab_산정기준:
    st.header("투입인원수 산정기준")

     # 사용자 입력값 불러오기
    공종        = st.session_state.get("선택공종")
    설계유형     = st.session_state.get("설계유형")
    면적        = st.session_state.get("면적", 0)
    성격        = st.session_state.get("대상지_성격")
    난이도       = st.session_state.get("난이도")
    전단계_활용  = st.session_state.get("전단계_활용", False)

    # 조경 + 세 가지 설계유형 모두 계산
    if 공종 == "조경" and 설계유형 in ["기본설계", "실시설계", "기본 및 실시설계"]:
        기준표 = load_기준인원수(설계유형).copy()
        환산계수 = (면적 / 5000) ** 0.4 if 면적 > 0 else 0
        성격계수 = {"도시공원":1.0, "공동주택":1.1, "녹지":0.8, "주제형":1.2}
        난이도계수 = {"단순":0.9, "보통":1.0, "복잡":1.1}

        for 직급 in 직급리스트:
            계산값, 계산식 = [], []
            for _, row in 기준표.iterrows():
                base = row[직급]
                v = base
                parts = []

                if row["환산계수(α₁)"] == "적용":
                    v *= 환산계수; parts.append(f"{환산계수:.3f}")

                if row["보정계수(α₂, α₃)"] == "적용" and row["업무구분"] not in ["조사", "기술협의"]:
                    a2 = 성격계수.get(성격,1.0); a3 = 난이도계수.get(난이도,1.0)
                    v *= a2 * a3; parts += [f"{a2:.3f}", f"{a3:.3f}"]

                if 전단계_활용 and row["업무구분"] == "설계안 작성":
                    v *= 0.7; parts.append("0.700")

                formula = f"{base} × " + " × ".join(parts) if parts else f"{base} (고정)"
                계산값.append(round(v,2)); 계산식.append(formula)
            기준표[직급] = 계산값
            기준표[f"{직급}_계산식"] = 계산식

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
        st.success("✅ 공제요율 정보를 불러왔습니다.")
        st.dataframe(공제요율_df)
    except Exception as e:
        st.error("❌ 공제요율 정보를 불러오지 못했습니다.")
        st.text(f"에러 메시지: {e}")
        st.stop()