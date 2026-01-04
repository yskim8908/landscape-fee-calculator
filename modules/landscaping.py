import streamlit as st
import pandas as pd
from io import BytesIO
import datetime as dt
from utils.data_loader import load_노임단가, load_손해보험요율, load_조경_기준
from utils.excel_builder import build_excel

def run():
    st.title("🌿 조경 설계 대가 산출 프로그램")  
    
    # ─── 1) 전역 데이터 정의 ───
    직급리스트 = ["기술사", "특급기술자", "고급기술자", "중급기술자", "초급기술자"]
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
            "단순 (야영장·둘레길·하천 경관 개선·생태통로·숲길조성)",
            "보통 (테마시설 조성·관광지·관광지 활성화 사업·가로 환경개선 등)",
            "복잡 (관광단지·동물원·골프장·스키장·2종 이상 복합 사업)",
        ],
    }

    성격_coeffs = {
        "도시공원":               1.0,
        "공동주택 및 대지의 조경": 1.1,
        "녹지 및 도시숲":          0.8,
        "주제형 사업":             1.2,
    }
    난이도_coeffs = {
        "도시공원":               {"단순":0.9, "보통":1.0, "복잡1":1.1, "복잡2":1.2},
        "공동주택 및 대지의 조경": {"보통":1.0, "복잡1":1.1, "복잡2":1.2},
        "녹지 및 도시숲":          {"단순":0.9, "보통":1.0, "복잡":1.1},
        "주제형 사업":             {"단순":0.9, "보통":1.0, "복잡":1.1},
    }  

    # ─── 2) 엑셀 빌드 함수 ───
    def build_조경_excel(template_path="templates/template_landscape.xlsx") -> BytesIO:
        갑지_데이터 = {
            "용역명": st.session_state.get("용역명_land", ""),
            "발주기관명": st.session_state.get("발주기관명_land", ""),
            "도급예정액": st.session_state.get("도급예정액_land", 0)
        }
        시트_데이터 = {
            "내역서": st.session_state.get("df_detail_land", pd.DataFrame()),
            "투입인원 및 내역": st.session_state.get("df_person_land", pd.DataFrame()),
            "투입인원수 산정기준": st.session_state.get("기준계산결과_land", pd.DataFrame()),
            "노임단가": st.session_state.get("최종_단가_land", pd.DataFrame()),
            "손해보험요율": st.session_state.get("df_ins_land", pd.DataFrame()),
        }
        return build_excel(template_path, 갑지_데이터, 시트_데이터)  

    # ─── 4) 탭별 UI 로직 ───
    tabs = st.tabs([
        "기초입력", "갑지", "내역서",
        "투입인원 및 내역", "투입인원수 산정기준",
        "노임단가", "손해보험요율"
    ])
    tab_기초입력, tab_갑지, tab_내역서, tab_투입인원및내역, tab_산정기준, tab_노임단가, tab_손해보험요율 = tabs

    # 4-1) 기초입력
    with tab_기초입력:
        st.header("기초입력")
        st.text_input("용역명", key="용역명_land")
        st.text_input("발주기관명", key="발주기관명_land")
        st.selectbox("설계유형", ["기본설계","실시설계","기본 및 실시설계"], key="설계유형_land")
        st.number_input("대상 면적 (㎡)", min_value=0.0, step=1.0, key="면적_land")
        st.selectbox("대상지 성격", list(난이도_map), key="성격_land")
        st.selectbox("업무난이도", 난이도_map.get(st.session_state["성격_land"], []), key="난이도_land")
        st.checkbox("전단계 활용 여부", key="전단계_활용")

    # 4-2) 갑지
    with tab_갑지:
        today = dt.date.today().strftime("%Y-%m-%d")
        st.markdown(f"##### 날짜: {today}")
        st.markdown(f"<h2 style='text-align:center;'>{st.session_state.get('용역명_land','')}</h2>",
                    unsafe_allow_html=True)
        if st.session_state.get("도급예정액_land", 0) > 0:
            st.download_button("⬇️ 조경 내역서 다운로드",
                               data=build_조경_excel(),
                               file_name=f"{st.session_state['용역명_land']}_조경설계.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.caption("※ 산출 완료 후 버튼이 활성화됩니다.")

    # 4-3) 내역서
    with tab_내역서:
        st.header("내역서")
        st.caption("※ 각 숫자를 수정한 뒤 **Enter** 를 눌러야 계산이 반영됩니다.")

        직접인건비 = st.session_state.get("직접인건비_land")
        if 직접인건비 is None:
            st.warning("먼저 ‘투입인원 및 내역’ 탭에서 직접인건비를 계산해 주세요.")
        else:
            제경비율   = st.number_input("제경비율 (110~120%)",     value=110.0, step=0.1)
            직접경비   = st.number_input("직접경비 금액 (원)", value=5_000_000, step=1_000)
            기술료율   = st.number_input("기술료율 (20~40%)",     value=20.0, step=0.1)
            공제율    = st.number_input("손해공제비율 (관람집회공사,0.432적용)",   value=0.432, step=0.001)
            부가세율   = st.number_input("부가가치세율 (%)",   value=10.0, step=0.1)

            제경비     = 직접인건비 * 제경비율   / 100
            기술료     = (직접인건비 + 제경비) * 기술료율 / 100
            손해공제비 = (직접인건비 + 제경비 + 직접경비 + 기술료) * 공제율   / 100
            부가세     = (직접인건비 + 제경비 + 직접경비 + 기술료 + 손해공제비) * 부가세율 / 100
            도급예정액  = 직접인건비 + 제경비 + 직접경비 + 기술료 + 손해공제비 + 부가세

            if st.button("✅ 산출 완료"):
                st.session_state["도급예정액_land"] = 도급예정액
                st.success(f"도급예정액 {도급예정액:,.0f}원이 저장되었습니다.")
            else:
                st.info("▶️ 값이 맞다면 ‘✅ 산출 완료’ 버튼을 눌러 주세요.")

            rows = [
                {"공종":"직접인건비", "규격":"-",         "수량":"-", "단위":"", "총액":직접인건비, "노무비":직접인건비, "경비":"",        "비고":""},
                {"공종":"제경비",     "규격":"직접인건비×율", "수량":"-", "단위":"", "총액":"",    "노무비":"",        "경비":제경비,    "비고":f"{제경비율}%"},
                {"공종":"직접경비",   "규격":"제출도서 인쇄",   "수량":1,   "단위":"식", "총액":"",    "노무비":"",        "경비":직접경비,  "비고":""},
                {"공종":"기술료",     "규격":"(직접인건비+제경비)×율","수량":"-", "단위":"", "총액":"",    "노무비":"",        "경비":기술료,    "비고":f"{기술료율}%"},
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
            st.session_state["df_detail_land"] = df

    # 4-4) 투입인원 및 내역
    with tab_투입인원및내역:
        st.header("투입인원 및 내역")

        기준결과    = st.session_state.get("기준계산결과_land")
        노임단가_df = st.session_state.get("최종_단가_land")

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

            st.session_state["직접인건비_land"] = sum_계

            st.session_state["df_person_land"] = final_df

            st.subheader("📊 기술자별 투입 인원 및 총액")
            st.dataframe(final_df)

    # 4-5) 산정기준
    with tab_산정기준:
        st.header("투입인원수 산정기준")
        공종        = st.session_state.get("선택공종")
        설계유형     = st.session_state.get("설계유형_land")
        대상_면적    = st.session_state.get("면적_land", 0)
        성격        = st.session_state.get("성격_land")
        전단계_활용  = st.session_state.get("전단계_활용", False)
        
        if 설계유형 in ["기본설계", "실시설계", "기본 및 실시설계"]:
            기준표 = load_조경_기준(설계유형).copy()
            for 직급 in 직급리스트:
                기준표[직급] = pd.to_numeric(기준표[직급], errors="coerce").fillna(0.0)
            if 대상_면적 <= 5000:
                환산계수 = (대상_면적 / 5000) ** 0.7
            else:
                환산계수 = (대상_면적 / 5000) ** 0.4
            full_nd_label = st.session_state.get("난이도_land", "")
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
            st.session_state["기준계산결과_land"] = 기준표

        else:
            st.info("‘기본설계’, ‘실시설계’, ‘기본 및 실시설계’ 중 하나를 모두 선택해야 계산이 표시됩니다.")

    # 4-6) 노임단가
    with tab_노임단가:
        st.header("노임단가")
        df = load_노임단가()
        st.dataframe(df)
        st.session_state["최종_단가_land"] = df

    # 4-7) 보험요율
    with tab_손해보험요율:
        st.header("손해보험요율")
        df = load_손해보험요율()
        st.dataframe(df)
        st.session_state["df_ins_land"] = df