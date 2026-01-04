import streamlit as st
import pandas as pd
from io import BytesIO
import datetime as dt
from utils.data_loader import load_노임단가, load_손해보험요율, load_환경_기준
from utils.excel_builder import build_excel
from utils.docx_builder import build_directive_docx

# ─── 환경영향평가 전용 로직 ───

직급리스트 = ["기술사", "특급기술자", "고급기술자", "중급기술자", "초급기술자"]

def generate_directive_docx_wrapper() -> BytesIO:
    """과업지시서(.docx) 생성 wrapper"""
    context = {
        "용역명":     st.session_state.get("용역명_env", ""),
        "발주기관명": st.session_state.get("발주기관명_env", ""),
        "날짜":       dt.date.today().strftime("%Y-%m-%d"),
        "과업대상지":  st.session_state.get("과업대상지_env", ""),
        "면적":       f"{st.session_state.get('면적_env', 0):,.0f}㎡",
        "과업기간":   f"{st.session_state.get('과업기간_env', 0)}일",
    }
    return build_directive_docx("templates/template_directive.docx", context)

def build_환경_excel(template_path="templates/template_env.xlsx") -> BytesIO:
    """환경영향평가 Excel 생성"""
    갑지_데이터 = {
        "용역명": st.session_state.get("용역명_env", ""),
        "발주기관명": st.session_state.get("발주기관명_env", ""),
        "도급예정액": st.session_state.get("도급예정액_env", 0)
    }
    시트_데이터 = {
        "내역서": st.session_state.get("df_detail_env", pd.DataFrame()),
        "투입인원 및 내역": st.session_state.get("투입인원DF_env", pd.DataFrame()),
        "투입인원수 산정기준": st.session_state.get("기준결과_env", pd.DataFrame()),
        "노임단가": st.session_state.get("최종_단가_env", pd.DataFrame()),
        "손해보험요율": st.session_state.get("보험요율DF_env", pd.DataFrame()),
    }
    return build_excel(template_path, 갑지_데이터, 시트_데이터)

def run():
    st.title("🌱 소규모 환경영향평가 대행 비용 산출 프로그램")

    tabs = st.tabs([
        "기초입력", "갑지", "내역서",
        "투입인원 및 내역", "투입인원수 산정기준",
        "노임단가", "손해보험요율"
    ])
    tab_기초입력, tab_갑지, tab_내역서, tab_투입인원및내역, tab_투입인원수산정기준, tab_노임단가, tab_손해보험요율 = tabs

    # ─── 기초입력 ───
    with tab_기초입력:
        st.header("기초입력")
        st.text_input("용역명", key="용역명_env")
        st.text_input("발주기관명", key="발주기관명_env")

        env_options = ["소규모 환경영향평가 대행"]
        st.radio("설계유형 선택", env_options, key="설계유형_env")
        st.text_input("과업대상지", key="과업대상지_env")
        st.number_input("대상 면적 (㎡)", min_value=0.0, step=10.0, key="면적_env")
        st.number_input("과업기간 (일수)", min_value=0, step=1, key="과업기간_env")

        st.markdown("**보정계수 산정용 추가 질문**")
        st.selectbox("1. 동·식물상 조사 특성",
                     ["생태*자연도 2·3등급 권역", "생태*자연도 1등급 권역 및 별도관리지역"],
                     key="보정_동식물")
        st.selectbox("2. 자연경관심의 대상", ["대상", "미대상"], key="보정_자연연경관심의")
        st.selectbox("3. 건강영향평가 대상", ["대상", "미대상"], key="보정_건강영향평가")
        st.selectbox("4. 수질오염총량계획 대상", ["대상", "미대상"], key="보정_수질오염총량계획")

    # ─── 갑지 ───
    with tab_갑지:
        today = dt.date.today().strftime("%Y-%m-%d")
        st.markdown(f"##### 날짜: {today}")
        st.markdown(f"<h2 style='text-align:center;'>{st.session_state.get('용역명_env','')}</h2>",
                    unsafe_allow_html=True)

        if st.session_state.get("도급예정액_env", 0) > 0:
            st.download_button("📄 과업지시서(.docx) 다운로드",
                               data=generate_directive_docx_wrapper(),
                               file_name=f"{st.session_state['용역명_env']}_과업지시서.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button("⬇️ 환경영향평가 내역서(Excel) 다운로드",
                               data=build_환경_excel(),
                               file_name=f"{st.session_state['용역명_env']}_환경영향평가_내역서.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.caption("※ 산출 완료 후 버튼이 활성화됩니다.")

    # ─── 내역서 ───
    with tab_내역서:
        st.header("내역서")
        st.caption("※ 각 숫자를 수정한 뒤 **Enter** 를 눌러야 계산이 반영됩니다.")

        직접인건비_env = st.session_state.get("직접인건비_env")
        if 직접인건비_env is None:
            st.warning("먼저 ‘투입인원 및 내역’ 탭에서 직접인건비를 계산해 주세요.")
        else:
            제경비율_env = st.number_input(
                "제경비율 (110~120%)",
                value=110.0,
                step=0.1,
                key="제경비율_env",
            )
            직접경비_env = st.number_input(
                "직접경비 금액 (원)",
                value=5_000_000,
                step=1_000,
                key="직접경비_env",
            )
            기술료율_env = st.number_input(
                "기술료율 (20~40%)",
                value=20.0,
                step=0.1,
                key="기술료율_env",
            )
            공제율_env = st.number_input(
                "손해공제비율 (%)",
                value=0.432,
                step=0.001,
                key="공제율_env",
            )
            부가세율_env = st.number_input(
                "부가가치세율 (%)",
                value=10.0,
                step=0.1,
                key="부가세율_env",
            )

            제경비_env = 직접인건비_env * 제경비율_env / 100
            기술료_env = (직접인건비_env + 제경비_env) * 기술료율_env / 100
            손해공제비_env = (
                (직접인건비_env + 제경비_env + 직접경비_env + 기술료_env)
                * 공제율_env
                / 100
            )
            부가세_env = (
                (직접인건비_env
                + 제경비_env
                + 직접경비_env
                + 기술료_env
                + 손해공제비_env)
                * 부가세율_env
                / 100
            )
            도급예정액_env = (
                직접인건비_env
                + 제경비_env
                + 직접경비_env
                + 기술료_env
                + 손해공제비_env
                + 부가세_env
            )

            if st.button("✅ 산출 완료"):
                st.session_state["도급예정액_env"] = 도급예정액_env
                st.success(f"도급예정액 {도급예정액_env:,.0f}원이 저장되었습니다.")
            else:
                st.info("▶️ 값이 맞다면 ‘✅ 산출 완료’ 버튼을 눌러 주세요.")

            rows_env = [
                {
                    "공종": "직접인건비",
                    "규격": "-",
                    "수량": "-",
                    "단위": "",
                    "총액": 직접인건비_env,
                    "노무비": 직접인건비_env,
                    "경비": "",
                    "비고": "",
                },
                {
                    "공종": "제경비",
                    "규격": "직접인건비×율",
                    "수량": "-",
                    "단위": "",
                    "총액": "",
                    "노무비": "",
                    "경비": 제경비_env,
                    "비고": f"{제경비율_env}%",
                },
                {
                    "공종": "직접경비",
                    "규격": "제출도서 인쇄",
                    "수량": 1,
                    "단위": "식",
                    "총액": "",
                    "노무비": "",
                    "경비": 직접경비_env,
                    "비고": "",
                },
                {
                    "공종": "기술료",
                    "규격": "(직접인건비+제경비)×율",
                    "수량": "-",
                    "단위": "",
                    "총액": "",
                    "노무비": "",
                    "경비": 기술료_env,
                    "비고": f"{기술료율_env}%",
                },
                {
                    "공종": "손해공제비",
                    "규격": "용역비×율",
                    "수량": "-",
                    "단위": "",
                    "총액": "",
                    "노무비": "",
                    "경비": 손해공제비_env,
                    "비고": f"{공제율_env}",
                },
                {
                    "공종": "부가가치세",
                    "규격": "합계×율",
                    "수량": "-",
                    "단위": "",
                    "총액": "",
                    "노무비": "",
                    "경비": 부가세_env,
                    "비고": f"{부가세율_env}%",
                },
                {
                    "공종": "도급예정액",
                    "규격": "",
                    "수량": "-",
                    "단위": "",
                    "총액": 도급예정액_env,
                    "노무비": "",
                    "경비": "",
                    "비고": "",
                },
            ]
            df_env = pd.DataFrame(rows_env)

            for c in ["총액", "노무비", "경비"]:
                df_env[c] = df_env[c].apply(
                    lambda x: f"{int(x):,}" if isinstance(x, (int, float)) else x
                )

            st.dataframe(df_env[[
                "공종","규격","수량","단위","총액","노무비","경비","비고"
            ]])

            st.session_state["df_detail_env"] = df_env

    with tab_투입인원및내역:
        st.header("투입인원 및 내역")

        기준결과_env = st.session_state.get("기준결과_env")
        노임단가_df_env = st.session_state.get("최종_단가_env")

        if 기준결과_env is None or 노임단가_df_env is None:
            st.warning("먼저 ‘산정기준’ 탭과 ‘노임단가’ 탭을 완료해주세요.")
        else:
            결과표 = 기준결과_env.copy()
            결과표 = 결과표[
                결과표["단위"].astype(str).str.strip() != ""
            ].reset_index(drop=True)

            for 직급 in 직급리스트:
                결과표[직급] = pd.to_numeric(결과표[직급], errors="coerce").fillna(0.0)

            is_technical = 결과표[직급리스트].sum(axis=1) > 0

            기간값 = {}
            n = len(결과표)
            half = (n + 1) // 2
            left, right = st.columns(2)

            with left:
                for idx, row in 결과표.iloc[:half].iterrows():
                    if not is_technical.iloc[idx]:
                        기간값[idx] = 0
                        continue
                    업무 = row["업무구분"]
                    단위 = str(row["단위"]).strip()

                    if 단위 == "식":
                        기본값 = 1
                        라벨 = f"{업무} (식)"
                    else:
                        기본값 = 1
                        라벨 = f"{업무} 기간 (일)"

                    값 = st.number_input(
                        라벨,
                        min_value=0,
                        step=1,
                        value=int(st.session_state.get(f"기간_{idx}", 기본값)),
                        key=f"기간_L_{idx}"
                    )
                    기간값[idx] = 값

            with right:
                for idx, row in 결과표.iloc[half:].iterrows():
                    if not is_technical.iloc[idx]:
                        기간값[idx] = 0
                        continue
                    업무 = row["업무구분"]
                    단위 = str(row["단위"]).strip()

                    if 단위 == "식":
                        기본값 = 1
                        라벨 = f"{업무} (식)"
                    else:
                        기본값 = 1
                        라벨 = f"{업무} 기간 (일)"

                    값 = st.number_input(
                        라벨,
                        min_value=0,
                        step=1,
                        value=int(st.session_state.get(f"기간_{idx}", 기본값)),
                        key=f"기간_R_{idx}"
                    )
                    기간값[idx] = 값

            결과표["기간"] = [기간값[i] for i in range(n)]

            노임단가_df_env.columns = [c.strip() for c in 노임단가_df_env.columns]
            노임단가_df_env["직종명"] = 노임단가_df_env["직종명"].astype(str).str.strip()

            env_col = "환경"
            노임단가_df_env[env_col] = (
                노임단가_df_env[env_col]
                .astype(str)
                .str.replace(",", "")
                .str.strip()
                .astype(float)
            )

            단가사전 = {}
            for 직급 in 직급리스트:
                sub = 노임단가_df_env[노임단가_df_env["직종명"] == 직급]
                단가사전[직급] = float(sub[env_col].iloc[0]) if not sub.empty else 0.0

            계산된_계 = []
            for idx, row in 결과표.iterrows():
                if not is_technical.iloc[idx]:
                    계산된_계.append(0.0)
                    continue
                인건비합 = sum(row[직급] * 단가사전.get(직급, 0.0) for 직급 in 직급리스트)
                사전체 = round(인건비합 * row["기간"], 2)
                계산된_계.append(사전체)

            결과표["계"] = 계산된_계

            표시열 = ["업무구분", "계"] + 직급리스트 + ["기간"]

            sum_계_env = 결과표["계"].sum()
            총계행 = {c: "" for c in 표시열}
            총계행["업무구분"] = "총계"
            총계행["계"] = sum_계_env
            total_df = pd.DataFrame([총계행])

            final_df = pd.concat([total_df, 결과표[표시열]], ignore_index=True)

            def fmt(x):
                if isinstance(x, (int, float)):
                    return "" if x == 0 else f"{x:,.2f}"
                return x

            for c in ["계", "기간"] + 직급리스트:
                final_df[c] = final_df[c].apply(fmt)

            st.session_state["직접인건비_env"] = sum_계_env
            st.session_state["투입인원DF_env"] = final_df

            st.subheader("📊 기술자별 투입 인원 및 총액")
            st.dataframe(final_df)

    # ─── 산정기준 탭 ───
    with tab_투입인원수산정기준:
        st.header("투입인원수 산정기준")

        설계유형_env = st.session_state.get("설계유형_env")
        대상_면적_env = st.session_state.get("면적_env", 0)
        q1 = st.session_state.get("보정_동식물")
        q2 = st.session_state.get("보정_자연연경관심의")
        q3 = st.session_state.get("보정_건강영향평가")
        q4 = st.session_state.get("보정_수질오염총량계획")

        if 설계유형_env == "소규모 환경영향평가 대행":
            기준표_env = load_환경_기준(설계유형_env).copy()
            A = float(대상_면적_env)

            α1_list = []
            for _, row in 기준표_env.iterrows():
                rule = row.get("환산계수", None)
                α1 = 1.00
                if rule == 1:
                    α1 = 1.00
                elif rule == 2:
                    if A <= 10000:
                        α1 = 0.25
                    elif A <= 100000:
                        α1 = round((A / 100000) ** 0.6, 3)
                    elif A <= 1000000:
                        α1 = round((A / 100000) ** 0.3, 3)
                    else:
                        α1 = 2.00
                elif rule == 3:
                    if A <= 30000:
                        α1 = 0.49
                    elif A <= 1000000:
                        α1 = round((A / 100000) ** 0.6, 3)
                    else:
                        α1 = 3.98
                else:
                    α1 = 1.00
                α1_list.append(α1)

            기준표_env["α₁(환산계수)"] = α1_list

            factor_가 = 1.00 if q1 == "생태*자연도 2등급 및 3등급 권역" else 1.20
            factor_나 = 1.00 if q2 == "미대상" else 2.00
            factor_다 = 1.00 if q3 == "미대상" else 1.15
            factor_라 = 1.00 if q4 == "미대상" else 1.40

            for 직급 in 직급리스트:
                계산값_list = []
                계산식_list = []
                for _, row in 기준표_env.iterrows():
                    base = row.get(직급, 0)
                    α1 = row["α₁(환산계수)"]
                    보정가 = row.get("보정계수(가)", "")
                    보정나 = row.get("보정계수(나)", "")
                    보정다 = row.get("보정계수(다)", "")
                    보정라 = row.get("보정계수(라)", "")

                    if not isinstance(base, (int, float)) or base <= 0:
                        계산값_list.append("")
                        계산식_list.append("")
                        continue

                    v = base * α1
                    parts = [f"{α1:.3f}"]

                    if 보정가 == "반영":
                        v *= factor_가
                        parts.append(f"{factor_가:.2f}")
                    if 보정나 == "반영":
                        v *= factor_나
                        parts.append(f"{factor_나:.2f}")
                    if 보정다 == "반영":
                        v *= factor_다
                        parts.append(f"{factor_다:.2f}")
                    if 보정라 == "반영":
                        v *= factor_라
                        parts.append(f"{factor_라:.2f}")

                    v_rounded = round(v, 2)
                    계산값_list.append(v_rounded)

                    formula = f"{base:.2f} × " + " × ".join(parts)
                    계산식_list.append(formula)

                기준표_env[직급] = 계산값_list
                기준표_env[f"{직급}_계산식"] = 계산식_list

            표시열 = ["업무구분", "단위"]
            for 직급 in 직급리스트:
                표시열 += [직급, f"{직급}_계산식"]

            df_display_env = 기준표_env[표시열].copy()
            df_display_env = df_display_env.replace(0, "").fillna("")

            st.subheader(f"📊 {설계유형_env} 기준표 (환산계수 + 보정계수 반영)")
            st.dataframe(df_display_env)

            st.session_state["기준결과_env"] = 기준표_env

        else:
            st.info("‘소규모 환경영향평가 대행’을 선택해야 산정기준이 표시됩니다.")


    # ─── 노임단가 ───
    with tab_노임단가:
        st.header("노임단가")
        df_wage_env = load_노임단가()
        st.dataframe(df_wage_env)
        st.session_state["최종_단가_env"] = df_wage_env

    # ─── 보험요율 ───
    with tab_손해보험요율:
        st.header("손해보험요율")
        df_ins_env = load_손해보험요율()
        st.dataframe(df_ins_env)
        st.session_state["보험요율DF_env"] = df_ins_env