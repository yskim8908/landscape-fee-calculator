import streamlit as st

# 모듈에서 run 함수 가져오기
from modules.urban_planning import run as run_urban
from modules.landscaping     import run as run_landscaping
from modules.environment    import run as run_environment

GA_MEASUREMENT_ID = "G-NRLLC2PD0P"

GA_SNIPPET = f"""
<script async src="https://www.googletagmanager.com/gtag/js?id={GA_MEASUREMENT_ID}"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){{dataLayer.push(arguments);}}
  gtag('js', new Date());
  gtag('config', '{GA_MEASUREMENT_ID}');
</script>
"""

def main():
    # GA4 스니펫 삽입 (페이지 로드 시 실행) - 나중에 활성화 가능
    # st.markdown(GA_SNIPPET, unsafe_allow_html=True)

    # 사이드바 상단 방문 통계 배지 (seeyoufarm 히트 카운터) - 나중에 활성화 가능
    # st.sidebar.markdown("### 📊 방문 통계")
    # st.sidebar.markdown(
    #     """
    #     <img src="https://hits.seeyoufarm.com/api/count/incr/badge.svg?url=your-app-unique-calc&count_bg=%23007ACC&title_bg=%23000000&icon=&icon_color=%23E7E7E7&title=visits&edge_flat=false" />
    #     """,
    #     unsafe_allow_html=True,
    # )
    # st.sidebar.markdown("---")
    
    st.sidebar.header("🗂️ 설계 분야 선택")
    option = st.sidebar.radio(
        "원하는 분야를 선택하세요",
        ("도시계획", "조경", "환경영향평가")
    )

    if option == "도시계획":
        run_urban()
    elif option == "조경":
        run_landscaping()
    elif option == "환경영향평가":
        run_environment()
    else:
        st.error("알 수 없는 옵션입니다.")

if __name__ == "__main__":
    main()