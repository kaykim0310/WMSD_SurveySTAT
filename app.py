import streamlit as st

# --- 웹 페이지 기본 설정 ---
st.set_page_config(
    page_title="근골증상 및 직무스트레스 통계 시스템",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🏥 근골증상 및 직무스트레스 통계 시스템")
st.write("근골격계 증상 분석과 직무스트레스 분석을 한 곳에서 처리할 수 있습니다.")

st.markdown("""
## 시스템 사용 방법

### 1️⃣ 근골격계 증상 분석
- 좌측 사이드바에서 **📊 근골격계 증상 분석** 선택
- 템플릿 다운로드 후 데이터 입력
- 파일 업로드 및 분석 수행

### 2️⃣ 직무스트레스 분석
- 좌측 사이드바에서 **🧠 직무스트레스 분석** 선택
- 템플릿 다운로드 후 설문 데이터 입력
- 파일 업로드 및 분석 수행
- 분석 완료 후 **📈 직무스트레스 결과** 페이지에서 결과 확인 및 다운로드

### 주요 기능
- 자동 통계 계산
- 성별/공정별 분석
- Excel 형식 결과 다운로드
- 초과자 명단 자동 생성

---

💡 **Tip**: 각 분석은 독립적으로 수행되며, 결과는 세션이 유지되는 동안 보존됩니다.
""")

# 세션 상태 초기화 메시지
if 'job_stress_calculated' in st.session_state and st.session_state.job_stress_calculated:
    st.success("✅ 직무스트레스 분석이 완료되었습니다. 좌측 메뉴에서 '직무스트레스 결과'를 선택하여 결과를 확인하세요.")

if 'musculo_calculated' in st.session_state and st.session_state.musculo_calculated:
    st.success("✅ 근골격계 증상 분석이 완료되었습니다.")
