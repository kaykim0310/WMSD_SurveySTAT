import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="직무스트레스 분석", layout="wide")
st.title("🧠 직무스트레스 분석")

# 세션 상태 초기화
if 'job_stress_calculated' not in st.session_state:
    st.session_state.job_stress_calculated = False
if 'df_calculated' not in st.session_state:
    st.session_state.df_calculated = None
if 'existing_stat_cols' not in st.session_state:
    st.session_state.existing_stat_cols = None

# 직무스트레스 템플릿 생성 함수
def create_job_stress_template():
    """직무스트레스 설문조사 템플릿 생성"""
    template_columns = [
        '대상', '성명', '연령', '성별', '현 직장경력', 'DURATION_G', 
        '작업 (수행작업)', '작업부서1', '작업부서2', '작업부서3', '작업부서4', 
        '결혼여부', '작업내용', '작업기간', '1일 근무시간', '휴식시간', 
        '현재작업을 하기전에 했던 작업', '작업기간'
    ]
    
    # 문1-1부터 문1-43까지 추가
    for i in range(1, 44):
        template_columns.append(f'문1-{i}')
    
    # 계산될 영역들 추가
    template_columns.extend([
        '물리환경', '직무요구', '직무자율', '관계갈등', 
        '직업불안정', '조직체계', '보상부적절', '직장문화', '총점'
    ])
    
    template_df = pd.DataFrame(columns=template_columns)
    title_df = pd.DataFrame(["직무스트레스 설문조사표"])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        title_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0, startcol=0)
        template_df.to_excel(writer, sheet_name='Sheet1', index=False, header=True, startrow=2)
    return output.getvalue()

st.write("---")
st.subheader("⬇️ 1단계: 템플릿 파일 다운로드")
st.write("아래 버튼으로 템플릿을 받아, **문1-1부터 문1-43까지의 설문 응답**을 입력해주세요.")
st.write("물리환경, 직무요구 등의 점수는 자동으로 계산됩니다.")

job_template_bytes = create_job_stress_template()
st.download_button(
    label="직무스트레스 엑셀 템플릿 다운로드", 
    data=job_template_bytes, 
    file_name="직무스트레스_데이터_입력_템플릿.xlsx", 
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="job_stress_template"
)

st.write("---")
st.subheader("⬆️ 2단계: 데이터 파일 업로드")
st.write("템플릿에 설문 응답을 모두 입력한 후, 완성된 파일을 여기에 업로드해주세요.")

job_stress_file = st.file_uploader(
    "분석할 직무스트레스 데이터 엑셀 파일을 업로드하세요.", 
    type=['xlsx', 'xls'], 
    key="job_stress_upload"
)

# 계산 방법 설명
st.info("""
📌 **직무스트레스 계산 방법**

각 영역별 점수는 다음과 같이 계산됩니다:
- **물리환경**: (문1-1 ~ 문1-3 합계 - 3) / 9 × 100
- **직무요구**: (문1-4 ~ 문1-11 합계 - 8) / 24 × 100
- **직무자율**: (문1-12 ~ 문1-16 합계 - 5) / 15 × 100
- **관계갈등**: (문1-17 ~ 문1-20 합계 - 4) / 12 × 100
- **직업불안정**: (문1-21 ~ 문1-26 합계 - 6) / 18 × 100
- **조직체계**: (문1-27 ~ 문1-33 합계 - 7) / 21 × 100
- **보상부적절**: (문1-34 ~ 문1-39 합계 - 6) / 18 × 100
- **직장문화**: (문1-40 ~ 문1-43 합계 - 4) / 12 × 100
- **총점**: 8개 영역 점수의 평균
""")

# 직무스트레스 계산 함수
def calculate_job_stress(df):
    """직무스트레스 점수 계산 - 제공된 공식 사용"""
    try:
        # 물리환경: 문1-1 ~ 문1-3
        physical_cols = ['문1-1', '문1-2', '문1-3']
        if all(col in df.columns for col in physical_cols):
            df['물리환경'] = ((df[physical_cols].sum(axis=1) - 3) / (12 - 3)) * 100
        
        # 직무요구: 문1-4 ~ 문1-11
        demand_cols = [f'문1-{i}' for i in range(4, 12)]
        if all(col in df.columns for col in demand_cols):
            df['직무요구'] = ((df[demand_cols].sum(axis=1) - 8) / (32 - 8)) * 100
        
        # 직무자율: 문1-12 ~ 문1-16
        autonomy_cols = [f'문1-{i}' for i in range(12, 17)]
        if all(col in df.columns for col in autonomy_cols):
            df['직무자율'] = ((df[autonomy_cols].sum(axis=1) - 5) / (20 - 5)) * 100
        
        # 관계갈등: 문1-17 ~ 문1-20
        relationship_cols = [f'문1-{i}' for i in range(17, 21)]
        if all(col in df.columns for col in relationship_cols):
            df['관계갈등'] = ((df[relationship_cols].sum(axis=1) - 4) / (16 - 4)) * 100
        
        # 직업불안정: 문1-21 ~ 문1-26
        insecurity_cols = [f'문1-{i}' for i in range(21, 27)]
        if all(col in df.columns for col in insecurity_cols):
            df['직업불안정'] = ((df[insecurity_cols].sum(axis=1) - 6) / (24 - 6)) * 100
        
        # 조직체계: 문1-27 ~ 문1-33
        system_cols = [f'문1-{i}' for i in range(27, 34)]
        if all(col in df.columns for col in system_cols):
            df['조직체계'] = ((df[system_cols].sum(axis=1) - 7) / (28 - 7)) * 100
        
        # 보상부적절: 문1-34 ~ 문1-39
        reward_cols = [f'문1-{i}' for i in range(34, 40)]
        if all(col in df.columns for col in reward_cols):
            df['보상부적절'] = ((df[reward_cols].sum(axis=1) - 6) / (24 - 6)) * 100
        
        # 직장문화: 문1-40 ~ 문1-43
        culture_cols = [f'문1-{i}' for i in range(40, 44)]
        if all(col in df.columns for col in culture_cols):
            df['직장문화'] = ((df[culture_cols].sum(axis=1) - 4) / (16 - 4)) * 100
        
        # 총점 계산: 8개 영역의 평균
        score_columns = ['물리환경', '직무요구', '직무자율', '관계갈등', 
                       '직업불안정', '조직체계', '보상부적절', '직장문화']
        existing_scores = [col for col in score_columns if col in df.columns]
        
        if existing_scores:
            df['총점'] = df[existing_scores].mean(axis=1)
        
        return df
        
    except Exception as e:
        st.error(f"❌ 직무스트레스 계산 중 오류: {e}")
        return df

# 파일 처리
if job_stress_file is not None:
    try:
        # 파일 로딩
        df_job = pd.read_excel(job_stress_file, header=2)
        df_job.columns = df_job.columns.str.strip()
        
        st.success("✅ 파일 로딩 성공!")
        
        # 직무스트레스 계산
        if st.button("📊 직무스트레스 분석 시작", key="analyze_job_stress"):
            with st.spinner("계산 중..."):
                df_calculated = calculate_job_stress(df_job.copy())
                
                # 성별 정규화 추가
                if '성별' in df_calculated.columns:
                    df_calculated['성별_정규화'] = df_calculated['성별'].astype(str).str.strip().replace({
                        '남': 'M', '남성': 'M', 'M': 'M', 'm': 'M', '1': 'M',
                        '여': 'F', '여성': 'F', 'F': 'F', 'f': 'F', '2': 'F'
                    })
                
                # 세션 상태에 저장
                st.session_state.df_calculated = df_calculated
                st.session_state.job_stress_calculated = True
                
                stat_columns = ['물리환경', '직무요구', '직무자율', '관계갈등', 
                              '직업불안정', '조직체계', '보상부적절', '직장문화', '총점']
                st.session_state.existing_stat_cols = [col for col in stat_columns if col in df_calculated.columns]
                
                st.success("✅ 계산 완료!")
                st.info("📊 분석 결과를 확인하려면 좌측 메뉴에서 **'직무스트레스 결과'**를 선택하세요.")
                
    except Exception as e:
        st.error(f"❌ 파일 처리 중 오류가 발생했습니다: {e}")
        st.error("파일 형식이 올바른지 확인해주세요.")

# 분석 완료 상태 표시
if st.session_state.job_stress_calculated:
    st.write("---")
    st.success("✅ 직무스트레스 분석이 완료되었습니다!")
    st.info("좌측 메뉴에서 **'직무스트레스 결과'**를 선택하여 상세 분석 결과를 확인하고 다운로드하세요.")
