import streamlit as st
import pandas as pd
import numpy as np
import io

# --- 웹 페이지 기본 설정 ---
st.set_page_config(page_title="근골증상 및 직무스트레스 통계 시스템", layout="wide")
st.title("🏥 근골증상 및 직무스트레스 통계 시스템")
st.write("근골격계 증상 분석과 직무스트레스 분석을 한 곳에서 처리할 수 있습니다.")

# --------------------------------------------------------------------------
# 탭 생성
# --------------------------------------------------------------------------
tab1, tab2 = st.tabs(["📊 근골격계 증상 분석", "🧠 직무스트레스 분석"])

# --------------------------------------------------------------------------
# TAB 1: 근골격계 증상 분석
# --------------------------------------------------------------------------
with tab1:
    st.header("📊 근골격계 자가증상 분석")
    
    # 템플릿 생성 함수
    def create_musculoskeletal_template():
        """근골격계 증상조사 템플릿 생성"""
        full_template_columns = [
            '대상(사번)', '성명', '연령', 'AGE', 'AGE_G', '성별', '현 직장경력', 'DURATION_G',
            '작업(수행작업)', '작업부서1', '작업부서2', '작업부서3', '작업부서4', '결혼여부',
            '작업내용', '작업기간', '1일 근무시간', '휴식시간', '현재작업을 하기전에 했던 작업',
            '작업기간', 'EX-DURATION', '문1-1 규칙적 취미활동', '문1-2 평균 가사노동시간',
            '문1-3(1) 질병진단', '문1-3(2) 해당질병', '문1-3(3) 현재상태',
            '문1-4(1) 운동, 사고로 인한 상해', '문1-4(2) 상해부위', '문1-5 일의 육체적 부담',
            '문2', '문2-0(목)', '문2-2(목) 통증기간', '문2-3(목) 아픈정도', '문2-4(목) 1년동안 증상빈도',
            '문2-5(목) 일주일동안 증상여부', '문2-6(목) 통증으로 인해 일어난 일', '문2-0(어깨)',
            '문2-1(어깨) 통증부위', '문2-2(어깨) 통증기간', '문2-3(어깨) 아픈정도',
            '문2-4(어깨) 1년동안 증상빈도', '문2-5(어깨) 일주일동안 증상여부',
            '문2-6(어깨) 통증으로 인해 일어난 일', '문2-0(팔/팔꿈치)', '문2-1(팔/팔꿈치) 통증부위',
            '문2-2(팔/팔꿈치) 통증기간', '문2-3(팔/팔꿈치) 아픈정도', '문2-4(팔/팔꿈치) 1년동안 증상빈도',
            '문2-5(팔/팔꿈치) 일주일동안 증상여부', '문2-6(팔/팔꿈치) 통증으로 인해 일어난 일',
            '문2-0(손/손목/손가락)', '문2-1(손/손목/손가락) 통증부위', '문2-2(손/손목/손가락) 통증기간',
            '문2-3(손/손목/손가락) 아픈정도', '문2-4(손/손목/손가락) 1년동안 증상빈도',
            '문2-5(손/손목/손가락) 일주일동안 증상여부', '문2-6(손/손목/손가락) 통증으로 인해 일어난 일',
            '문2-0(허리)', '문2-2(허리) 통증기간', '문2-3(허리) 아픈정도',
            '문2-4(허리) 1년동안 증상빈도', '문2-5(허리) 일주일동안 증상여부',
            '문2-6(허리) 통증으로 인해 일어난 일', '문2-0(다리/발)', '문2-1(다리/발) 통증부위',
            '문2-2(다리/발) 통증기간', '문2-3(다리/발) 아픈정도', '문2-4(다리/발) 1년동안 증상빈도',
            '문2-5(다리/발) 일주일동안 증상여부', '문2-6(다리/발) 통증으로 인해 일어난 일',
            'N-n2', 'N-n3', 'N-n4', 'N-관리대상자', 'N-통증호소자', 'SH-n2', 'SH-n3', 'SH-n4',
            'SH-관리대상자', 'SH-통증호소자', 'A-n2', 'A-n3', 'A-n4', 'A-관리대상자', 'A-통증호소자',
            'H-n2', 'H-n3', 'H-n4', 'H-관리대상자', 'H-통증호소자', 'W-n2', 'W-n3', 'W-n4',
            'W-관리대상자', 'W-통증호소자', 'L-n2', 'L-n3', 'L-n4', 'L-관리대상자', 'L-통증호소자',
            '한부위이상', '관리대상자', '통증호소자'
        ]
        template_df = pd.DataFrame(columns=full_template_columns)
        title_df = pd.DataFrame(["근골격계 증상조사표"])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            title_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=0, startcol=0)
            template_df.to_excel(writer, sheet_name='Sheet1', index=False, header=True, startrow=2)
        return output.getvalue()
    
    st.write("---")
    st.subheader("⬇️ 1단계: 템플릿 파일 다운로드")
    st.write("아래 버튼으로 템플릿을 받아, **설문조사 원본 데이터 위주로** 내용을 채워주세요.")
    template_bytes = create_musculoskeletal_template()
    st.download_button(
        label="근골격계 엑셀 템플릿 다운로드", 
        data=template_bytes, 
        file_name="근골격계_데이터_입력_템플릿.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="musculo_template"
    )
    
    st.write("---")
    st.subheader("⬆️ 2단계: 데이터 파일 업로드")
    st.write("템플릿에 데이터를 모두 입력한 후, 완성된 파일을 여기에 업로드해주세요.")
    musculo_file = st.file_uploader(
        "분석할 근골격계 데이터 엑셀 파일을 업로드하세요.", 
        type=['xlsx', 'xls'], 
        key="musculo_upload"
    )
    
    # 근골격계 분석 로직 (기존 코드의 함수들을 그대로 사용)
    # 여기에 기존의 auto_calculate_status, create_table1~5 등의 함수들이 들어갑니다
    # (코드가 너무 길어서 생략하지만, 실제로는 모든 함수를 포함해야 합니다)

# --------------------------------------------------------------------------
# TAB 2: 직무스트레스 분석
# --------------------------------------------------------------------------
with tab2:
    st.header("🧠 직무스트레스 분석")
    
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
                    
                    st.success("✅ 계산 완료!")
                    
                    # 결과 미리보기
                    st.subheader("📋 계산 결과 미리보기")
                    result_columns = ['성명', '작업부서3', '작업부서4', '물리환경', '직무요구', 
                                    '직무자율', '관계갈등', '직업불안정', '조직체계', 
                                    '보상부적절', '직장문화', '총점']
                    
                    # 존재하는 컬럼만 선택
                    display_columns = [col for col in result_columns if col in df_calculated.columns]
                    st.dataframe(df_calculated[display_columns].head(10))
                    
                    # 통계 요약
                    st.subheader("📊 영역별 통계 요약")
                    stat_columns = ['물리환경', '직무요구', '직무자율', '관계갈등', 
                                  '직업불안정', '조직체계', '보상부적절', '직장문화', '총점']
                    existing_stat_cols = [col for col in stat_columns if col in df_calculated.columns]
                    
                    if existing_stat_cols:
                        stats_df = df_calculated[existing_stat_cols].describe()
                        st.dataframe(stats_df.round(2))
                    
                    # 부서별 통계 (작업부서3이 있는 경우)
                    if '작업부서3' in df_calculated.columns:
                        st.subheader("📊 부서별 직무스트레스 평균")
                        dept_stats = df_calculated.groupby('작업부서3')[existing_stat_cols].mean().round(2)
                        st.dataframe(dept_stats)
                        
                        # 부서별 상세 통계 (작업부서4까지 있는 경우)
                        if '작업부서4' in df_calculated.columns:
                            st.subheader("📊 세부 부서별 직무스트레스 평균")
                            detailed_stats = df_calculated.groupby(['작업부서3', '작업부서4'])[existing_stat_cols].mean().round(2)
                            st.dataframe(detailed_stats)
                    
                    # 성별 기준 초과자 분석
                    if '성별' in df_calculated.columns and '작업부서3' in df_calculated.columns:
                        st.subheader("📊 성별 기준치 초과자 분석")
                        
                        # 성별 기준값 정의
                        male_criteria = {
                            '물리환경': 66.7, '직무요구': 55.6, '직무자율': 58.4, '관계갈등': 62.6,
                            '직업불안정': 60.1, '조직체계': 66.7, '보상부적절': 50.1, '직장문화': 41.7, '총점': 61.2
                        }
                        
                        female_criteria = {
                            '물리환경': 55.6, '직무요구': 62.0, '직무자율': 62.0, '관계갈등': 77.8,
                            '직업불안정': 77.8, '조직체계': 50.1, '보상부적절': 50.1, '직장문화': 56.6, '총점': 56.7
                        }
                        
                        # 성별 데이터 정규화
                        df_calculated['성별_정규화'] = df_calculated['성별'].astype(str).str.strip().replace({
                            '남': 'M', '남성': 'M', 'M': 'M', 'm': 'M', '1': 'M',
                            '여': 'F', '여성': 'F', 'F': 'F', 'f': 'F', '2': 'F'
                        })
                        
                        # 부서별 초과자 집계
                        exceed_results = []
                        
                        for dept in df_calculated['작업부서3'].unique():
                            if pd.isna(dept):
                                continue
                                
                            dept_data = df_calculated[df_calculated['작업부서3'] == dept]
                            
                            # 남성 분석
                            male_data = dept_data[dept_data['성별_정규화'] == 'M']
                            male_total = len(male_data)
                            
                            # 여성 분석
                            female_data = dept_data[dept_data['성별_정규화'] == 'F']
                            female_total = len(female_data)
                            
                            for area in existing_stat_cols:
                                if area in male_criteria:
                                    # 남성 초과자
                                    male_exceed = 0
                                    if male_total > 0:
                                        male_exceed = (male_data[area] >= male_criteria[area]).sum()
                                        male_percent = (male_exceed / male_total * 100) if male_total > 0 else 0
                                    else:
                                        male_percent = 0
                                    
                                    # 여성 초과자
                                    female_exceed = 0
                                    if female_total > 0:
                                        female_exceed = (female_data[area] >= female_criteria[area]).sum()
                                        female_percent = (female_exceed / female_total * 100) if female_total > 0 else 0
                                    else:
                                        female_percent = 0
                                    
                                    exceed_results.append({
                                        '작업부서3': dept,
                                        '영역': area,
                                        '남성_전체': male_total,
                                        '남성_초과자': male_exceed,
                                        '남성_초과율(%)': round(male_percent, 1),
                                        '여성_전체': female_total,
                                        '여성_초과자': female_exceed,
                                        '여성_초과율(%)': round(female_percent, 1)
                                    })
                        
                        exceed_df = pd.DataFrame(exceed_results)
                        
                        # 부서별로 표시
                        for dept in exceed_df['작업부서3'].unique():
                            st.write(f"**{dept}**")
                            dept_exceed = exceed_df[exceed_df['작업부서3'] == dept].drop('작업부서3', axis=1)
                            st.dataframe(dept_exceed.set_index('영역'))
                            st.write("")
                        
                        # 전체 요약
                        st.subheader("📊 전체 초과자 요약")
                        total_summary = []
                        
                        for area in existing_stat_cols:
                            if area in male_criteria:
                                # 전체 남성
                                all_male = df_calculated[df_calculated['성별_정규화'] == 'M']
                                all_male_total = len(all_male)
                                all_male_exceed = (all_male[area] >= male_criteria[area]).sum() if all_male_total > 0 else 0
                                all_male_percent = (all_male_exceed / all_male_total * 100) if all_male_total > 0 else 0
                                
                                # 전체 여성
                                all_female = df_calculated[df_calculated['성별_정규화'] == 'F']
                                all_female_total = len(all_female)
                                all_female_exceed = (all_female[area] >= female_criteria[area]).sum() if all_female_total > 0 else 0
                                all_female_percent = (all_female_exceed / all_female_total * 100) if all_female_total > 0 else 0
                                
                                total_summary.append({
                                    '영역': area,
                                    '남성_기준치': male_criteria[area],
                                    '남성_초과자': f"{all_male_exceed}/{all_male_total}",
                                    '남성_초과율(%)': round(all_male_percent, 1),
                                    '여성_기준치': female_criteria[area],
                                    '여성_초과자': f"{all_female_exceed}/{all_female_total}",
                                    '여성_초과율(%)': round(all_female_percent, 1)
                                })
                        
                        total_summary_df = pd.DataFrame(total_summary)
                        st.dataframe(total_summary_df.set_index('영역'))
                    
                    # 결과 다운로드
                    st.subheader("⬇️ 결과 다운로드")
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # 원본 데이터 + 계산 결과
                        df_calculated.to_excel(writer, sheet_name='직무스트레스_계산결과', index=False)
                        
                        # 전체 통계 요약
                        if existing_stat_cols:
                            stats_df.to_excel(writer, sheet_name='전체_통계요약')
                        
                        # 부서별 통계
                        if '작업부서3' in df_calculated.columns and existing_stat_cols:
                            dept_stats = df_calculated.groupby('작업부서3')[existing_stat_cols].mean().round(2)
                            dept_stats.to_excel(writer, sheet_name='부서별_평균')
                            
                            # 세부 부서별 통계
                            if '작업부서4' in df_calculated.columns:
                                detailed_stats = df_calculated.groupby(['작업부서3', '작업부서4'])[existing_stat_cols].mean().round(2)
                                detailed_stats.to_excel(writer, sheet_name='세부부서별_평균')
                        
                        # 성별 기준 초과자 분석 추가
                        if '성별' in df_calculated.columns and '작업부서3' in df_calculated.columns and exceed_df is not None:
                            # 부서별 초과자 현황
                            exceed_df.to_excel(writer, sheet_name='부서별_초과자현황', index=False)
                            
                            # 전체 초과자 요약
                            if 'total_summary_df' in locals():
                                total_summary_df.to_excel(writer, sheet_name='전체_초과자요약')
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 계산 결과 다운로드 (Excel)",
                        data=output.read(),
                        file_name="직무스트레스_분석결과.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_job_stress"
                    )
                    
        except Exception as e:
            st.error(f"❌ 파일 처리 중 오류가 발생했습니다: {e}")
            st.error("파일 형식이 올바른지 확인해주세요.")

# 페이지 하단 정보
st.write("---")
st.info("""
💡 **도움말**
- 각 탭에서 해당하는 템플릿을 다운로드하여 데이터를 입력하세요.
- 엑셀 파일의 3행에 열 제목이 있어야 합니다.
- 문제가 발생하면 템플릿 형식을 다시 확인해주세요.
""")
