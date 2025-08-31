import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="직무스트레스 결과", layout="wide")
st.title("📈 직무스트레스 분석 결과")

# 세션 상태 확인
if 'job_stress_calculated' not in st.session_state or not st.session_state.job_stress_calculated:
    st.warning("⚠️ 먼저 직무스트레스 분석을 수행해주세요.")
    st.info("좌측 메뉴에서 '직무스트레스 분석'을 선택하여 데이터를 업로드하고 분석을 시작하세요.")
    st.stop()

# 데이터 로드
df_calculated = st.session_state.df_calculated
existing_stat_cols = st.session_state.existing_stat_cols

# 성별 통계 표시
if '성별' in df_calculated.columns:
    st.header("📊 성별 직무스트레스 통계")
    
    gender_stats = []
    
    for gender, gender_name in [('M', '남'), ('F', '여')]:
        gender_data = df_calculated[df_calculated['성별_정규화'] == gender]
        
        if len(gender_data) > 0:
            # 평균
            mean_values = gender_data[existing_stat_cols].mean()
            mean_row = {'성별': gender_name, '통계': '평균'}
            for col in existing_stat_cols:
                mean_row[col] = round(mean_values[col], 2) if not pd.isna(mean_values[col]) else '-'
            gender_stats.append(mean_row)
            
            # 표준편차
            std_values = gender_data[existing_stat_cols].std()
            std_row = {'성별': gender_name, '통계': '표준편차'}
            for col in existing_stat_cols:
                std_row[col] = round(std_values[col], 2) if not pd.isna(std_values[col]) else '-'
            gender_stats.append(std_row)
    
    if gender_stats:
        gender_stats_df = pd.DataFrame(gender_stats)
        gender_stats_df = gender_stats_df.set_index(['성별', '통계'])
        st.dataframe(gender_stats_df)

# 공정별 성별 통계
if '작업부서3' in df_calculated.columns and '작업부서1' in df_calculated.columns:
    st.header("📊 공정별 성별 직무스트레스 통계")
    
    # 성별 기준값 정의
    male_criteria = {
        '물리환경': 66.7, '직무요구': 55.6, '직무자율': 58.4, '관계갈등': 62.6,
        '직업불안정': 60.1, '조직체계': 66.7, '보상부적절': 50.1, '직장문화': 41.7, '총점': 61.2
    }
    
    female_criteria = {
        '물리환경': 55.6, '직무요구': 62.0, '직무자율': 62.0, '관계갈등': 77.8,
        '직업불안정': 77.8, '조직체계': 50.1, '보상부적절': 50.1, '직장문화': 56.6, '총점': 56.7
    }
    
    for dept in df_calculated['작업부서3'].unique():
        if pd.isna(dept):
            continue
        
        dept_data = df_calculated[df_calculated['작업부서3'] == dept]
        # 작업부서1/작업부서2/작업부서3 형식으로 표시
        dept1 = dept_data['작업부서1'].iloc[0] if len(dept_data) > 0 else ''
        dept2 = dept_data['작업부서2'].iloc[0] if len(dept_data) > 0 and '작업부서2' in dept_data.columns else ''
        display_name = f"{dept1}/{dept2}/{dept}"
        
        st.subheader(f"**{display_name}**")
        
        dept_gender_stats = []
        for gender, gender_name in [('M', '남'), ('F', '여')]:
            gender_dept_data = dept_data[dept_data['성별_정규화'] == gender]
            
            if len(gender_dept_data) > 0:
                # 평균
                mean_values = gender_dept_data[existing_stat_cols].mean()
                mean_row = {'성별': gender_name, '통계': '평균'}
                for col in existing_stat_cols:
                    mean_row[col] = round(mean_values[col], 2) if not pd.isna(mean_values[col]) else '-'
                dept_gender_stats.append(mean_row)
                
                # 표준편차
                std_values = gender_dept_data[existing_stat_cols].std()
                std_row = {'성별': gender_name, '통계': '표준편차'}
                for col in existing_stat_cols:
                    std_row[col] = round(std_values[col], 2) if not pd.isna(std_values[col]) else '-'
                dept_gender_stats.append(std_row)
                
                # 초과자수
                exceed_row = {'성별': gender_name, '통계': '초과자수'}
                exceed_rate_row = {'성별': gender_name, '통계': '초과율(%)'}
                
                criteria = male_criteria if gender == 'M' else female_criteria
                total_count = len(gender_dept_data)
                
                for col in existing_stat_cols:
                    if col in criteria:
                        exceed_count = (gender_dept_data[col] >= criteria[col]).sum()
                        exceed_row[col] = exceed_count
                        exceed_rate_row[col] = round((exceed_count / total_count * 100), 1) if total_count > 0 else 0
                    else:
                        exceed_row[col] = '-'
                        exceed_rate_row[col] = '-'
                
                dept_gender_stats.append(exceed_row)
                dept_gender_stats.append(exceed_rate_row)
        
        if dept_gender_stats:
            dept_gender_df = pd.DataFrame(dept_gender_stats)
            dept_gender_df = dept_gender_df.set_index(['성별', '통계'])
            st.dataframe(dept_gender_df)

# 결과 다운로드
st.header("⬇️ 결과 다운로드")

col1, col2 = st.columns(2)

# 자동계산 결과 엑셀
with col1:
    st.subheader("📥 자동계산 결과")
    
    # 자동계산 엑셀 생성 함수
    def create_auto_calc_excel():
        output_calc = io.BytesIO()
        with pd.ExcelWriter(output_calc, engine='xlsxwriter') as writer:
            # 원본 데이터 + 계산 결과
            df_calculated.to_excel(writer, sheet_name='직무스트레스_계산결과', index=False)
            
            # 한 항목 이상 초과자 명단 추가
            if '성별' in df_calculated.columns and existing_stat_cols:
                # 성별 기준값 정의
                male_criteria = {
                    '물리환경': 66.7, '직무요구': 55.6, '직무자율': 58.4, '관계갈등': 62.6,
                    '직업불안정': 60.1, '조직체계': 66.7, '보상부적절': 50.1, '직장문화': 41.7, '총점': 61.2
                }
                
                female_criteria = {
                    '물리환경': 55.6, '직무요구': 62.0, '직무자율': 62.0, '관계갈등': 77.8,
                    '직업불안정': 77.8, '조직체계': 50.1, '보상부적절': 50.1, '직장문화': 56.6, '총점': 56.7
                }
                
                # 초과자 찾기
                exceed_list = []
                for idx, row in df_calculated.iterrows():
                    gender = row.get('성별_정규화', '')
                    
                    if gender == 'M':
                        criteria = male_criteria
                    elif gender == 'F':
                        criteria = female_criteria
                    else:
                        continue
                    
                    # 각 항목별 초과 여부 확인
                    is_exceed = False
                    for area in existing_stat_cols:
                        if area in criteria and row[area] >= criteria[area]:
                            is_exceed = True
                            break
                    
                    if is_exceed:
                        exceed_row = {
                            '대상': row.get('대상', ''),
                            '성명': row.get('성명', ''),
                            '연령': row.get('연령', ''),
                            '성별': row.get('성별', ''),
                            '작업부서1': row.get('작업부서1', ''),
                            '작업부서2': row.get('작업부서2', ''),
                            '작업부서3': row.get('작업부서3', '')
                        }
                        
                        # 각 영역 점수 추가
                        for area in existing_stat_cols:
                            exceed_row[area] = round(row[area], 2) if not pd.isna(row[area]) else '-'
                        
                        exceed_list.append(exceed_row)
                
                if exceed_list:
                    exceed_df = pd.DataFrame(exceed_list)
                    exceed_df.to_excel(writer, sheet_name='초과자명단', index=False)
        
        output_calc.seek(0)
        return output_calc.read()
    
    calc_excel = create_auto_calc_excel()
    st.download_button(
        label="📥 자동계산 결과 다운로드",
        data=calc_excel,
        file_name="직무스트레스_자동계산결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_calc"
    )

# 통계분석 결과 엑셀
with col2:
    st.subheader("📊 통계분석 결과")
    
    # 통계분석 엑셀 생성 함수
    def create_stats_excel():
        output_stats = io.BytesIO()
        all_stats_rows = []
        
        # 성별 기준값 정의
        male_criteria = {
            '물리환경': 66.7, '직무요구': 55.6, '직무자율': 58.4, '관계갈등': 62.6,
            '직업불안정': 60.1, '조직체계': 66.7, '보상부적절': 50.1, '직장문화': 41.7, '총점': 61.2
        }
        
        female_criteria = {
            '물리환경': 55.6, '직무요구': 62.0, '직무자율': 62.0, '관계갈등': 77.8,
            '직업불안정': 77.8, '조직체계': 50.1, '보상부적절': 50.1, '직장문화': 56.6, '총점': 56.7
        }
        
        with pd.ExcelWriter(output_stats, engine='xlsxwriter') as writer:
            # 전체 성별 통계 추가
            if '성별_정규화' in df_calculated.columns:
                for gender, gender_name in [('M', '남'), ('F', '여')]:
                    gender_data = df_calculated[df_calculated['성별_정규화'] == gender]
                    
                    if len(gender_data) > 0:
                        # 평균
                        mean_values = gender_data[existing_stat_cols].mean()
                        mean_row = {
                            '구분': '전체',
                            '작업부서1': '-',
                            '작업부서2': '-',
                            '작업부서3': '전체',
                            '성별': gender_name,
                            '통계': '평균'
                        }
                        for col in existing_stat_cols:
                            mean_row[col] = round(mean_values[col], 2) if not pd.isna(mean_values[col]) else '-'
                        all_stats_rows.append(mean_row)
                        
                        # 표준편차
                        std_values = gender_data[existing_stat_cols].std()
                        std_row = {
                            '구분': '전체',
                            '작업부서1': '-',
                            '작업부서2': '-',
                            '작업부서3': '전체',
                            '성별': gender_name,
                            '통계': '표준편차'
                        }
                        for col in existing_stat_cols:
                            std_row[col] = round(std_values[col], 2) if not pd.isna(std_values[col]) else '-'
                        all_stats_rows.append(std_row)
            
            # 공정별 성별 통계 + 초과자 현황
            if '작업부서3' in df_calculated.columns and '작업부서1' in df_calculated.columns:
                for dept in df_calculated['작업부서3'].unique():
                    if pd.isna(dept):
                        continue
                    
                    dept_data = df_calculated[df_calculated['작업부서3'] == dept]
                    dept1 = dept_data['작업부서1'].iloc[0] if len(dept_data) > 0 else ''
                    dept2 = dept_data['작업부서2'].iloc[0] if len(dept_data) > 0 and '작업부서2' in dept_data.columns else ''
                    
                    for gender, gender_name in [('M', '남'), ('F', '여')]:
                        gender_dept_data = dept_data[dept_data['성별_정규화'] == gender]
                        
                        if len(gender_dept_data) > 0:
                            # 평균
                            mean_values = gender_dept_data[existing_stat_cols].mean()
                            mean_row = {
                                '구분': '공정별',
                                '작업부서1': dept1,
                                '작업부서2': dept2,
                                '작업부서3': dept,
                                '성별': gender_name,
                                '통계': '평균'
                            }
                            for col in existing_stat_cols:
                                mean_row[col] = round(mean_values[col], 2) if not pd.isna(mean_values[col]) else '-'
                            all_stats_rows.append(mean_row)
                            
                            # 표준편차
                            std_values = gender_dept_data[existing_stat_cols].std()
                            std_row = {
                                '구분': '공정별',
                                '작업부서1': dept1,
                                '작업부서2': dept2,
                                '작업부서3': dept,
                                '성별': gender_name,
                                '통계': '표준편차'
                            }
                            for col in existing_stat_cols:
                                std_row[col] = round(std_values[col], 2) if not pd.isna(std_values[col]) else '-'
                            all_stats_rows.append(std_row)
                            
                            # 초과자수
                            exceed_row = {
                                '구분': '공정별',
                                '작업부서1': dept1,
                                '작업부서2': dept2,
                                '작업부서3': dept,
                                '성별': gender_name,
                                '통계': '초과자수'
                            }
                            
                            # 초과율
                            exceed_rate_row = {
                                '구분': '공정별',
                                '작업부서1': dept1,
                                '작업부서2': dept2,
                                '작업부서3': dept,
                                '성별': gender_name,
                                '통계': '초과율(%)'
                            }
                            
                            criteria = male_criteria if gender == 'M' else female_criteria
                            total_count = len(gender_dept_data)
                            
                            for col in existing_stat_cols:
                                if col in criteria:
                                    exceed_count = (gender_dept_data[col] >= criteria[col]).sum()
                                    exceed_row[col] = exceed_count
                                    exceed_rate_row[col] = round((exceed_count / total_count * 100), 1) if total_count > 0 else 0
                                else:
                                    exceed_row[col] = '-'
                                    exceed_rate_row[col] = '-'
                            
                            all_stats_rows.append(exceed_row)
                            all_stats_rows.append(exceed_rate_row)
            
            # 전체 초과자 요약 추가
            if '성별_정규화' in df_calculated.columns and existing_stat_cols:
                for gender, gender_name in [('M', '남'), ('F', '여')]:
                    criteria = male_criteria if gender == 'M' else female_criteria
                    gender_data = df_calculated[df_calculated['성별_정규화'] == gender]
                    total_count = len(gender_data)
                    
                    if total_count > 0:
                        # 전체 초과자수
                        exceed_row = {
                            '구분': '전체',
                            '작업부서1': '-',
                            '작업부서2': '-',
                            '작업부서3': '전체',
                            '성별': gender_name,
                            '통계': '초과자수'
                        }
                        
                        # 전체 초과율
                        exceed_rate_row = {
                            '구분': '전체',
                            '작업부서1': '-',
                            '작업부서2': '-',
                            '작업부서3': '전체',
                            '성별': gender_name,
                            '통계': '초과율(%)'
                        }
                        
                        for col in existing_stat_cols:
                            if col in criteria:
                                exceed_count = (gender_data[col] >= criteria[col]).sum()
                                exceed_row[col] = exceed_count
                                exceed_rate_row[col] = round((exceed_count / total_count * 100), 1)
                            else:
                                exceed_row[col] = '-'
                                exceed_rate_row[col] = '-'
                        
                        all_stats_rows.append(exceed_row)
                        all_stats_rows.append(exceed_rate_row)
            
            # 하나의 시트에 모든 통계 저장
            if all_stats_rows:
                all_stats_df = pd.DataFrame(all_stats_rows)
                all_stats_df.to_excel(writer, sheet_name='통계분석결과', index=False)
        
        output_stats.seek(0)
        return output_stats.read()
    
    stats_excel = create_stats_excel()
    st.download_button(
        label="📊 통계분석 결과 다운로드",
        data=stats_excel,
        file_name="직무스트레스_통계분석결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_stats"
    )