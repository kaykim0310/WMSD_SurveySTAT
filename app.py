import streamlit as st
import pandas as pd
import numpy as np
import io

# --- 웹 페이지 기본 설정 ---
st.set_page_config(page_title="근골격계 자가증상 분석", layout="wide")
st.title("📊 근골격계 자가증상 분석 프로그램")

# --------------------------------------------------------------------------
# 1. 템플릿 생성 및 다운로드 기능
# --------------------------------------------------------------------------
def create_template_excel():
    """사용자가 제공한 전체 열 목록을 기반으로 템플릿 엑셀 파일을 생성하는 함수"""
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
template_bytes = create_template_excel()
st.download_button(label="엑셀 템플릿 다운로드", data=template_bytes, file_name="데이터_입력_템플릿.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.write("---")

st.subheader("⬆️ 2단계: 데이터 파일 업로드")
st.write("템플릿에 데이터를 모두 입력한 후, 완성된 파일을 여기에 업로드해주세요.")
uploaded_file = st.file_uploader("분석할 데이터 엑셀 파일을 업로드하세요.", type=['xlsx', 'xls'], label_visibility='collapsed')

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='자동계산결과', index=False)
    return output.getvalue()

def results_to_excel_bytes(results_dict):
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            sheet_name, start_row = '통합 분석결과', 0
            workbook = writer.book
            header_format = workbook.add_format({'bold': True, 'font_size': 14, 'bottom': 1})
            subheader_format = workbook.add_format({'bold': True, 'font_size': 11})
            
            for group_title, tables in results_dict.items():
                try:
                    worksheet = writer.sheets.get(sheet_name)
                    if worksheet is None: 
                        worksheet = workbook.add_worksheet(sheet_name)
                        writer.sheets[sheet_name] = worksheet
                    
                    worksheet.write(start_row, 0, group_title, header_format)
                    start_row += 2
                    
                    for table_title, df in tables.items():
                        try:
                            if df is None or df.empty: 
                                continue
                                
                            worksheet.write(start_row, 0, table_title, subheader_format)
                            start_row += 1
                            
                            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=True)
                            header_rows = df.columns.nlevels if isinstance(df.columns, pd.MultiIndex) else 1
                            start_row += len(df.index) + header_rows + 2
                            
                        except Exception as e:
                            st.warning(f"⚠️ 테이블 '{table_title}' 처리 중 오류: {e}")
                            continue
                            
                except Exception as e:
                    st.warning(f"⚠️ 그룹 '{group_title}' 처리 중 오류: {e}")
                    continue
                    
        return output.getvalue()
        
    except Exception as e:
        st.error(f"❌ Excel 파일 생성 중 오류: {e}")
        # 빈 Excel 파일 반환
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pd.DataFrame({'오류': ['Excel 파일 생성 중 오류가 발생했습니다.']}).to_excel(writer, index=False)
        return output.getvalue()

def auto_calculate_status(df):
    try:
        body_parts_map = {'목': 'N', '어깨': 'SH', '팔/팔꿈치': 'A', '손/손목/손가락': 'H', '허리': 'W', '다리/발': 'L'}
        
        # 각 신체부위별 계산
        for part_kr, part_en in body_parts_map.items():
            duration_col = f'문2-2({part_kr}) 통증기간'
            intensity_col = f'문2-3({part_kr}) 아픈정도'
            frequency_col = f'문2-4({part_kr}) 1년동안 증상빈도'
            
            n2_col = f'{part_en}-n2'
            n3_col = f'{part_en}-n3'
            n4_col = f'{part_en}-n4'
            manage_col = f'{part_en}-관리대상자'
            complain_col = f'{part_en}-통증호소자'
            
            # 필수 컬럼 확인
            required_cols = [duration_col, intensity_col, frequency_col]
            if not all(col in df.columns for col in required_cols):
                continue
            
            # 숫자 변환
            duration = pd.to_numeric(df[duration_col], errors='coerce')
            intensity = pd.to_numeric(df[intensity_col], errors='coerce')
            frequency = pd.to_numeric(df[frequency_col], errors='coerce')
            
            # 조건 설정
            cond_n2 = (duration >= 3)
            cond_n3 = (frequency >= 3)
            cond_n4 = (intensity >= 2)
            
            # 결과 설정
            df[n2_col] = np.where(cond_n2, '유병', '')
            df[n3_col] = np.where(cond_n3, '유병', '')
            df[n4_col] = np.where(cond_n4, '유병', '')
            
            # 관리대상자 및 통증호소자 판정
            is_management_target = cond_n2 | cond_n3 | cond_n4
            is_complainant = cond_n2 & cond_n3 & cond_n4
            
            df[manage_col] = np.where(is_management_target, 'Y', 'N')
            df[complain_col] = np.where(is_complainant, 'Y', 'N')
        
        # 전체 관리대상자 및 통증호소자 판정
        manage_cols = [f'{en}-관리대상자' for en in body_parts_map.values()]
        complain_cols = [f'{en}-통증호소자' for en in body_parts_map.values()]
        
        # 컬럼 존재 확인
        existing_manage_cols = [col for col in manage_cols if col in df.columns]
        existing_complain_cols = [col for col in complain_cols if col in df.columns]
        
        if existing_manage_cols and existing_complain_cols:
            is_complainant_any = (df[existing_complain_cols] == 'Y').any(axis=1)
            is_management_any = (df[existing_manage_cols] == 'Y').any(axis=1)
            
            df['최종상태'] = np.select(
                [is_complainant_any, ~is_complainant_any & is_management_any],
                ['통증호소자', '관리대상자'],
                default='정상'
            )
            
            df['관리대상자'] = np.where(df['최종상태'] == '관리대상자', 'Y', 'N')
            df['통증호소자'] = np.where(df['최종상태'] == '통증호소자', 'Y', 'N')
        else:
            # 기본값 설정
            df['최종상태'] = '정상'
            df['관리대상자'] = 'N'
            df['통증호소자'] = 'N'
        
        return df
        
    except Exception as e:
        st.error(f"❌ auto_calculate_status 오류: {e}")
        # 오류 발생 시 기본값으로 설정
        df['최종상태'] = '정상'
        df['관리대상자'] = 'N'
        df['통증호소자'] = 'N'
        return df

def create_table1(df, department3_name):
    try:
        df_filtered = df[df['작업부서3'] == department3_name].copy()
        if df_filtered.empty: 
            return None
        
        # 필수 컬럼 확인
        required_cols = ['작업부서4', '연령', '성별']
        missing_cols = [col for col in required_cols if col not in df_filtered.columns]
        if missing_cols:
            st.warning(f"⚠️ create_table1: 필수 컬럼 누락 - {missing_cols}")
            return None
        
        # 데이터 전처리
        df_filtered['성별'] = df_filtered['성별'].astype(str).str.strip()
        df_filtered['연령'] = pd.to_numeric(df_filtered['연령'], errors='coerce')
        
        # 성별 데이터 정규화
        df_filtered['성별'] = df_filtered['성별'].replace({
            '남': '남', '남성': '남', 'M': '남', 'm': '남', '1': '남',
            '여': '여', '여성': '여', 'F': '여', 'f': '여', '2': '여'
        })
        
        # 통계 계산
        age_stats = df_filtered.groupby('작업부서4')['연령'].agg(['count', 'mean', 'std']).rename(
            columns={'count': '응답자(명)', 'mean': '평균(세)', 'std': '표준편차'}
        )
        gender_counts = df_filtered.groupby(['작업부서4', '성별']).size().unstack(fill_value=0)
        
        result_table = pd.concat([age_stats, gender_counts], axis=1)
        
        # 컬럼 존재 확인 및 기본값 설정
        if '남' not in result_table.columns: 
            result_table['남'] = 0
        if '여' not in result_table.columns: 
            result_table['여'] = 0
            
        result_table['합계'] = result_table['남'] + result_table['여']
        result_table = result_table.rename(columns={'남': '남자(명)', '여': '여자(명)'})
        
        final_columns = ['응답자(명)', '평균(세)', '표준편차', '남자(명)', '여자(명)', '합계']
        result_table = result_table.reindex(columns=final_columns, fill_value=0).round(2)
        
        # 인덱스 이름을 '단위작업'으로 변경 (reindex 후에 다시 설정)
        result_table.index.name = '단위작업'
        
        return result_table
        

        
    except Exception as e:
        st.error(f"❌ create_table1 오류: {e}")
        return None

def create_table2(df, department3_name):
    try:
        df_filtered = df[df['작업부서3'] == department3_name].copy()
        if df_filtered.empty: 
            return None
        
        # 필수 컬럼 확인
        required_cols = ['작업부서4', '현 직장경력', '작업기간']
        missing_cols = [col for col in required_cols if col not in df_filtered.columns]
        if missing_cols:
            st.warning(f"⚠️ create_table2: 필수 컬럼 누락 - {missing_cols}")
            return None
        
        # 숫자 데이터 추출 및 변환
        df_filtered['현 직장경력_숫자'] = df_filtered['현 직장경력'].astype(str).str.extract(r'(\d+\.?\d*)').astype(float)
        df_filtered['작업기간_숫자'] = df_filtered['작업기간'].astype(str).str.extract(r'(\d+\.?\d*)').astype(float)
        
        # 범위 설정
        bins_current, labels_current = [-np.inf, 1, 3, 5, np.inf], ['<1년', '<3년', '<5년', '≥5년']
        bins_previous, labels_previous = [-np.inf, 1, 2, 3, np.inf], ['<1년', '<2년', '<3년', '≥3년']
        
        # 범주화
        df_filtered['현재작업기간_범위'] = pd.cut(df_filtered['현 직장경력_숫자'], bins=bins_current, labels=labels_current, right=False)
        df_filtered['이전작업기간_범위'] = pd.cut(df_filtered['작업기간_숫자'], bins=bins_previous, labels=labels_previous, right=False)
        
        # 통계 계산
        current_counts = df_filtered.groupby(['작업부서4', '현재작업기간_범위']).size().unstack(fill_value=0)
        current_counts['무응답'] = df_filtered.groupby('작업부서4').size() - current_counts.sum(axis=1)
        current_counts['합계'] = df_filtered.groupby('작업부서4').size()
        
        previous_counts = df_filtered.groupby(['작업부서4', '이전작업기간_범위']).size().unstack(fill_value=0)
        previous_counts['무응답'] = df_filtered.groupby('작업부서4').size() - previous_counts.sum(axis=1)
        previous_counts['합계'] = df_filtered.groupby('작업부서4').size()
        
        result_table = pd.concat([current_counts, previous_counts], axis=1, keys=['현재 작업기간', '이전 작업기간']).fillna(0).astype(int)
        
        # 인덱스 이름을 '단위작업'으로 변경
        result_table.index.name = '단위작업'
        
        return result_table
        
    except Exception as e:
        st.error(f"❌ create_table2 오류: {e}")
        return None

def create_table3(df, department3_name):
    try:
        df_filtered = df[df['작업부서3'] == department3_name].copy()
        if df_filtered.empty: 
            return None
        
        burden_column_name = '문1-5 일의 육체적 부담'
        
        # 필수 컬럼 확인
        if burden_column_name not in df_filtered.columns:
            st.warning(f"⚠️ create_table3: 필수 컬럼 누락 - {burden_column_name}")
            return None
        
        burden_map = {1: "전혀 힘들지 않음", 2: "견딜만 함", 3: "약간 힘듦", 4: "힘듦", 5: "매우 힘듦"}
        df_filtered['부담정도'] = df_filtered[burden_column_name].map(burden_map)
        
        # 통계 계산
        result_table = df_filtered.groupby(['작업부서4', '부담정도']).size().unstack(fill_value=0)
        result_table['합계'] = result_table.sum(axis=1)
        
        final_columns = ["전혀 힘들지 않음", "견딜만 함", "약간 힘듦", "힘듦", "매우 힘듦", "합계"]
        result_table = result_table.reindex(columns=final_columns, fill_value=0).fillna(0).astype(int)
        
        # 인덱스 이름을 '단위작업'으로 변경 (reindex 후에 다시 설정)
        result_table.index.name = '단위작업'
        
        return result_table
        
    except Exception as e:
        st.error(f"❌ create_table3 오류: {e}")
        return None

def create_table4(df, department3_name):
    try:
        df_filtered = df[df['작업부서3'] == department3_name].copy()
        if df_filtered.empty: 
            return None
        
        part_map = {
            '목': ('N-관리대상자', 'N-통증호소자'), 
            '어깨': ('SH-관리대상자', 'SH-통증호소자'), 
            '팔/팔꿈치': ('A-관리대상자', 'A-통증호소자'), 
            '손/손목/손가락': ('H-관리대상자', 'H-통증호소자'), 
            '허리': ('W-관리대상자', 'W-통증호소자'), 
            '다리/발': ('L-관리대상자', 'L-통증호소자')
        }
        
        # 필수 컬럼 확인
        required_cols = ['작업부서4', '관리대상자', '통증호소자'] + [col for cols in part_map.values() for col in cols]
        missing_cols = [col for col in required_cols if col not in df_filtered.columns]
        if missing_cols:
            st.warning(f"⚠️ create_table4: 필수 컬럼 누락 - {missing_cols}")
            return None
        
        total_col_map = {'관리대상자': '관리대상자', '통증호소자': '통증호소자'}
        body_parts, final_rows = list(part_map.keys()), []
        
        for dept4_name, dept4_df in df_filtered.groupby('작업부서4'):
            total_people = len(dept4_df)
            normal_row = {'단위작업': dept4_name, '상태': '정상'}
            manage_row = {'단위작업': dept4_name, '상태': '관리대상자'}
            complain_row = {'단위작업': dept4_name, '상태': '통증호소자'}
            
            for part_name, (manage_col, complain_col) in part_map.items():
                manage_count = (dept4_df[manage_col] == 'Y').sum()
                complain_count = (dept4_df[complain_col] == 'Y').sum()
                manage_row[part_name] = manage_count
                complain_row[part_name] = complain_count
                normal_row[part_name] = total_people - manage_count - complain_count
            
            manage_row['합계'] = (dept4_df[total_col_map['관리대상자']] == 'Y').sum()
            complain_row['합계'] = (dept4_df[total_col_map['통증호소자']] == 'Y').sum()
            normal_row['합계'] = total_people - manage_row['합계'] - complain_row['합계']
            
            final_rows.extend([normal_row, manage_row, complain_row])
        
        result_table = pd.DataFrame(final_rows).set_index(['단위작업', '상태'])
        return result_table[body_parts + ['합계']].fillna(0).astype(int)
        
    except Exception as e:
        st.error(f"❌ create_table4 오류: {e}")
        return None

def create_table5(df, department3_name):
    try:
        df_filtered = df[df['작업부서3'] == department3_name].copy()
        if df_filtered.empty: 
            return None
        
        feature_cols = {
            '개인취미생활': '문1-1 규칙적 취미활동', 
            '가사노동': '문1-2 평균 가사노동시간', 
            '개인병력_질병': '문1-3(1) 질병진단', 
            '개인병력_상해': '문1-4(1) 운동, 사고로 인한 상해'
        }
        
        # 필수 컬럼 확인
        required_cols = ['작업부서4', '최종상태'] + list(feature_cols.values())
        missing_cols = [col for col in required_cols if col not in df_filtered.columns]
        if missing_cols:
            st.warning(f"⚠️ create_table5: 필수 컬럼 누락 - {missing_cols}")
            return None
        
        # 개인병력 처리
        cond_illness = df_filtered[feature_cols['개인병력_질병']].isin([1, '1', 'Y', 'y', '예']).fillna(False)
        cond_injury = df_filtered[feature_cols['개인병력_상해']].isin([1, '1', 'Y', 'y', '예']).fillna(False)
        df_filtered['개인병력_결과'] = np.where(cond_illness | cond_injury, '예', '아니오')
        
        # 매핑 정의
        hobby_map = {1: "컴퓨터관련 활동", 2: "악기연주", 3: "뜨게질,붓글씨", 4: "스포츠활동", 5: "해당없음"}
        housework_map = {1: "거의안함", 2: "1시간미만", 3: "1-2시간미만", 4: "2-3시간미만", 5: "3시간이상"}
        
        # 데이터 매핑
        df_filtered['취미_분류'] = df_filtered[feature_cols['개인취미생활']].map(hobby_map)
        df_filtered['가사노동_분류'] = df_filtered[feature_cols['가사노동']].map(housework_map)
        
        # 통계 계산
        hobby_counts = df_filtered.groupby(['작업부서4', '최종상태', '취미_분류']).size().unstack(fill_value=0)
        housework_counts = df_filtered.groupby(['작업부서4', '최종상태', '가사노동_분류']).size().unstack(fill_value=0)
        history_counts = df_filtered.groupby(['작업부서4', '최종상태', '개인병력_결과']).size().unstack(fill_value=0)
        
        # 인덱스 이름을 '단위작업'으로 변경
        hobby_counts.index.names = ['단위작업', '상태']
        housework_counts.index.names = ['단위작업', '상태']
        history_counts.index.names = ['단위작업', '상태']
        
        # MultiIndex 컬럼 설정
        hobby_counts.columns = pd.MultiIndex.from_product([['개인취미'], hobby_counts.columns])
        housework_counts.columns = pd.MultiIndex.from_product([['가사노동'], housework_counts.columns])
        history_counts.columns = pd.MultiIndex.from_product([['개인병력'], history_counts.columns])
        
        result_table = pd.concat([hobby_counts, housework_counts, history_counts], axis=1)
        result_table.index.names = ['단위작업', '상태']
        
        # 전체 합계 계산
        total_hobby = df_filtered.groupby(['최종상태', '취미_분류']).size().unstack(fill_value=0)
        total_housework = df_filtered.groupby(['최종상태', '가사노동_분류']).size().unstack(fill_value=0)
        total_history = df_filtered.groupby(['최종상태', '개인병력_결과']).size().unstack(fill_value=0)
        
        total_hobby.columns = pd.MultiIndex.from_product([['개인취미'], total_hobby.columns])
        total_housework.columns = pd.MultiIndex.from_product([['가사노동'], total_housework.columns])
        total_history.columns = pd.MultiIndex.from_product([['개인병력'], total_history.columns])
        
        total_table = pd.concat([total_hobby, total_housework, total_history], axis=1)
        total_table.index = pd.MultiIndex.from_product([['전체 합계'], total_table.index], names=['단위작업', '상태'])
        
        final_table = pd.concat([result_table, total_table])
        
        # 최종 컬럼 정의
        final_columns = pd.MultiIndex.from_tuples(
            [('개인취미', label) for label in hobby_map.values()] + 
            [('가사노동', label) for label in housework_map.values()] + 
            [('개인병력', label) for label in ['예', '아니오']]
        )
        
        # 인덱스 정렬
        all_depts_4 = sorted(list(df_filtered['작업부서4'].unique())) + ['전체 합계']
        all_statuses = ['정상', '관리대상자', '통증호소자']
        full_index = pd.MultiIndex.from_product([all_depts_4, all_statuses], names=['단위작업', '상태'])
        
        return final_table.reindex(index=full_index, columns=final_columns, fill_value=0).fillna(0).astype(int)
        
    except Exception as e:
        st.error(f"❌ create_table5 오류: {e}")
        return None

# --------------------------------------------------------------------------
# 메인 로직
# --------------------------------------------------------------------------
if uploaded_file is not None:
    df = None
    all_results_for_excel = {}
    try:
        # 파일 로딩
        df = pd.read_excel(uploaded_file, header=2)
        df.columns = df.columns.str.strip().str.replace('\n', '', regex=False)
        st.success("✅ 파일 로딩 성공! 자동 계산을 시작합니다.")
        
        # 자동 계산
        df_calculated = auto_calculate_status(df.copy())
        st.success("✅ 자동 계산 완료! 통계 분석을 시작합니다.")
        st.write("---")
        
        # 필수 컬럼 확인
        required_columns = ['작업부서3', '작업부서4', '연령', '성별']
        missing_columns = [col for col in required_columns if col not in df_calculated.columns]
        if missing_columns:
            st.error(f"❌ 필수 컬럼이 누락되었습니다: {', '.join(missing_columns)}")
            st.stop()
        
        # 부서별 분석
        department3_list = df_calculated['작업부서3'].dropna().unique()
        if len(department3_list) == 0:
            st.error("❌ '작업부서3'에 데이터가 없습니다. 데이터를 확인해주세요.")
            st.stop()
            
        for dept3 in department3_list:
            try:
                group_title = f"분석 결과: {dept3}"
                st.header(group_title)
                all_results_for_excel[group_title] = {}
                
                # 각 테이블 생성 및 에러 처리
                try:
                    table1 = create_table1(df_calculated, dept3)
                    if table1 is not None:
                        st.subheader("1. 기초현황")
                        # 인덱스 이름을 '단위작업'으로 설정하여 표시
                        display_table = table1.reset_index()
                        display_table = display_table.rename(columns={'index': '단위작업'})
                        display_table = display_table.set_index('단위작업')
                        st.dataframe(display_table, use_container_width=True)
                        all_results_for_excel[group_title]["1. 기초현황"] = table1
                    else:
                        st.warning("⚠️ 1. 기초현황: 해당 부서의 데이터가 부족합니다.")
                except Exception as e:
                    st.error(f"❌ 1. 기초현황 생성 중 오류: {e}")
                
                try:
                    table2 = create_table2(df_calculated, dept3)
                    if table2 is not None:
                        st.subheader("2. 작업기간")
                        # 인덱스 이름을 '단위작업'으로 설정하여 표시
                        display_table = table2.reset_index()
                        display_table = display_table.rename(columns={'index': '단위작업'})
                        display_table = display_table.set_index('단위작업')
                        st.dataframe(display_table, use_container_width=True)
                        all_results_for_excel[group_title]["2. 작업기간"] = table2
                    else:
                        st.warning("⚠️ 2. 작업기간: 해당 부서의 데이터가 부족합니다.")
                except Exception as e:
                    st.error(f"❌ 2. 작업기간 생성 중 오류: {e}")
                
                try:
                    table3 = create_table3(df_calculated, dept3)
                    if table3 is not None:
                        st.subheader("3. 육체적 부담정도")
                        # 인덱스 이름을 '단위작업'으로 설정하여 표시
                        display_table = table3.reset_index()
                        display_table = display_table.rename(columns={'index': '단위작업'})
                        display_table = display_table.set_index('단위작업')
                        st.dataframe(display_table, use_container_width=True)
                        all_results_for_excel[group_title]["3. 부담정도"] = table3
                    else:
                        st.warning("⚠️ 3. 육체적 부담정도: 해당 부서의 데이터가 부족합니다.")
                except Exception as e:
                    st.error(f"❌ 3. 육체적 부담정도 생성 중 오류: {e}")
                
                try:
                    table4 = create_table4(df_calculated, dept3)
                    if table4 is not None:
                        st.subheader("4. 신체부위별 통증 호소 현황")
                        # 인덱스 이름을 '단위작업'으로 설정하여 표시
                        display_table = table4.reset_index()
                        display_table = display_table.rename(columns={'level_0': '단위작업', 'level_1': '상태'})
                        display_table = display_table.set_index(['단위작업', '상태'])
                        st.dataframe(display_table, use_container_width=True)
                        all_results_for_excel[group_title]["4. 통증현황"] = table4
                    else:
                        st.warning("⚠️ 4. 신체부위별 통증 호소 현황: 해당 부서의 데이터가 부족합니다.")
                except Exception as e:
                    st.error(f"❌ 4. 신체부위별 통증 호소 현황 생성 중 오류: {e}")
                
                try:
                    table5 = create_table5(df_calculated, dept3)
                    if table5 is not None:
                        st.subheader("5. 개인 특성")
                        # 인덱스 이름을 '단위작업'으로 설정하여 표시
                        display_table = table5.reset_index()
                        display_table = display_table.rename(columns={'level_0': '단위작업', 'level_1': '상태'})
                        display_table = display_table.set_index(['단위작업', '상태'])
                        st.dataframe(display_table, use_container_width=True)
                        all_results_for_excel[group_title]["5. 개인특성"] = table5
                    else:
                        st.warning("⚠️ 5. 개인 특성: 해당 부서의 데이터가 부족합니다.")
                except Exception as e:
                    st.error(f"❌ 5. 개인 특성 생성 중 오류: {e}")
                
                st.write("---")
                
            except Exception as e:
                st.error(f"❌ {dept3} 부서 분석 중 오류가 발생했습니다: {e}")
                continue

        # 결과 다운로드
        if all_results_for_excel:
            st.header("⬇️ 전체 결과 다운로드")
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("1. 자동계산 원본 데이터")
                try:
                    excel_bytes = to_excel_bytes(df_calculated)
                    st.download_button(
                        label="자동계산 원본 다운로드 (Excel)", 
                        data=excel_bytes, 
                        file_name="결과_자동계산_원본.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ 원본 데이터 다운로드 생성 중 오류: {e}")
                    
            with col2:
                st.subheader("2. 최종 통계표 (통합)")
                try:
                    results_bytes = results_to_excel_bytes(all_results_for_excel)
                    st.download_button(
                        label="최종 통계표 다운로드 (Excel)", 
                        data=results_bytes, 
                        file_name="결과_최종_통계표.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"❌ 통합 통계표 다운로드 생성 중 오류: {e}")
        else:
            st.warning("⚠️ 생성된 결과가 없어 다운로드할 수 없습니다.")
            
    except Exception as e:
        st.error(f"❌ 파일을 처리하는 중 오류가 발생했습니다: {e}")
        st.error("💡 문제 해결 방법:")
        st.error("1. 엑셀 파일이 올바른 형식인지 확인해주세요.")
        st.error("2. 템플릿을 사용하여 데이터를 입력했는지 확인해주세요.")
        st.error("3. 필수 컬럼들이 모두 포함되어 있는지 확인해주세요.")
