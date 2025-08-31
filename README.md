# 🏥 근골증상 및 직무스트레스 통계 시스템

근골격계 증상 분석과 직무스트레스 분석을 통합한 웹 기반 분석 시스템입니다.

## 🚀 주요 기능

### 1. 근골격계 증상 분석
- 신체 부위별 통증 분석 (목, 어깨, 팔/팔꿈치, 손/손목/손가락, 허리, 다리/발)
- 관리대상자 및 통증호소자 자동 분류
- 부서별, 성별, 연령별 통계 분석
- 육체적 부담정도 및 개인특성 분석

### 2. 직무스트레스 분석
- 8개 영역별 스트레스 점수 자동 계산
  - 물리환경, 직무요구, 직무자율, 관계갈등
  - 직업불안정, 조직체계, 보상부적절, 직장문화
- 성별/공정별 통계 분석
- 기준치 초과자 자동 식별 및 명단 생성

## 📋 시작하기

### 사전 요구사항
- Python 3.8 이상
- pip (Python 패키지 관리자)

### 설치 방법

1. 저장소 클론
```bash
git clone https://github.com/[your-username]/stress-analysis-system.git
cd stress-analysis-system
```

2. 가상환경 생성 (권장)
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

3. 필요한 패키지 설치
```bash
pip install -r requirements.txt
```

4. 애플리케이션 실행
```bash
streamlit run app.py
```

5. 웹 브라우저에서 자동으로 열립니다 (기본: http://localhost:8501)

## 📁 프로젝트 구조

```
stress-analysis-system/
│
├── app.py                              # 메인 애플리케이션
├── pages/                              # 멀티페이지 구조
│   ├── 1_📊_근골격계_증상_분석.py
│   ├── 2_🧠_직무스트레스_분석.py
│   └── 3_📈_직무스트레스_결과.py
├── requirements.txt                    # 필요한 패키지 목록
├── README.md                          # 프로젝트 문서
├── LICENSE                            # 라이선스 정보
└── .gitignore                         # Git 제외 파일

```

## 💻 사용 방법

### 근골격계 증상 분석
1. 좌측 메뉴에서 "📊 근골격계 증상 분석" 선택
2. 템플릿 다운로드 후 데이터 입력
3. 완성된 엑셀 파일 업로드
4. 자동 분석 및 결과 확인
5. 결과 엑셀 파일 다운로드

### 직무스트레스 분석
1. 좌측 메뉴에서 "🧠 직무스트레스 분석" 선택
2. 템플릿 다운로드 후 설문 데이터 입력 (문1-1 ~ 문1-43)
3. 완성된 엑셀 파일 업로드
4. "분석 시작" 버튼 클릭
5. "📈 직무스트레스 결과" 페이지에서 결과 확인 및 다운로드

## 📊 데이터 형식

### 엑셀 파일 구조
- 3행부터 데이터 시작 (1행: 제목, 2행: 빈 행, 3행: 열 제목)
- UTF-8 인코딩 권장
- .xlsx 또는 .xls 형식 지원

### 필수 열 정보
- 공통: 대상, 성명, 연령, 성별, 작업부서1, 작업부서2, 작업부서3
- 근골격계: 각 신체부위별 통증 관련 문항
- 직무스트레스: 문1-1 ~ 문1-43 설문 응답

## 🔧 문제 해결

### 일반적인 문제
1. **파일 업로드 오류**: 템플릿 형식 확인, 3행에 열 제목이 있는지 확인
2. **계산 오류**: 필수 열이 모두 있는지 확인
3. **한글 깨짐**: UTF-8 인코딩으로 저장

### 기술 지원
- 이슈 등록: [GitHub Issues](https://github.com/[your-username]/stress-analysis-system/issues)
- 이메일: kangyoon.kim@ihealse.com

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 🤝 기여하기

프로젝트 개선에 기여하고 싶으시다면:
1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 연락처

프로젝트 관리자 - kangyoon.kim@ihealse.com


프로젝트 링크: [https://github.com/[kykim0310]/stress-analysis-system](https://github.com/[your-username]/stress-analysis-system)
