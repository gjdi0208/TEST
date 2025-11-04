# 📊 자동 판매 데이터 분석 및 보고서 시스템

CSV 데이터를 읽어서 자동으로 분석하고, 시각화 차트가 포함된 워드 보고서를 생성한 후 이메일로 전송하는 완전 자동화 시스템입니다.

## 🚀 주요 기능

- **📊 데이터 분석**: CSV 파일의 판매 데이터를 자동으로 분석
- **📈 시각화**: 카테고리별, 지역별, 일별 매출 차트 생성
- **📄 보고서 생성**: 분석 결과를 포함한 전문적인 워드 문서 생성
- **📧 자동 전송**: 생성된 보고서를 이메일로 자동 전송
- **🔧 데이터 정리**: 자동 데이터 검증 및 정리

## 📁 파일 구조

```
📦 프로젝트 폴더
├── 📄 cicd_data.csv                    # 원본 데이터 파일
├── 🐍 automated_sales_report.py        # 메인 통합 시스템
├── 🐍 word_report_generator.py         # 워드 보고서 생성기
├── 🐍 email_sender.py                  # 독립 이메일 전송기
├── 🐍 simple_sales_report.py           # 간단한 분석기
├── 📚 EMAIL_SETUP_GUIDE.md             # 이메일 설정 가이드
└── 📚 README.md                        # 이 파일
```

## 🛠️ 설치 및 설정

### 1. 필요한 라이브러리 설치

```bash
pip install pandas matplotlib seaborn numpy openpyxl python-docx
```

### 2. 이메일 설정

`automated_sales_report.py` 파일에서 이메일 설정을 수정하세요:

```python
EMAIL_CONFIG = {
    'sender_email': 'your_email@gmail.com',      # 보내는 사람
    'sender_password': 'your_app_password',      # 앱 비밀번호
    'recipient_emails': [                        # 받는 사람들
        'manager@company.com',
        'team@company.com'
    ],
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587
}
```

> 📋 자세한 이메일 설정 방법은 `EMAIL_SETUP_GUIDE.md`를 참조하세요.

## 🎯 사용 방법

### 🔄 전체 자동화 시스템 실행

```bash
python automated_sales_report.py
```

**실행 과정:**
1. 📊 CSV 데이터 로드 및 정리
2. 📈 시각화 차트 3개 생성
3. 📄 워드 보고서 생성 (차트 포함)
4. 📧 이메일 자동 전송

### 📄 워드 보고서만 생성

```bash
python word_report_generator.py
```

### 📧 기존 보고서 이메일 전송

```bash
python email_sender.py
```

## 📊 생성되는 보고서 내용

### 📈 시각화 차트
1. **카테고리별 매출 비율** (파이차트)
2. **지역별 매출 비교** (막대차트)
3. **일별 매출 추이** (선 그래프)

### 📋 분석 섹션
1. **📊 요약 통계**: 총 매출액, 수량, 평균 주문 금액 등
2. **📊 카테고리별 분석**: 카테고리별 매출, 수량, 제품 수
3. **🌍 지역별 분석**: 지역별 성과 비교
4. **🏆 베스트셀러 제품**: TOP 10 제품 순위
5. **📈 일별 매출 추이**: 최고/최저 매출일, 평균 매출
6. **💡 결론 및 제안사항**: 데이터 기반 인사이트

## 📁 생성되는 파일

### 자동 생성 파일들
- `sales_analysis_report_YYYYMMDD_HHMMSS.docx` - 메인 보고서
- `chart_category_pie.png` - 카테고리 차트
- `chart_region_bar.png` - 지역별 차트  
- `chart_daily_trend.png` - 일별 추이 차트

## ⚙️ 설정 옵션

### 이메일 서비스별 SMTP 설정

| 서비스 | SMTP 서버 | 포트 |
|--------|-----------|------|
| Gmail | smtp.gmail.com | 587 |
| Outlook | smtp-mail.outlook.com | 587 |
| Yahoo | smtp.mail.yahoo.com | 587 |
| Naver | smtp.naver.com | 587 |

### 차트 스타일 커스터마이징

`create_charts()` 함수에서 다음을 수정할 수 있습니다:
- 색상 팔레트
- 차트 크기
- 폰트 설정
- 레이아웃

## 🔧 데이터 요구사항

CSV 파일은 다음 컬럼을 포함해야 합니다:

| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| Date | 날짜 | 판매 날짜 (YYYY-MM-DD) |
| ProductID | 문자열 | 제품 ID |
| ProductName | 문자열 | 제품명 |
| Category | 문자열 | 제품 카테고리 |
| Quantity | 숫자 | 판매 수량 |
| UnitPrice | 숫자 | 단가 |
| TotalPrice | 숫자 | 총 금액 |
| Region | 문자열 | 판매 지역 |
| Salesperson | 문자열 | 영업사원 |

## ⚠️ 주의사항

1. **이메일 보안**: 앱 비밀번호 사용 권장
2. **파일 크기**: 첨부파일은 25MB 이하로 제한
3. **데이터 품질**: 결측값이나 오류 데이터는 자동으로 제거됨
4. **네트워크**: 이메일 전송 시 인터넷 연결 필요

## 🐛 문제 해결

### 일반적인 오류들

#### 1. 모듈 없음 오류
```bash
pip install pandas matplotlib python-docx
```

#### 2. 이메일 인증 실패
- Gmail 2단계 인증 활성화
- 앱 비밀번호 생성 및 사용

#### 3. 한글 폰트 오류
- Windows: 'Malgun Gothic' 폰트 확인
- 다른 OS: 적절한 한글 폰트로 변경

#### 4. 차트 생성 실패
```python
matplotlib.use('Agg')  # GUI 없는 환경에서 사용
```

## 📈 향후 개선 계획

- [ ] 웹 대시보드 인터페이스 추가
- [ ] 다양한 차트 타입 지원
- [ ] 스케줄링 기능 (정기 자동 실행)
- [ ] 데이터베이스 연동
- [ ] PDF 보고서 옵션
- [ ] 슬랙/팀즈 연동

## 📞 지원 및 문의

시스템 사용 중 문제가 발생하면:
1. 오류 메시지 확인
2. 로그 파일 검토
3. 설정 파일 점검
4. 네트워크 연결 확인

---

**🎉 이제 자동화된 보고서 시스템을 즐겨보세요!**