# 🚀 GitHub Actions 자동 보고서 시스템

매주 월요일 오전 9시(한국시간)에 자동으로 판매 데이터 분석 보고서를 생성하는 GitHub Actions 워크플로우입니다.

## 📁 파일 구조

```
.github/
└── workflows/
    ├── weekly-report.yml       # 완전한 워크플로우
    └── simple-report.yml       # 간단한 워크플로우

requirements.txt                # 필요한 Python 패키지
github_actions_report.py        # GitHub Actions 전용 스크립트
word_report_generator.py        # 원본 스크립트
cicd_data.csv                  # 데이터 파일 (필수)
```

## ⚙️ 설정 방법

### 1. 리포지토리 설정

```bash
# 1. 리포지토리 생성 및 클론
git clone https://github.com/your-username/your-repo.git
cd your-repo

# 2. 워크플로우 디렉토리 생성
mkdir -p .github/workflows

# 3. 필요한 파일들 추가
# - 워크플로우 파일들 (.github/workflows/*.yml)
# - 스크립트 파일들 (*.py)
# - 데이터 파일 (cicd_data.csv)
# - requirements.txt
```

### 2. 데이터 파일 준비

`cicd_data.csv` 파일을 리포지토리 루트에 추가하세요:

```csv
Date,ProductID,ProductName,Category,Quantity,UnitPrice,TotalPrice,Region,Salesperson
2025-09-01,P1001,Alpha Widget,Electronics,12,50,600,North,John Doe
2025-09-01,P2003,Orion T-shirt,Apparel,60,18,1080,South,Leo Martin
...
```

### 3. GitHub 설정

#### 3.1 Actions 권한 설정
- Repository Settings → Actions → General
- "Allow all actions and reusable workflows" 선택
- "Read and write permissions" 활성화

#### 3.2 Pages 설정 (선택사항)
- Repository Settings → Pages
- Source: "GitHub Actions" 선택

## 🕐 스케줄 설정

### 현재 설정: 매주 월요일 오전 9시 (한국시간)

```yaml
schedule:
  - cron: '0 0 * * MON'  # UTC 기준 월요일 00:00 = 한국시간 09:00
```

### 다른 스케줄 예시

```yaml
# 매일 오전 9시
- cron: '0 0 * * *'

# 매주 금요일 오후 5시 (한국시간)
- cron: '0 8 * * FRI'

# 매월 1일 오전 10시 (한국시간)
- cron: '0 1 1 * *'

# 매일 오전 9시, 오후 6시 (한국시간)
- cron: '0 0,9 * * *'
```

## 🚀 워크플로우 실행

### 자동 실행
- 매주 월요일 오전 9시에 자동 실행
- GitHub Actions 탭에서 실행 상태 확인 가능

### 수동 실행
1. GitHub 리포지토리 → Actions 탭
2. "Weekly Sales Report Generation" 워크플로우 선택
3. "Run workflow" 버튼 클릭

### 실행 과정
1. **환경 설정**: Python 3.11, 필요한 패키지 설치
2. **데이터 검증**: CSV 파일 존재 여부 확인
3. **보고서 생성**: 워드 문서 + 차트 생성
4. **결과 저장**: Artifacts 업로드 및 Release 생성

## 📊 생성되는 결과물

### 1. Artifacts
- **파일명**: `weekly-sales-report-{실행번호}`
- **포함 내용**: 
  - `sales_analysis_report_YYYYMMDD_HHMMSS.docx`
  - `chart_category_pie.png`
  - `chart_region_bar.png`
  - `chart_daily_trend.png`
- **보관 기간**: 30일

### 2. GitHub Releases
- **태그**: `report-{실행번호}-{실행ID}`
- **제목**: `Weekly Sales Report - {실행번호}`
- **첨부파일**: 워드 보고서 + 차트들

## 🔧 환경별 차이점

| 구분 | 로컬 환경 | GitHub Actions |
|------|-----------|----------------|
| 운영체제 | Windows | Ubuntu Linux |
| 한글 폰트 | Malgun Gothic | DejaVu Sans |
| 이메일 전송 | 지원 | 비활성화 |
| 파일 저장 | 로컬 폴더 | Artifacts + Releases |

## 🐛 문제 해결

### 일반적인 오류들

#### 1. 데이터 파일 없음
```
❌ cicd_data.csv 파일이 없습니다.
```
**해결**: 리포지토리 루트에 `cicd_data.csv` 파일 추가

#### 2. 패키지 설치 실패
```
❌ pip install 실패
```
**해결**: `requirements.txt` 파일 확인 및 버전 호환성 검토

#### 3. 권한 오류
```
❌ GITHUB_TOKEN 권한 부족
```
**해결**: Repository Settings → Actions → General에서 권한 설정

#### 4. 한글 폰트 오류
```
❌ Font 'Malgun Gothic' not found
```
**해결**: 자동으로 DejaVu Sans로 변경됨 (정상 동작)

### 로그 확인 방법
1. GitHub Actions 탭 → 실행된 워크플로우 클릭
2. 각 단계별 로그 확인
3. 오류 메시지를 통해 문제점 파악

## 📈 모니터링 및 알림

### 성공/실패 알림
현재는 기본 GitHub 알림만 지원합니다. 추가 알림이 필요한 경우:

1. **이메일 알림**: GitHub 계정 설정에서 활성화
2. **Slack 알림**: 워크플로우에 Slack action 추가
3. **Discord 알림**: Discord webhook 연동

### 실행 통계
- Actions 탭에서 실행 히스토리 확인
- 성공률, 평균 실행 시간 등 모니터링 가능

## 🔒 보안 고려사항

### 1. 데이터 보안
- 민감한 데이터는 private repository 사용
- 필요시 데이터 암호화 적용

### 2. 토큰 관리
- GITHUB_TOKEN은 자동으로 제공됨
- 추가 시크릿이 필요한 경우 Repository Secrets 사용

### 3. 권한 제한
- 워크플로우는 최소 권한으로 실행
- 불필요한 권한 제거

## 🛠️ 커스터마이징

### 보고서 내용 변경
`github_actions_report.py` 파일을 수정하여:
- 차트 스타일 변경
- 분석 항목 추가/제거
- 보고서 형식 수정

### 스케줄 변경
워크플로우 파일의 `cron` 표현식 수정:
```yaml
schedule:
  - cron: '0 2 * * TUE'  # 매주 화요일 오전 11시 (한국시간)
```

### 알림 추가
워크플로우에 알림 단계 추가:
```yaml
- name: Send Slack notification
  uses: 8398a7/action-slack@v3
  with:
    status: ${{ job.status }}
    webhook_url: ${{ secrets.SLACK_WEBHOOK }}
```

## 📚 참고 자료

- [GitHub Actions 문서](https://docs.github.com/en/actions)
- [Cron 표현식 생성기](https://crontab.guru/)
- [GitHub Marketplace](https://github.com/marketplace?type=actions)

---

**🎉 이제 자동화된 보고서 시스템을 즐겨보세요!**