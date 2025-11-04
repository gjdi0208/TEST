# 📧 이메일 설정 가이드

이 문서는 자동 보고서 시스템에서 이메일 전송 기능을 사용하기 위한 설정 방법을 설명합니다.

## 🔧 이메일 설정 방법

### 1. Gmail 사용 시 (권장)

1. **2단계 인증 활성화**
   - Google 계정 설정 → 보안 → 2단계 인증 활성화

2. **앱 비밀번호 생성**
   - Google 계정 설정 → 보안 → 앱 비밀번호
   - "메일" 앱 선택 후 16자리 비밀번호 생성

3. **설정 파일 수정**
   ```python
   EMAIL_CONFIG = {
       'sender_email': 'your_actual_email@gmail.com',  # 실제 Gmail 주소
       'sender_password': 'abcd efgh ijkl mnop',       # 생성된 16자리 앱 비밀번호
       'recipient_emails': [                           # 받는 사람들
           'manager@company.com',
           'team@company.com'
       ],
       'smtp_server': 'smtp.gmail.com',
       'smtp_port': 587
   }
   ```

### 2. Outlook/Hotmail 사용 시

```python
EMAIL_CONFIG = {
    'sender_email': 'your_email@outlook.com',
    'sender_password': 'your_password',
    'recipient_emails': ['recipient@company.com'],
    'smtp_server': 'smtp-mail.outlook.com',
    'smtp_port': 587
}
```

### 3. 네이버 메일 사용 시

```python
EMAIL_CONFIG = {
    'sender_email': 'your_email@naver.com',
    'sender_password': 'your_password',
    'recipient_emails': ['recipient@company.com'],
    'smtp_server': 'smtp.naver.com',
    'smtp_port': 587
}
```

## 🚀 실제 이메일 전송 활성화

`automated_sales_report.py` 파일에서 다음 라인을 수정하세요:

**현재 (테스트 모드):**
```python
# email_success = send_email_with_report(report_file, EMAIL_CONFIG)
print("📧 이메일 전송 기능은 설정 완료 후 사용하세요.")
email_success = True  # 테스트용
```

**실제 전송 모드로 변경:**
```python
email_success = send_email_with_report(report_file, EMAIL_CONFIG)
# print("📧 이메일 전송 기능은 설정 완료 후 사용하세요.")
# email_success = True  # 테스트용
```

## 📋 사용 예제

### 예제 1: 간단한 설정
```python
EMAIL_CONFIG = {
    'sender_email': 'report@mycompany.com',
    'sender_password': 'myapppassword1234',
    'recipient_emails': ['boss@mycompany.com'],
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587
}
```

### 예제 2: 여러 수신자
```python
EMAIL_CONFIG = {
    'sender_email': 'analytics@company.com',
    'sender_password': 'app_password_here',
    'recipient_emails': [
        'ceo@company.com',
        'sales_manager@company.com',
        'marketing_team@company.com'
    ],
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587
}
```

## ⚠️ 보안 주의사항

1. **앱 비밀번호 사용**: 일반 계정 비밀번호가 아닌 앱 전용 비밀번호를 사용하세요.
2. **코드 보안**: 이메일 비밀번호를 코드에 직접 입력하지 말고 환경변수나 별도 설정 파일을 사용하세요.
3. **권한 관리**: 필요한 사람에게만 보고서를 전송하도록 수신자 목록을 관리하세요.

## 🔍 문제 해결

### 인증 오류 발생 시
- Gmail: 앱 비밀번호가 올바른지 확인
- 2단계 인증이 활성화되어 있는지 확인
- 계정이 잠기지 않았는지 확인

### 전송 실패 시
- 네트워크 연결 확인
- SMTP 서버 주소와 포트 번호 확인
- 첨부파일 크기 확인 (Gmail: 25MB 제한)

## 📞 지원

문제가 계속 발생하면 다음을 확인하세요:
1. 이메일 서비스 제공업체의 SMTP 설정 문서
2. 방화벽이나 보안 소프트웨어 설정
3. 회사 네트워크의 이메일 정책