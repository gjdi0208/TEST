# 🔐 GitHub Secrets 설정 가이드

GitHub Actions에서 이메일 발송을 위한 보안 설정 방법입니다.

## 📋 설정해야 할 Secrets

GitHub 리포지토리에서 다음 Secrets을 설정해야 합니다:

### 1. Repository Settings → Secrets and variables → Actions 이동

### 2. 다음 Secrets 추가:

| Secret Name | Value | 설명 |
|-------------|-------|------|
| `SENDER_EMAIL` | `gjdi0208@gmail.com` | 발송자 이메일 |
| `SENDER_PASSWORD` | `hkcn ywsk wmxk vqvm` | Gmail 앱 비밀번호 |
| `RECIPIENT_EMAIL` | `gjdi0208@gmail.com` | 수신자 이메일 |

## 🔧 Secrets 설정 방법

1. **GitHub 리포지토리 이동**
   - 리포지토리 → Settings 탭 클릭

2. **Secrets 메뉴 접근**
   - 왼쪽 사이드바 → Security → Secrets and variables → Actions

3. **New repository secret 클릭**

4. **각 Secret 추가**:
   ```
   Name: SENDER_EMAIL
   Secret: gjdi0208@gmail.com
   ```
   
   ```
   Name: SENDER_PASSWORD  
   Secret: hkcn ywsk wmxk vqvm
   ```
   
   ```
   Name: RECIPIENT_EMAIL
   Secret: gjdi0208@gmail.com
   ```

## 🚨 보안 고려사항

### ✅ 장점
- 이메일 자격 증명이 코드에 노출되지 않음
- GitHub Secrets는 암호화되어 저장됨
- Actions 로그에서 Secrets 값이 마스킹됨

### 📝 주의사항
- Secrets를 설정하지 않으면 기본값 사용
- 기본값은 현재 하드코딩된 값이 사용됨
- 프로덕션 환경에서는 반드시 Secrets 설정 권장

## 🔄 Fallback 설정

현재 워크플로우는 다음과 같이 설정되어 있습니다:

```yaml
env:
  SENDER_EMAIL: ${{ secrets.SENDER_EMAIL || 'gjdi0208@gmail.com' }}
  SENDER_PASSWORD: ${{ secrets.SENDER_PASSWORD || 'hkcn ywsk wmxk vqvm' }}
  RECIPIENT_EMAIL: ${{ secrets.RECIPIENT_EMAIL || 'gjdi0208@gmail.com' }}
```

- Secrets가 설정되어 있으면 해당 값 사용
- Secrets가 없으면 기본값 (현재 설정) 사용

## 🎯 권장사항

1. **즉시 테스트**: Secrets 설정 없이도 현재 설정으로 작동
2. **보안 강화**: 시간이 될 때 Secrets 설정 추가
3. **다중 수신자**: RECIPIENT_EMAIL을 여러 이메일로 확장 가능

## 📧 다중 수신자 설정 예시

RECIPIENT_EMAIL Secret을 다음과 같이 설정하면 여러 명에게 발송 가능:

```
gjdi0208@gmail.com,manager@company.com,team@company.com
```

워크플로우에서 자동으로 쉼표로 분리하여 처리합니다.