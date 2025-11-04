import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import getpass

def send_report_email(docx_file_path, recipient_emails, sender_email=None, sender_password=None, 
                     smtp_server="smtp.gmail.com", smtp_port=587, custom_subject=None, custom_body=None):
    """
    워드 보고서를 첨부하여 이메일 전송
    
    Args:
        docx_file_path (str): 전송할 워드 파일 경로
        recipient_emails (list): 받는 사람 이메일 주소 리스트
        sender_email (str): 보내는 사람 이메일 주소
        sender_password (str): 보내는 사람 이메일 비밀번호
        smtp_server (str): SMTP 서버 주소
        smtp_port (int): SMTP 포트 번호
        custom_subject (str): 사용자 정의 제목
        custom_body (str): 사용자 정의 본문
    
    Returns:
        bool: 전송 성공 여부
    """
    print("="*50)
    print("📧 이메일 전송 시스템")
    print("="*50)
    
    try:
        # 파일 존재 확인
        if not os.path.exists(docx_file_path):
            print(f"❌ 파일을 찾을 수 없습니다: {docx_file_path}")
            return False
        
        # 파일 크기 확인 (25MB 제한 - Gmail 기준)
        file_size_mb = os.path.getsize(docx_file_path) / (1024 * 1024)
        print(f"📎 첨부파일: {os.path.basename(docx_file_path)} ({file_size_mb:.2f}MB)")
        
        if file_size_mb > 25:
            print("⚠️  파일 크기가 25MB를 초과합니다. Gmail에서 전송이 제한될 수 있습니다.")
            continue_choice = input("계속 진행하시겠습니까? (y/n): ")
            if continue_choice.lower() != 'y':
                return False
        
        # 이메일 정보 입력받기
        if not sender_email:
            sender_email = input("📮 보내는 사람 이메일 주소: ")
        
        if not sender_password:
            sender_password = getpass.getpass("🔐 이메일 비밀번호 (앱 비밀번호): ")
        
        print(f"📧 받는 사람: {', '.join(recipient_emails)}")
        print(f"📡 SMTP 서버: {smtp_server}:{smtp_port}")
        
        # 이메일 메시지 생성
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ", ".join(recipient_emails)
        
        # 제목 설정
        if custom_subject:
            msg['Subject'] = custom_subject
        else:
            msg['Subject'] = f"판매 데이터 분석 보고서 - {datetime.now().strftime('%Y년 %m월 %d일')}"
        
        # 본문 설정
        if custom_body:
            body = custom_body
        else:
            body = f"""
안녕하세요,

첨부된 파일은 {datetime.now().strftime('%Y년 %m월 %d일')}에 생성된 판매 데이터 분석 보고서입니다.

📊 보고서 주요 내용:
┌─────────────────────────────────┐
│ • 매출 요약 통계                    │
│ • 카테고리별 판매 분석               │
│ • 지역별 성과 비교                  │
│ • 베스트셀러 제품 순위               │
│ • 영업사원별 실적                   │
│ • 일별 매출 추이                    │
│ • 시각화 차트 및 그래프              │
└─────────────────────────────────┘

📈 주요 인사이트와 분석 결과가 포함되어 있으니 검토 후 피드백 부탁드립니다.

궁금한 사항이 있으시면 언제든지 연락주세요.

감사합니다.

---
⚡ 자동화 시스템으로 생성된 보고서입니다.
🕒 생성 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
        
        # 본문을 이메일에 추가
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # 워드 파일 첨부
        print("📎 파일 첨부 중...")
        with open(docx_file_path, "rb") as attachment:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
            part.set_payload(attachment.read())
        
        # 파일을 base64로 인코딩
        encoders.encode_base64(part)
        
        # 헤더 추가
        filename = os.path.basename(docx_file_path)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {filename}',
        )
        
        # 이메일에 파일 첨부
        msg.attach(part)
        print(f"✅ 파일 '{filename}' 첨부 완료")
        
        # SMTP 서버 연결 및 이메일 전송
        print("📡 SMTP 서버 연결 중...")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # TLS 보안 연결 시작
        print("🔐 로그인 중...")
        server.login(sender_email, sender_password)
        
        # 이메일 전송
        print("📤 이메일 전송 중...")
        text = msg.as_string()
        server.sendmail(sender_email, recipient_emails, text)
        server.quit()
        
        print("\n" + "="*50)
        print("🎉 이메일이 성공적으로 전송되었습니다!")
        print(f"📧 받는 사람: {', '.join(recipient_emails)}")
        print(f"📎 첨부파일: {filename}")
        print(f"📅 전송 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*50)
        
        return True
        
    except smtplib.SMTPAuthenticationError:
        print("\n❌ 이메일 인증에 실패했습니다.")
        print("💡 해결 방법 (Gmail 기준):")
        print("   1. Google 계정 > 보안 > 2단계 인증 활성화")
        print("   2. 앱 비밀번호 생성")
        print("   3. 생성된 16자리 앱 비밀번호 사용")
        print("   4. 일반 비밀번호가 아닌 앱 비밀번호 입력 필요")
        return False
        
    except smtplib.SMTPRecipientsRefused as e:
        print(f"\n❌ 받는 사람 주소 오류: {e}")
        print("💡 이메일 주소를 다시 확인해주세요.")
        return False
        
    except smtplib.SMTPServerDisconnected:
        print("\n❌ SMTP 서버 연결이 끊어졌습니다.")
        print("💡 네트워크 연결을 확인하거나 잠시 후 다시 시도해주세요.")
        return False
        
    except smtplib.SMTPException as e:
        print(f"\n❌ SMTP 오류: {e}")
        return False
        
    except Exception as e:
        print(f"\n❌ 예상치 못한 오류가 발생했습니다: {e}")
        return False

def get_smtp_settings():
    """이메일 서비스별 SMTP 설정 반환"""
    smtp_settings = {
        '1': ('Gmail', 'smtp.gmail.com', 587),
        '2': ('Outlook/Hotmail', 'smtp-mail.outlook.com', 587),
        '3': ('Yahoo', 'smtp.mail.yahoo.com', 587),
        '4': ('Naver', 'smtp.naver.com', 587),
    }
    
    print("\n📡 이메일 서비스를 선택하세요:")
    for key, (name, server, port) in smtp_settings.items():
        print(f"{key}. {name} ({server}:{port})")
    print("5. 직접 입력")
    
    choice = input("\n선택 (1-5): ").strip()
    
    if choice in smtp_settings:
        name, server, port = smtp_settings[choice]
        print(f"✅ {name} 선택됨")
        return server, port
    elif choice == '5':
        server = input("SMTP 서버 주소: ")
        port = int(input("SMTP 포트 번호: "))
        return server, port
    else:
        print("❌ 잘못된 선택입니다. Gmail을 기본값으로 사용합니다.")
        return 'smtp.gmail.com', 587

def main():
    """메인 함수 - 대화형 이메일 전송"""
    print("🚀 보고서 이메일 전송 시스템")
    print("="*50)
    
    try:
        # 보고서 파일 확인
        default_file = 'sales_analysis_report.docx'
        if os.path.exists(default_file):
            print(f"📄 발견된 보고서: {default_file}")
            use_default = input("이 파일을 사용하시겠습니까? (y/n): ").lower().strip()
            
            if use_default == 'y':
                docx_file = default_file
            else:
                docx_file = input("워드 파일 경로를 입력하세요: ")
        else:
            print("❌ 기본 보고서 파일을 찾을 수 없습니다.")
            docx_file = input("워드 파일 경로를 입력하세요: ")
        
        if not os.path.exists(docx_file):
            print(f"❌ 파일을 찾을 수 없습니다: {docx_file}")
            return
        
        # 받는 사람 이메일 주소 입력
        print("\n📮 받는 사람 정보 입력")
        print("(여러 명에게 보낼 경우 쉼표로 구분)")
        print("예: user1@gmail.com, user2@company.com")
        
        recipients_input = input("\n받는 사람 이메일: ")
        recipient_emails = [email.strip() for email in recipients_input.split(',')]
        
        # 이메일 주소 검증
        invalid_emails = []
        for email in recipient_emails:
            if '@' not in email or '.' not in email.split('@')[1]:
                invalid_emails.append(email)
        
        if invalid_emails:
            print(f"❌ 잘못된 이메일 형식: {', '.join(invalid_emails)}")
            return
        
        print(f"✅ {len(recipient_emails)}명에게 전송 예정")
        
        # SMTP 설정 선택
        smtp_server, smtp_port = get_smtp_settings()
        
        # 사용자 정의 옵션
        print("\n✏️  이메일 내용 사용자 정의 (선택사항)")
        custom_subject = input("사용자 정의 제목 (엔터시 기본값 사용): ").strip()
        if not custom_subject:
            custom_subject = None
        
        print("\n📤 이메일 전송을 시작합니다...")
        
        # 이메일 전송 실행
        success = send_report_email(
            docx_file_path=docx_file,
            recipient_emails=recipient_emails,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            custom_subject=custom_subject
        )
        
        if success:
            print("\n🎊 모든 작업이 완료되었습니다!")
        else:
            print("\n❌ 이메일 전송에 실패했습니다.")
            print("💡 설정을 확인하고 다시 시도해주세요.")
            
    except KeyboardInterrupt:
        print("\n\n❌ 사용자가 취소했습니다.")
    except Exception as e:
        print(f"\n❌ 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()