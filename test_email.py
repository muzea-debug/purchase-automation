"""
이메일 발송 테스트 스크립트
실행: python test_email.py
설정은 .env 파일에서 읽음 (SMTP_SERVER, SMTP_PORT, SMTP_FROM, SMTP_PASSWORD, SMTP_TO, SMTP_CC)
"""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os

load_dotenv()

SMTP_서버 = os.environ.get("SMTP_SERVER", "smtp.naver.com")
SMTP_포트 = int(os.environ.get("SMTP_PORT", "465"))
발신자    = os.environ.get("SMTP_FROM", "")
비밀번호  = os.environ.get("SMTP_PASSWORD", "")
수신자    = os.environ.get("SMTP_TO", "")
참조      = os.environ.get("SMTP_CC", "")


def test():
    if not 발신자 or not 비밀번호:
        print("SMTP_FROM, SMTP_PASSWORD 가 .env 에 없습니다.")
        return

    msg = MIMEMultipart()
    msg["From"]    = 발신자
    msg["To"]      = 수신자
    msg["Cc"]      = 참조
    msg["Subject"] = "[발주 자동화] 이메일 테스트"

    본문 = "발주 자동화 이메일 테스트입니다.\n정상 수신되면 설정 완료입니다."
    msg.attach(MIMEText(본문, "plain", "utf-8"))

    수신자_목록 = [수신자] + [a.strip() for a in 참조.split(",") if a.strip()]

    print(f"접속 중: {SMTP_서버}:{SMTP_포트}")
    try:
        if SMTP_포트 == 465:
            with smtplib.SMTP_SSL(SMTP_서버, SMTP_포트) as server:
                server.ehlo("localhost")
                print("로그인 중...")
                server.login(발신자, 비밀번호)
                server.sendmail(발신자, 수신자_목록, msg.as_string())
        else:
            with smtplib.SMTP(SMTP_서버, SMTP_포트) as server:
                server.ehlo("localhost")
                server.starttls()
                server.ehlo("localhost")
                print("로그인 중...")
                server.login(발신자, 비밀번호)
                server.sendmail(발신자, 수신자_목록, msg.as_string())
        print(f"발송 완료 -> {수신자_목록}")
    except smtplib.SMTPAuthenticationError:
        print("로그인 실패: 이메일/비밀번호 확인")
    except smtplib.SMTPException as e:
        print(f"SMTP 오류: {e}")
    except Exception as e:
        print(f"오류: {e}")


if __name__ == "__main__":
    test()
