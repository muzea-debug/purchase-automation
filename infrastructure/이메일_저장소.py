"""
이메일_저장소 - SMTP 발주서 첨부 이메일 발송
"""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path


class 이메일_저장소:

    def __init__(
        self,
        smtp_서버: str,
        smtp_포트: int,
        발신자: str,
        비밀번호: str,
        수신자: str,
        참조: str = "",
    ):
        self._smtp_서버 = smtp_서버
        self._smtp_포트 = smtp_포트
        self._발신자 = 발신자
        self._비밀번호 = 비밀번호
        self._수신자 = 수신자
        self._참조 = 참조

    # 사유 → 중국어 매핑
    _사유_중문: dict[str, str] = {
        "신제품": "新订单",
        "정기발주": "定期订单",
    }

    def 발주서_발송(
        self,
        파일_경로들: "list[Path] | Path",
        po코드들: "list[str] | str",
        사유: str = "정기발주",
        총수량: int = 0,
    ) -> None:
        """발주서 xlsx(여러 개 가능)를 첨부하여 수신자(+참조)에게 이메일 발송."""
        # 단일 값도 허용
        if isinstance(파일_경로들, Path):
            파일_경로들 = [파일_경로들]
        if isinstance(po코드들, str):
            po코드들 = [po코드들]

        중문_유형 = self._사유_중문.get(사유, "定期订单")
        po코드_str = ", ".join(po코드들)

        msg = MIMEMultipart()
        msg["From"] = self._발신자
        msg["To"] = self._수신자
        if self._참조:
            msg["Cc"] = self._참조
        msg["Subject"] = f"请收到 {중문_유형} {po코드_str} {총수량}"

        본문 = f"请收到 {중문_유형} {po코드_str} {총수량}"
        msg.attach(MIMEText(본문, "plain", "utf-8"))

        for 파일_경로 in 파일_경로들:
            with open(파일_경로, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=("utf-8", "", 파일_경로.name),
                )
                msg.attach(part)

        수신자_목록 = [self._수신자]
        if self._참조:
            수신자_목록 += [a.strip() for a in self._참조.split(",") if a.strip()]

        # 포트 465 → SSL, 그 외 → STARTTLS
        if self._smtp_포트 == 465:
            with smtplib.SMTP_SSL(self._smtp_서버, self._smtp_포트) as server:
                server.ehlo("localhost")
                server.login(self._발신자, self._비밀번호)
                server.sendmail(self._발신자, 수신자_목록, msg.as_string())
        else:
            with smtplib.SMTP(self._smtp_서버, self._smtp_포트) as server:
                server.ehlo("localhost")
                server.starttls()
                server.ehlo("localhost")
                server.login(self._발신자, self._비밀번호)
                server.sendmail(self._발신자, 수신자_목록, msg.as_string())
