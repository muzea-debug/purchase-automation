"""
Slack Webhook 알림 발송
"""
import json
import config

try:
    import requests as _requests
except ImportError:
    _requests = None


def 슬랙_발송(메시지: str) -> bool:
    """Webhook으로 메시지 발송. 성공 True, 실패/미설정 False."""
    if not config.SLACK_WEBHOOK_URL or _requests is None:
        print(f"[슬랙] 미설정 또는 requests 없음. 메시지: {메시지}")
        return False
    try:
        r = _requests.post(
            config.SLACK_WEBHOOK_URL,
            data=json.dumps({"text": 메시지}),
            headers={"Content-Type": "application/json"},
            timeout=5,
        )
        return r.status_code == 200
    except Exception as e:
        print(f"[슬랙] 발송 실패: {e}")
        return False


def 담당자_멘션(이름: str) -> str:
    """이름 → <@U12345> 형식. 맵에 없으면 이름 그대로."""
    uid = config.SLACK_담당자_맵.get(이름, "")
    return f"<@{uid}>" if uid else 이름


def pi_알림_발송(담당자: str, po코드: str, 차수: str, pct: float, 금액: float,
               제품목록: list = None) -> bool:
    """
    예: @홍길동 [NFA001] 제품A, 제품B 외 2건 30% 송금 리마인드 드립니다. (¥14,100)
    """
    멘션 = 담당자_멘션(담당자)

    제품명_str = ""
    if 제품목록:
        names = [
            p.get("product_name_kr") or p.get("description") or ""
            for p in 제품목록
            if p.get("product_name_kr") or p.get("description")
        ]
        if names:
            제품명_str = ", ".join(names[:3])
            if len(names) > 3:
                제품명_str += f" 외 {len(names) - 3}건"

    메시지 = f"{멘션} [{po코드}]"
    if 제품명_str:
        메시지 += f" {제품명_str}"
    메시지 += f" {int(pct * 100)}% 송금 리마인드 드립니다. (¥{금액:,.0f})"
    return 슬랙_발송(메시지)
