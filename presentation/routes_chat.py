"""
챗봇 + 피드백 API
"""
from __future__ import annotations
from datetime import datetime
from flask import Blueprint, request, jsonify
import config

bp_chat = Blueprint("chat", __name__)

_PAGE_CONTEXT: dict[str, str] = {
    "발주_실행": (
        "바코드 또는 엑셀 파일로 상품을 조회한 뒤 공급사별 발주서를 생성하고 이메일로 발송하는 화면입니다. "
        "진행 순서: 바코드+수량 직접 입력 또는 엑셀 업로드 → [상품 조회] → PO코드 확인 및 수정 → [발주 실행] → 이메일 발송."
    ),
    "재고현황": (
        "재고현황 엑셀 파일을 업로드해 SKU별 재고 수준과 발주 권고를 분석하는 화면입니다. "
        "파일 업로드 → 분석 실행 → 발주 권고 목록 확인 → Google Sheets 아카이브 저장 순서로 사용합니다."
    ),
    "dashboard": (
        "카테고리별 재고 추이와 소진 예상일을 시각화하는 대시보드입니다. "
        "좌측 카테고리 선택 → 우측 SKU 목록에서 상품 클릭 → 하단 상세 차트에서 예상 소진일(빨간 수직선) 확인."
    ),
    "pi_관리": (
        "선금계약서(Proforma Invoice)를 생성하고 관리하는 화면입니다. "
        "공급사·품목·금액 입력 후 PI Excel 생성, 슬랙 알림 발송, 이력 조회가 가능합니다."
    ),
    "발주서_생성": (
        "공급사에 제출할 원본 발주서 Excel 파일을 만드는 화면입니다. "
        "상품 이미지·정보·외박스 이미지를 입력하면 발주서 양식을 자동 생성합니다."
    ),
}

_SYSTEM = """\
너는 구매 자동화 시스템에 탑재된 AI 도우미 "소씽이"야.
이 시스템은 한국 회사 구매팀이 사용하는 발주 자동화 프로그램이야.

시스템 전체 기능:
- 원본 발주서 생성: 상품 정보로 공급사 제출용 발주서 Excel 생성
- 발주 실행: 바코드/수량 → 공급사별 발주서 생성 + 이메일 자동 발송
- 재고현황판: 재고 파일 업로드 → SKU별 발주 권고 분석
- 재고 대시보드: 카테고리별 재고 추이, 소진 예상일 시각화
- 선금계약서(PI) 관리: PI 생성 및 슬랙 알림

현재 사용자가 보고 있는 화면: {page}
{page_context}

소씽이 행동 지침:
- 항상 한국어로 자연스럽고 친근하게 말해 (딱딱한 공문체 금지)
- 현재 화면에서 할 수 있는 것 위주로 안내해
- 답변은 3~5문장으로 간결하게
- 모르는 건 솔직하게 모른다고 말해
- 자기소개 요청 시: 이름(소씽이)과 현재 화면에서 뭘 도와줄 수 있는지 한두 문장으로 소개해\
"""

_피드백_파일 = config.BASE_DIR / "output" / "피드백.txt"


@bp_chat.route("/api/chat", methods=["POST"])
def chat():
    if not config.ANTHROPIC_API_KEY:
        return jsonify({"success": False, "error": "ANTHROPIC_API_KEY가 설정되지 않았습니다."}), 503

    data = request.get_json(force=True)
    message = str(data.get("message", "")).strip()
    page = str(data.get("page", "")).strip()
    history = data.get("history", [])

    if not message:
        return jsonify({"success": False, "error": "메시지를 입력해주세요."}), 400

    # 패널 첫 열림 시 자동 인사
    if message == "__intro__":
        message = "지금 내가 보고 있는 화면이 어떤 화면인지 파악하고, 소씽이로서 짧게 자기소개 하면서 이 화면에서 뭘 도와줄 수 있는지 자연스럽게 한두 문장으로 말해줘."

    system = _SYSTEM.format(
        page=page or "알 수 없음",
        page_context=_PAGE_CONTEXT.get(page, ""),
    )

    messages = []
    for h in history[-10:]:
        role = h.get("role", "")
        content = h.get("content", "")
        if role in ("user", "assistant") and content:
            messages.append({"role": role, "content": content})
    messages.append({"role": "user", "content": message})

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=config.ANTHROPIC_API_KEY)
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            system=system,
            messages=messages,
        )
        return jsonify({"success": True, "reply": resp.content[0].text})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@bp_chat.route("/api/feedback", methods=["POST"])
def feedback_save():
    data = request.get_json(force=True)
    page = str(data.get("page", "")).strip() or "알 수 없음"
    content = str(data.get("content", "")).strip()

    if not content:
        return jsonify({"success": False, "error": "내용을 입력해주세요."}), 400

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] [{page}]\n{content}\n{'─' * 50}\n"

    _피드백_파일.parent.mkdir(parents=True, exist_ok=True)
    with _피드백_파일.open("a", encoding="utf-8") as f:
        f.write(entry)

    return jsonify({"success": True})
