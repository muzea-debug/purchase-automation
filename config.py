"""
발주 자동화 - 설정 파일
경로가 바뀌면 이 파일만 수정
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent

# 브랜드별 외박스 기본 이미지 (원본 발주서에서 추출)
ILYON_외박스_기본이미지        = BASE_DIR / "assets" / "ilyon_외박스_기본이미지.png"
누아트스튜디오_외박스_기본이미지 = BASE_DIR / "assets" / "nuart_studio_외박스_기본이미지.png"

# 통합발주서원본 폴더 (저장 대상) - Docker 환경에서는 사용 안 함 (브라우저 다운로드)
발주서원본_DIR = Path(os.environ.get("발주서원본_DIR", str(BASE_DIR / "output")))

# 이미지 임시 업로드 폴더
UPLOAD_DIR = BASE_DIR / "uploads"

# Flask 설정
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", 5006))
DEBUG = os.environ.get("DEBUG", "false").lower() == "true"

# ── 발주 실행 경로 ─────────────────────────────────────────────
통합발주서원본_DIR = Path(os.environ.get(
    "통합발주서원본_DIR",
    str(BASE_DIR / "통합발주서원본_test")   # 로컬 테스트용 fallback
))
발주서_DIR = Path(os.environ.get(
    "발주서_DIR",
    str(BASE_DIR / "발주서_output")         # 로컬 테스트용 fallback
))

# ── Google Sheets ──────────────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_SHEET_NAME = os.environ.get("GOOGLE_SHEET_NAME", "엔대시")

# 엔대시 발주요청 시트명 (재고현황 자동화에서 직접 읽는 시트)
ENDASH_발주요청_시트 = os.environ.get("ENDASH_발주요청_시트", "[엔대시]중국 발주요청리스트(交期管理表)")
# 엔대시 아카이브 시트명 (발주요청 데이터를 날짜별로 기록)
ENDASH_아카이브_시트 = os.environ.get("ENDASH_아카이브_시트", "재고현황_아카이브")

# ── 재고현황 업로드 임시 폴더 ──────────────────────────────────
재고현황_UPLOAD_DIR = BASE_DIR / "uploads" / "재고현황"

# ── 재고현황 최근 분석 결과 저장 ────────────────────────────────
재고현황_저장_파일 = BASE_DIR / "output" / "재고현황_최근.json"

# ── 재고현황 누적 DB (날짜별 스냅샷) ────────────────────────────
재고현황_아카이브_DIR = BASE_DIR / "재고현황_아카이브"
재고현황_누적_파일   = BASE_DIR / "output" / "재고현황_누적.json"

# ── 로컬 시트 (Google Sheets 대체용 Excel) ──────────────────────
로컬시트_DIR  = BASE_DIR / "구글 시트 대체용"
로컬시트_파일 = 로컬시트_DIR / "엔대시_로컬.xlsx"

# 초기화용 소스 파일 경로 (없으면 빈 탭으로 생성)
로컬시트_엔대시소스 = Path(os.environ.get(
    "로컬시트_엔대시소스",
    ""  # 비어있으면 빈 발주요청 탭 생성
))
로컬시트_주문프로그램 = Path(os.environ.get(
    "로컬시트_주문프로그램",
    ""  # 비어있으면 빈 제품DB 탭 생성
))
로컬시트_재고현황표 = Path(os.environ.get(
    "로컬시트_재고현황표",
    ""  # 비어있으면 카테고리 없이 생성
))

# ── 브랜드 → PO 코드 base (일반 상품) ────────────────────────────
# 실제 prefix = base + 알파벳(A~Z), 번호 1~999 후 다음 알파벳 자동 전환
BRAND_PO_BASE: dict[str, str] = {
    "NUART":       "N",
    "누아트스튜디오": "N",
    "ILYON":       "I",
    "일리온":       "I",
    "패션":         "F",
    "FASHION":     "F",
}

# ── 브랜드 → PO 코드 base (전장품 카테고리 전용) ────────────────
BRAND_전장품_PO_BASE: dict[str, str] = {
    "NUART":       "EN",
    "누아트스튜디오": "EN",
    "ILYON":       "EI",
    "일리온":       "EI",
}

# ── 전장품으로 인식할 카테고리1 값 ──────────────────────────────
전장품_카테고리: frozenset = frozenset({"전장품", "전장"})

# ── PO 번호 상한 (초과 시 다음 알파벳으로 전환) ─────────────────
PO_번호_상한: int = 999

# ── 대시보드 표시용 시리즈 정의 ──────────────────────────────────
PO_시리즈: list[dict] = [
    {"base": "N",  "name": "누아트"},
    {"base": "I",  "name": "일리온"},
    {"base": "F",  "name": "패션"},
    {"base": "EN", "name": "누아트전장품"},
    {"base": "EI", "name": "일리온전장품"},
]

# 하위호환용 (기존 코드 참조 시 대비)
BRAND_PO_PREFIX: dict[str, str] = {k: v + "A" for k, v in BRAND_PO_BASE.items()}

# ── 이메일 수신자 설정 (UI에서 관리, env 보다 우선) ──────────────
이메일_설정_파일 = BASE_DIR / "output" / "이메일_설정.json"

# ── SMTP ───────────────────────────────────────────────────────
SMTP_서버 = os.environ.get("SMTP_SERVER", "smtp.naver.com")
SMTP_포트 = int(os.environ.get("SMTP_PORT", "587"))
SMTP_발신자 = os.environ.get("SMTP_FROM", "")
SMTP_비밀번호 = os.environ.get("SMTP_PASSWORD", "")
SMTP_수신자 = os.environ.get("SMTP_TO", "")    # 중국지사 담당자 이메일
SMTP_참조 = os.environ.get("SMTP_CC", "")      # 참조(CC)

# ── PI (Proforma Invoice) ───────────────────────────────────────
PI_DIR = BASE_DIR / "output" / "PI"
PI_데이터_파일 = BASE_DIR / "output" / "pi_데이터.json"
PI_템플릿 = BASE_DIR / "assets" / "pi_template.xlsx"
재고현황_템플릿 = BASE_DIR / "assets" / "재고현황_템플릿.xlsx"
누아트_도장 = BASE_DIR / "assets" / "nuart_도장.png"
업체도장_DIR = BASE_DIR / "assets" / "업체도장"
계좌DB_파일 = BASE_DIR / "output" / "계좌_db.json"

# ── Slack ──────────────────────────────────────────────────────
SLACK_WEBHOOK_URL = os.environ.get("SLACK_WEBHOOK_URL", "")
SLACK_담당자_맵_JSON = os.environ.get("SLACK_담당자_맵", "{}")

import json as _json
try:
    SLACK_담당자_맵: dict[str, str] = _json.loads(SLACK_담당자_맵_JSON)
except Exception:
    SLACK_담당자_맵 = {}

# ── Anthropic (스크린샷 OCR) ───────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
