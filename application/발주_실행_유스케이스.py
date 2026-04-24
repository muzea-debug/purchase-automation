"""
발주_실행_유스케이스 - 상품 조회 및 발주 실행 흐름 조율
"""
from __future__ import annotations
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import config
from infrastructure.발주서원본_파서 import 발주서원본_파서
from infrastructure.발주서_실행_저장소 import 발주서_실행_저장소
from infrastructure.google_sheets_저장소 import GoogleSheets저장소
from infrastructure.엑셀_시트_저장소 import 엑셀시트저장소
from infrastructure.이메일_저장소 import 이메일_저장소
from domain.services.발주_코드_서비스 import 발주_코드_서비스


class 발주_실행_유스케이스:

    def __init__(self):
        self._파서 = 발주서원본_파서(config.통합발주서원본_DIR)
        self._저장소 = 발주서_실행_저장소(config.발주서_DIR)
        self._sheets: Optional[GoogleSheets저장소] = self._sheets_연결()
        self._이메일: Optional[이메일_저장소] = self._이메일_연결()

    # ── 상품 조회 ────────────────────────────────────────────────

    def 조회(self, items: list[dict]) -> dict:
        """
        items: [{"barcode": "...", "quantity": 500}, ...]
        반환: {"success": True, "groups": [...], "미확인": [...]}
        """
        카탈로그 = self._파서.카탈로그_로드()
        lt_맵 = self._lt_맵_조회()
        제품DB = self._제품DB_읽기()   # 카테고리 정보 포함
        그룹_맵: dict[str, dict] = {}
        미확인: list[str] = []

        for item in items:
            바코드 = str(item.get("barcode", "")).strip()
            수량 = int(item.get("quantity", 0))

            if 바코드 not in 카탈로그:
                미확인.append(바코드)
                continue

            상품 = 카탈로그[바코드]
            키 = 상품.원본파일.name

            if 키 not in 그룹_맵:
                카테고리1 = 제품DB.get(바코드, {}).get("카테고리_1", "")
                base = 발주_코드_서비스.브랜드_po_base(상품.브랜드, 카테고리1)
                prefix, 현재번호 = self._현재_활성_prefix(base)
                그룹_맵[키] = {
                    "제조사": 상품.제조사,
                    "브랜드": 상품.브랜드,
                    "원본파일": 키,
                    "po_prefix": prefix,
                    "po_번호_제안": 현재번호 + 1,
                    "po_base": base,
                    "max_lt": 0,
                    "계좌id": 상품.계좌id,
                    "items": [],
                }

            lt = int(lt_맵.get(바코드, 0))
            그룹_맵[키]["max_lt"] = max(그룹_맵[키]["max_lt"], lt)

            그룹_맵[키]["items"].append({
                "barcode": 바코드,
                "이름": 상품.이름,
                "원가": 상품.원가,
                "수량": 수량,
            })

        return {
            "success": True,
            "groups": list(그룹_맵.values()),
            "미확인": 미확인,
        }

    # ── 발주 실행 ────────────────────────────────────────────────

    def 실행(self, payload: dict) -> dict:
        """
        payload:
          groups: [{원본파일, po코드, max_lt, items: [{barcode, quantity}]}]
        """
        카탈로그 = self._파서.카탈로그_로드()
        제품DB = self._제품DB_읽기()

        결과_목록 = []
        전체_출력_경로: list[Path] = []
        전체_po코드: list[str] = []
        전체_수량 = 0
        대표_사유 = "정기발주"
        발주현황_행들: list[dict] = []

        오늘 = date.today()
        발주일 = f"{오늘.year}. {오늘.month}. {오늘.day}"

        for 그룹 in payload.get("groups", []):
            원본파일 = 그룹["원본파일"]
            po코드 = 그룹["po코드"]
            items = 그룹["items"]
            max_lt = int(그룹.get("max_lt", 0))

            if not 발주_코드_서비스.유효성_검사(po코드):
                return {"success": False, "error": f"잘못된 PO 코드: {po코드}"}

            # 예상 도착일: 발주일 + LT 최대 + 5일 → 다음 화/금
            예상_도착일 = self._예상도착일_계산(오늘, max_lt)

            원본_경로 = config.통합발주서원본_DIR / 원본파일
            출력_경로 = self._저장소.출력_경로_계산("정기발주", po코드)
            바코드_수량 = {str(i["barcode"]): int(i["quantity"]) for i in items}

            self._저장소.발주서_생성(
                원본_경로=원본_경로,
                출력_경로=출력_경로,
                바코드_수량=바코드_수량,
                예상_도착일=예상_도착일,
            )

            첫_바코드 = str(items[0]["barcode"])
            첫_상품 = 카탈로그.get(첫_바코드)
            제조사명 = 첫_상품.제조사 if 첫_상품 else ""

            # Google Sheets 기록 (엔대시 발주이력)
            sheets_ok = self._sheets_기록(items, po코드, 제조사명, 카탈로그)

            # 중국발주요청리스트 추가
            발주요청_ok = self._중국발주요청리스트_기록(items, po코드, 예상_도착일, 제조사명, 카탈로그)

            # N월 발주현황 행 수집
            for item in items:
                bc = str(item["barcode"])
                상품 = 카탈로그.get(bc)
                db = 제품DB.get(bc, {})
                발주현황_행들.append({
                    "발주일": 발주일,
                    "바코드": bc,
                    "주문번호": po코드,
                    "발주수량": int(item["quantity"]),
                    "제품명": 상품.이름 if 상품 else "",
                    "브랜드명": 상품.브랜드 if 상품 else "",
                    "업체명": 상품.제조사 if 상품 else "",
                    "비고": "",
                    "색상코드": "",
                    "시즌구분": db.get("시즌구분", ""),
                    "카테고리_1": db.get("대분류", ""),
                    "카테고리_2": db.get("소분류", ""),
                    "단가": db.get("최종원가", ""),
                })

            전체_출력_경로.append(출력_경로)
            전체_po코드.append(po코드)
            전체_수량 += sum(int(i["quantity"]) for i in items)
            대표_사유 = 그룹.get("사유", "정기발주")

            결과_목록.append({
                "po코드": po코드,
                "파일": str(출력_경로.relative_to(config.발주서_DIR)),
                "예상_도착일": 예상_도착일,
                "sheets_기록": sheets_ok,
                "중국발주요청_완료": 발주요청_ok,
            })

        # N월 발주현황 시트 기록
        self._발주현황_기록(발주현황_행들)

        # 이메일은 자동 발송하지 않음 → 클라이언트에서 버튼으로 발송
        이메일_준비 = {
            "파일_경로들": [str(p) for p in 전체_출력_경로],
            "po코드들": 전체_po코드,
            "사유": 대표_사유,
            "총수량": 전체_수량,
        }

        return {"success": True, "results": 결과_목록, "이메일_준비": 이메일_준비}

    def 이메일_발송_직접(self, 파일_경로들: list, po코드들: list, 사유: str, 총수량: int) -> str:
        """수동 호출용 이메일 발송. 반환: '' 성공 / '미설정' / 오류메시지"""
        경로들 = [Path(p) for p in 파일_경로들]
        return self._이메일_발송(경로들, po코드들, 사유, 총수량)

    # ── 내부 헬퍼 ────────────────────────────────────────────────

    def _현재_활성_prefix(self, base: str) -> tuple[str, int]:
        """base 시리즈에서 현재 활성 prefix와 최대번호 반환."""
        def 최대_조회(prefix: str) -> int:
            if self._sheets:
                try:
                    n = self._sheets.마지막_po번호_조회(prefix)
                    if n > 0:
                        return n
                except Exception as e:
                    print(f"[발주_실행] Sheets PO번호 조회 실패: {e}")
            return 발주_코드_서비스.최대번호_로컬(config.발주서_DIR, prefix)

        return 발주_코드_서비스.현재_활성_prefix(base, 최대_조회, config.PO_번호_상한)

    def _다음_po번호(self, prefix: str) -> int:
        """단일 prefix 기준 다음 번호 (하위호환용)."""
        if self._sheets:
            try:
                n = self._sheets.마지막_po번호_조회(prefix)
                if n > 0:
                    return n + 1
            except Exception as e:
                print(f"[발주_실행] Sheets PO번호 조회 실패: {e}")
        return 발주_코드_서비스.다음_번호_로컬(config.발주서_DIR, prefix)

    def _제품DB_읽기(self) -> dict:
        """바코드 → 제품DB 메타. sheets 없으면 빈 dict."""
        if not self._sheets:
            return {}
        try:
            return self._sheets.제품DB_읽기()
        except Exception as e:
            print(f"[발주_실행] 제품DB 읽기 실패: {e}")
            return {}

    def _lt_맵_조회(self) -> dict[str, float]:
        """바코드 → 리드타임(일) 맵. sheets 없으면 빈 dict."""
        제품DB = self._제품DB_읽기()
        return {bc: info.get("기본_lt", 0) for bc, info in 제품DB.items()}

    def _예상도착일_계산(self, 발주일: date, lt_일: int) -> str:
        """발주일 + lt + 5일 → 다음 화요일(1) 또는 금요일(4), 공휴일 제외."""
        목표일 = 발주일 + timedelta(days=lt_일 + 5)
        for delta in range(0, 14):
            d = 목표일 + timedelta(days=delta)
            if d.weekday() in (1, 4) and not self._한국공휴일(d):
                return f"{d.year}. {d.month}. {d.day}"
        return f"{목표일.year}. {목표일.month}. {목표일.day}"

    @staticmethod
    def _한국공휴일(d: date) -> bool:
        try:
            import holidays
            return d in holidays.KR(years=d.year)
        except ImportError:
            return False

    def _sheets_기록(self, items, po코드, 제조사명, 카탈로그) -> bool:
        if not self._sheets:
            return False
        try:
            for item in items:
                바코드 = str(item["barcode"])
                상품 = 카탈로그.get(바코드)
                self._sheets.발주_기록(
                    바코드=바코드,
                    po코드=po코드,
                    수량=int(item["quantity"]),
                    제품명=상품.이름 if 상품 else "",
                    제조사=제조사명,
                )
            return True
        except Exception as e:
            print(f"[발주_실행] Sheets 기록 실패: {e}")
            return False

    def _중국발주요청리스트_기록(self, items, po코드, 예상_도착일, 제조사명, 카탈로그) -> bool:
        if not self._sheets or not hasattr(self._sheets, "중국발주요청리스트_추가"):
            return False
        try:
            행들 = []
            for item in items:
                바코드 = str(item["barcode"])
                상품 = 카탈로그.get(바코드)
                행들.append({
                    "바코드":    바코드,
                    "po코드":    po코드,
                    "수량":      int(item["quantity"]),
                    "제품명":    상품.이름 if 상품 else "",
                    "입고예정일": 예상_도착일,
                    "제조사":    제조사명,
                })
            self._sheets.중국발주요청리스트_추가(행들)
            return True
        except Exception as e:
            print(f"[발주_실행] 중국발주요청리스트 기록 실패: {e}")
            return False

    def _발주현황_기록(self, 행들: list[dict]) -> bool:
        if not self._sheets or not 행들:
            return False
        try:
            self._sheets.발주현황_기록(행들)
            return True
        except Exception as e:
            print(f"[발주_실행] 발주현황 기록 실패: {e}")
            return False

    def _이메일_발송(self, 출력_경로들: list, po코드들: list, 사유: str, 총수량: int) -> str:
        """성공: '', 미설정: '미설정', 실패: 오류메시지"""
        if not self._이메일:
            return "미설정"
        try:
            self._이메일.발주서_발송(
                파일_경로들=출력_경로들,
                po코드들=po코드들,
                사유=사유,
                총수량=총수량,
            )
            return ""
        except Exception as e:
            import traceback
            traceback.print_exc()
            return str(e)

    def _sheets_연결(self):
        """Google Sheets 우선, 없으면 로컬 Excel 파일 사용."""
        if config.GOOGLE_SERVICE_ACCOUNT_JSON and config.GOOGLE_SHEET_ID:
            try:
                return GoogleSheets저장소(
                    config.GOOGLE_SERVICE_ACCOUNT_JSON,
                    config.GOOGLE_SHEET_ID,
                    config.GOOGLE_SHEET_NAME,
                )
            except Exception as e:
                print(f"[발주_실행] Google Sheets 연결 실패: {e}")

        if config.로컬시트_파일.exists():
            return 엑셀시트저장소(config.로컬시트_파일)

        return None

    def _이메일_연결(self) -> Optional[이메일_저장소]:
        if not config.SMTP_발신자 or not config.SMTP_비밀번호:
            return None
        # JSON 설정 파일 우선, 없으면 env 값 fallback
        수신자 = config.SMTP_수신자
        참조   = config.SMTP_참조
        try:
            if config.이메일_설정_파일.exists():
                import json as _json
                설정 = _json.loads(config.이메일_설정_파일.read_text(encoding="utf-8"))
                수신자 = 설정.get("수신자") or 수신자
                참조   = 설정.get("참조", 참조)
        except Exception:
            pass
        if not 수신자:
            return None
        return 이메일_저장소(
            config.SMTP_서버,
            config.SMTP_포트,
            config.SMTP_발신자,
            config.SMTP_비밀번호,
            수신자,
            참조,
        )
