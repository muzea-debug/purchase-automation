"""
재고현황_유스케이스 - 7개 파일 + 엔대시 Sheets 읽기 → 재고현황판 데이터 생성
"""
from __future__ import annotations
from datetime import date, datetime
from pathlib import Path
from typing import Optional, List

import config
from infrastructure.readers.재고현황표_리더 import 재고현황표_리더
from infrastructure.readers.사방넷_리더 import 사방넷_리더
from infrastructure.readers.basic_리더 import Basic_리더
from infrastructure.readers.forecast_리더 import Forecast_리더
from infrastructure.readers.누아트sku_리더 import 누아트SKU_리더
from infrastructure.readers.po_리더 import PO_리더
from infrastructure.readers.주문서_리더 import 주문서_리더
from infrastructure.readers.엔대시_파일_리더 import 엔대시파일_리더
from infrastructure.readers._공통 import 컬럼_찾기, 안전_숫자, 안전_문자


class 재고현황_유스케이스:

    def __init__(self, sheets_저장소=None):
        self._sheets = sheets_저장소
        self._마스터_리더  = 재고현황표_리더()
        self._사방넷_리더  = 사방넷_리더()
        self._basic_리더   = Basic_리더()
        self._forecast_리더 = Forecast_리더()
        self._sku_리더     = 누아트SKU_리더()
        self._po_리더      = PO_리더()
        self._주문서_리더  = 주문서_리더()

    # ── 공개 API ────────────────────────────────────────────────

    def 분석(
        self,
        사방넷_파일: Optional[Path] = None,
        basic_파일: Optional[Path] = None,
        forecast_파일: Optional[Path] = None,
        누아트sku_파일: Optional[Path] = None,
        po_파일들: Optional[list[Path]] = None,
        주문서_파일: Optional[Path] = None,
        재고현황표_파일: Optional[Path] = None,
        제외_날짜들: list[str] | None = None,
        기존_행맵: dict | None = None,
        엔대시_파일: Optional[Path] = None,
    ) -> dict:
        """
        전체 재고현황 계산.
        제품 마스터: sheets_저장소.제품DB_읽기() 우선, 없으면 재고현황표_파일 사용.

        Returns:
          {
            "success": True,
            "입고_그룹": [...],   # 오늘 이전 누아트입고일자 그룹
            "rows": [...],        # 재고현황 행 목록
            "요약": {...}
          }
        """
        제외_날짜들 = 제외_날짜들 or []
        기존_행맵 = 기존_행맵 or {}

        def _기존(바코드, 키, 기본=0.0):
            """파일 미업로드 시 기존 저장 데이터에서 값 반환."""
            return 기존_행맵.get(바코드, {}).get(키, 기본)

        # 1. 제품 마스터 — 로컬 Excel 제품DB 우선 (나중엔 스프레드시트), 없으면 재고현황표 파일
        마스터: dict[str, dict] = {}
        if self._sheets and hasattr(self._sheets, '제품DB_읽기'):
            try:
                마스터 = self._sheets.제품DB_읽기()
                print(f"[재고현황] 제품DB에서 {len(마스터)}개 제품 로드")
            except Exception as e:
                print(f"[재고현황] 제품DB 읽기 실패: {e}")
        if not 마스터 and 재고현황표_파일:
            마스터 = self._마스터_리더.읽기(재고현황표_파일)

        # 2. 엔대시 Sheets 읽기 (파일 직접 지정 가능)
        엔대시_데이터, 입고_그룹 = self._엔대시_읽기(제외_날짜들, 엔대시_파일=엔대시_파일)

        # 3. 나머지 소스 파일 (None이면 None 유지 → 기존 값 fallback 처리)
        사방넷   = self._사방넷_리더.읽기(사방넷_파일) if 사방넷_파일 else None
        basic    = self._basic_리더.읽기(basic_파일) if basic_파일 else None
        sku_맵   = self._sku_리더.읽기(누아트sku_파일) if 누아트sku_파일 else {}
        forecast = self._forecast_리더.읽기(forecast_파일, sku_맵) if forecast_파일 else None
        po       = self._po_리더.읽기(po_파일들) if po_파일들 else None
        주문서   = self._주문서_리더.읽기(주문서_파일) if 주문서_파일 else None

        # 4. 행 조립
        rows = []
        for 바코드, 마스터_행 in 마스터.items():
            엔 = 엔대시_데이터.get(바코드, {})
            # 구글 시트 데이터가 있으면 사용, 없으면 기존 저장 값 유지
            시트_있음 = bool(엔대시_데이터) or self._sheets is not None
            중국재고_쿠팡  = 엔.get("중국재고_쿠팡",  _기존(바코드, "중국재고_쿠팡"))  if 시트_있음 else _기존(바코드, "중국재고_쿠팡")
            중국재고_일반  = 엔.get("중국재고_일반",  _기존(바코드, "중국재고_일반"))  if 시트_있음 else _기존(바코드, "중국재고_일반")
            중국쿠팡_미입고 = 엔.get("중국쿠팡_미입고", _기존(바코드, "중국쿠팡_미입고")) if 시트_있음 else _기존(바코드, "중국쿠팡_미입고")
            중국일반_미입고 = 엔.get("중국일반_미입고", _기존(바코드, "중국일반_미입고")) if 시트_있음 else _기존(바코드, "중국일반_미입고")
            입고예정_코드  = 엔.get("입고예정_코드",  _기존(바코드, "입고예정_코드", ""))
            입고일자       = 엔.get("입고일자",       _기존(바코드, "입고일자", ""))

            사방넷재고   = 사방넷.get(바코드, 0.0)  if 사방넷  is not None else _기존(바코드, "사방넷재고")
            fc재고       = basic.get(바코드, 0.0)   if basic    is not None else _기존(바코드, "쿠팡fc재고")
            쿠팡주간판매  = forecast.get(바코드, 0.0) if forecast is not None else _기존(바코드, "쿠팡주간판매")
            쿠팡발주수량  = po.get(바코드, 0.0)      if po       is not None else _기존(바코드, "쿠팡발주수량")
            일반채널판매  = 주문서.get(바코드, 0.0)  if 주문서   is not None else _기존(바코드, "일반채널판매")

            재고합계 = (
                중국재고_쿠팡 + 중국재고_일반
                + 사방넷재고
                + 중국쿠팡_미입고 + 중국일반_미입고
            )  # 쿠팡FC재고(fc재고)는 제외

            # 리드타임 먼저 계산
            lt = 마스터_행.get("기본_lt", 0) or 0
            _안전 = lt * 2.5
            if _안전 < 14:
                안전재고일수 = 42
            else:
                안전재고일수 = min(90, round(_안전))

            # 총 리드타임: 양산(기본_lt 일) + 운송 5일, 주 단위로 환산
            총_lt_주 = (lt + 5) / 7
            예상판매_lt = 쿠팡발주수량 + (쿠팡주간판매 + 일반채널판매) * 총_lt_주
            # 판매 데이터 없으면 발주량 산출 불가 → 0
            발주량산출 = round(재고합계 - 예상판매_lt, 1) if 예상판매_lt > 0 else 0

            # 재고 상태 분류
            상태 = self._재고_상태(발주량산출, 재고합계, 예상판매_lt)

            # 재고소진일: 재고합계 / 일일판매량
            일판매 = (쿠팡주간판매 + 일반채널판매) / 7
            재고소진일 = round(재고합계 / 일판매) if 일판매 > 0 else None

            rows.append({
                **마스터_행,
                "중국재고_쿠팡":   round(중국재고_쿠팡, 0),
                "중국재고_일반":   round(중국재고_일반, 0),
                "사방넷재고":      round(사방넷재고, 0),
                "쿠팡fc재고":      round(fc재고, 0),
                "입고예정_코드":   입고예정_코드,
                "입고일자":        입고일자,
                "중국쿠팡_미입고": round(중국쿠팡_미입고, 0),
                "중국일반_미입고": round(중국일반_미입고, 0),
                "쿠팡발주수량":    round(쿠팡발주수량, 0),
                "쿠팡주간판매":    round(쿠팡주간판매, 1),
                "일반채널판매":    round(일반채널판매, 0),
                "재고합계":        round(재고합계, 0),
                "재고소진일":      재고소진일,
                "안전재고일수":    안전재고일수,
                "예상판매_lt":     round(예상판매_lt, 1),
                "발주량산출":      round(발주량산출, 1),
                "재고상태":        상태,
            })

        # 5. 요약
        요약 = self._요약_계산(rows)

        return {
            "success": True,
            "입고_그룹": 입고_그룹,
            "rows": rows,
            "요약": 요약,
        }

    def 분석_아카이브(self, 폴더명: str) -> dict:
        """
        아카이브 폴더의 소스 파일로 과거 시점 재고현황 분석.
        폴더명: YYYY.MM.DD 형식 (재고현황_아카이브/ 하위)
        """
        from infrastructure.아카이브_스캐너 import 아카이브_스캐너
        scanner = 아카이브_스캐너()
        try:
            파일맵 = scanner.날짜_파일맵(폴더명)
        except FileNotFoundError as e:
            return {"success": False, "error": str(e)}

        엔대시_파일 = 파일맵.get("엔대시")
        po_파일들 = 파일맵.get("po_목록") or ([파일맵["po"]] if "po" in 파일맵 else [])

        return self.분석(
            사방넷_파일=파일맵.get("사방넷"),
            basic_파일=파일맵.get("basic"),
            forecast_파일=파일맵.get("forecast"),
            누아트sku_파일=파일맵.get("누아트sku"),
            po_파일들=po_파일들 or None,
            주문서_파일=파일맵.get("주문서"),
            엔대시_파일=엔대시_파일,
        )

    # ── 엔대시 Sheets 읽기 ──────────────────────────────────────

    def _엔대시_읽기(self, 제외_날짜들: list[str], 엔대시_파일: Optional[Path] = None) -> tuple[dict, list]:
        """
        엔대시 발주요청 시트 읽기.
        엔대시_파일 지정 시 xlsx 파일 직접 읽기 (아카이브 분석용).
        Returns: (바코드별_집계, 입고확인_그룹목록)
        """
        if 엔대시_파일:
            try:
                헤더, 행들 = 엔대시파일_리더().읽기(엔대시_파일)
            except Exception as e:
                print(f"[재고현황] 엔대시 파일 읽기 실패: {e}")
                return {}, []
        elif self._sheets:
            try:
                헤더, 행들 = self._sheets.발주요청_읽기(config.ENDASH_발주요청_시트)
            except Exception as e:
                print(f"[재고현황] 엔대시 읽기 실패: {e}")
                return {}, []
        else:
            return {}, []

        # 컬럼 인덱스 — 시트 구조 고정값 사용
        # A(0)=쿠팡중국재고, B(1)=일반중국재고, C(2)=발주일자,
        # D(3)=바코드, E(4)=발주번호, F(5)=쿠팡미입고, G(6)=일반미입고,
        # H(7)=제품명, I(8~)=비고/중국직원 사용
        쿠팡재고_col     = 0
        일반재고_col     = 1
        바코드_col       = 3
        발주번호_col     = 4
        쿠팡발주수량_col  = 5
        일반수량_col     = 6
        제품명_col       = 7
        입고일자_col     = 10  # 누아트 입고일자 (K열)

        오늘 = date.today()

        # 입고 확인이 필요한 행 그룹핑 (날짜별)
        입고_날짜_그룹: dict[str, list] = {}
        유효_행: list[list] = []

        for row in 행들:
            def _v(col):
                if col is None or len(row) <= col:
                    return ''
                return row[col]

            입고일자_raw = _v(입고일자_col)
            입고_날짜 = self._날짜_파싱(입고일자_raw)

            if 입고_날짜 and 입고_날짜 < 오늘:
                날짜_str = 입고_날짜.strftime("%Y-%m-%d")
                if 날짜_str not in 제외_날짜들:
                    # 입고 확인 필요
                    if 날짜_str not in 입고_날짜_그룹:
                        입고_날짜_그룹[날짜_str] = []
                    입고_날짜_그룹[날짜_str].append({
                        "바코드":   안전_문자(_v(바코드_col)),
                        "제품명":   안전_문자(_v(제품명_col)),
                        "발주번호": 안전_문자(_v(발주번호_col)),
                        "수량":     안전_숫자(_v(쿠팡발주수량_col)) + 안전_숫자(_v(일반수량_col)),
                        "입고일자": 입고일자_raw,
                    })
                    continue  # 제외 날짜 미지정 시 이 행은 집계에서 제외

            유효_행.append(row)

        # 유효 행으로 바코드별 집계
        집계: dict[str, dict] = {}
        for row in 유효_행:
            def _v(col):
                if col is None or len(row) <= col:
                    return ''
                return row[col]

            바코드 = 안전_문자(_v(바코드_col))
            if not 바코드:
                continue

            if 바코드 not in 집계:
                집계[바코드] = {
                    "중국재고_쿠팡": 0.0,
                    "중국재고_일반": 0.0,
                    "중국쿠팡_미입고": 0.0,
                    "중국일반_미입고": 0.0,
                    "입고예정_코드": [],
                    "입고일자": "",
                }

            # 중국재고: 마지막 비어있지 않은 값 사용 (빈 셀은 무시)
            if _v(쿠팡재고_col) != '':
                집계[바코드]["중국재고_쿠팡"] = 안전_숫자(_v(쿠팡재고_col))
            if _v(일반재고_col) != '':
                집계[바코드]["중국재고_일반"] = 안전_숫자(_v(일반재고_col))
            # 미입고량: 누적 합산
            집계[바코드]["중국쿠팡_미입고"] += 안전_숫자(_v(쿠팡발주수량_col))
            집계[바코드]["중국일반_미입고"] += 안전_숫자(_v(일반수량_col))
            # 발주코드: 여러 개면 콤마 구분
            코드 = 안전_문자(_v(발주번호_col))
            if 코드 and 코드 not in 집계[바코드]["입고예정_코드"]:
                집계[바코드]["입고예정_코드"].append(코드)
            # 입고일자: 가장 가까운 미래 값
            입고_날짜 = self._날짜_파싱(_v(입고일자_col))
            if 입고_날짜 and 입고_날짜 >= 오늘:
                기존 = self._날짜_파싱(집계[바코드]["입고일자"])
                if 기존 is None or 입고_날짜 < 기존:
                    집계[바코드]["입고일자"] = 입고_날짜.strftime("%Y-%m-%d")

        # 발주코드 리스트 → 문자열
        for v in 집계.values():
            v["입고예정_코드"] = ", ".join(v["입고예정_코드"]) or "-"
            if not v["입고일자"]:
                v["입고일자"] = "-"

        # 입고 그룹 포맷
        def _날짜_표시(날짜_str: str) -> str:
            try:
                d = datetime.strptime(날짜_str, "%Y-%m-%d")
                return f"{d.month}월 {d.day}일"
            except ValueError:
                return 날짜_str

        입고_그룹 = [
            {
                "날짜": 날짜,
                "날짜_표시": _날짜_표시(날짜),
                "개수": len(항목들),
                "항목들": 항목들,
            }
            for 날짜, 항목들 in sorted(입고_날짜_그룹.items())
        ]

        return 집계, 입고_그룹

    # ── 헬퍼 ────────────────────────────────────────────────────

    @staticmethod
    def _날짜_파싱(v) -> Optional[date]:
        if not v:
            return None
        if isinstance(v, (date, datetime)):
            return v.date() if isinstance(v, datetime) else v
        s = str(v).strip()
        for fmt in ("%Y-%m-%d", "%Y. %m. %d", "%Y.%m.%d", "%Y/%m/%d"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        return None

    @staticmethod
    def _재고_상태(발주량산출: float, 재고합계: float, 예상판매: float) -> str:
        if 예상판매 <= 0:
            return "정보없음"
        if 발주량산출 < 0:
            return "위험"
        if 발주량산출 < 예상판매:
            return "주의"
        return "정상"

    @staticmethod
    def _요약_계산(rows: list[dict]) -> dict:
        총계 = len(rows)
        위험 = sum(1 for r in rows if r["재고상태"] == "위험")
        주의 = sum(1 for r in rows if r["재고상태"] == "주의")
        정상 = sum(1 for r in rows if r["재고상태"] == "정상")
        return {
            "총제품수": 총계,
            "위험": 위험,
            "주의": 주의,
            "정상": 정상,
            "정보없음": 총계 - 위험 - 주의 - 정상,
        }
