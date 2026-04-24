"""
po_리더 - PO_SKU_LIST_*.csv → 바코드별 쿠팡 발주수량 합산

헤더: Row 1 (UTF-8-BOM)
식별: 파일명 'PO_SKU_LIST' 포함
복수 파일 지원: 모두 합산
"""
import csv
from pathlib import Path
from ._공통 import 컬럼_찾기, 안전_숫자, 안전_문자


class PO_리더:

    def 읽기(self, 파일들: list[Path]) -> dict[str, float]:
        """바코드 → 발주수량 합계 dict 반환 (복수 파일 통합)."""
        결과: dict[str, float] = {}
        for 파일 in 파일들:
            if 파일.stat().st_size < 10:
                continue  # 빈 파일 스킵
            self._단일_파일_읽기(파일, 결과)
        return 결과

    @staticmethod
    def _단일_파일_읽기(파일: Path, 누적: dict) -> None:
        with open(파일, encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            헤더 = next(reader)

            바코드_col   = 컬럼_찾기(헤더, ["sku barcode", "skubarcode", "바코드"])
            발주수량_col  = 컬럼_찾기(헤더, ["발주수량"])
            발주현황_col  = 컬럼_찾기(헤더, ["발주현황"])

            if 바코드_col is None or 발주수량_col is None:
                return

            for row in reader:
                if len(row) <= max(바코드_col, 발주수량_col):
                    continue
                # 발주확정 행 제외 — 거래처확인요청 등 미확정 행만 집계
                if 발주현황_col is not None and len(row) > 발주현황_col:
                    if 안전_문자(row[발주현황_col]) == "발주확정":
                        continue
                바코드 = 안전_문자(row[바코드_col])
                if not 바코드:
                    continue
                수량 = 안전_숫자(row[발주수량_col])
                누적[바코드] = 누적.get(바코드, 0.0) + 수량
