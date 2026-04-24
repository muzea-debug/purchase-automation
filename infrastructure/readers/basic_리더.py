"""
basic_리더 - basic_operation_rocket_*.csv → 바코드별 쿠팡FC 현재재고수량

헤더: Row 1 (UTF-8-BOM)
식별: 파일명 'basic' 으로 시작
컬럼: '바코드', '현재재고수량', '센터'
집계 대상 센터: FC, RC
"""
import csv
from pathlib import Path
from ._공통 import 컬럼_찾기, 안전_숫자, 안전_문자

_집계_센터 = {"FC", "RC"}


class Basic_리더:

    def 읽기(self, 파일: Path) -> dict[str, float]:
        """바코드 → 쿠팡FC 현재재고 dict 반환. FC·RC 센터만 합산."""
        with open(파일, encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            헤더 = next(reader)

            바코드_col = 컬럼_찾기(헤더, ["바코드"])
            재고_col   = 컬럼_찾기(헤더, ["현재재고수량"])
            센터_col   = 컬럼_찾기(헤더, ["센터"])

            if 바코드_col is None or 재고_col is None:
                raise ValueError("basic 파일에서 '바코드' 또는 '현재재고수량' 컬럼을 찾을 수 없습니다.")

            결과: dict[str, float] = {}
            for row in reader:
                if len(row) <= max(바코드_col, 재고_col):
                    continue
                # 센터 컬럼이 있으면 FC·RC만 통과
                if 센터_col is not None and len(row) > 센터_col:
                    if row[센터_col].strip() not in _집계_센터:
                        continue
                바코드 = 안전_문자(row[바코드_col])
                if not 바코드 or 바코드 == '0':
                    continue
                재고 = 안전_숫자(row[재고_col])
                결과[바코드] = 결과.get(바코드, 0.0) + 재고

        return 결과
