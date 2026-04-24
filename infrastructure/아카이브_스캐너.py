"""
아카이브 스캐너 - 재고현황_아카이브/YYYY.MM.DD/ 폴더 스캔
각 날짜 폴더에서 소스 파일 타입 자동 매핑
"""
from __future__ import annotations
import re
from pathlib import Path
from datetime import datetime

import config


_날짜_패턴 = re.compile(r"^(\d{4})\.(\d{2})\.(\d{2})$")


class 아카이브_스캐너:

    def __init__(self):
        self._dir = config.재고현황_아카이브_DIR

    def 날짜_목록(self) -> list[dict]:
        """사용 가능한 날짜 폴더 목록 반환. 최신 순 정렬."""
        if not self._dir.exists():
            return []
        결과 = []
        for 폴더 in sorted(self._dir.iterdir(), reverse=True):
            if not 폴더.is_dir():
                continue
            m = _날짜_패턴.match(폴더.name)
            if not m:
                continue
            날짜 = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            파일맵 = self._파일_매핑(폴더)
            결과.append({
                "폴더명": 폴더.name,
                "날짜": 날짜,
                "파일맵": {k: str(v) for k, v in 파일맵.items()},
                "소스_수": len(파일맵),
            })
        return 결과

    def 날짜_파일맵(self, 폴더명: str) -> dict[str, Path]:
        """특정 날짜 폴더의 파일 타입 → Path 맵 반환."""
        폴더 = self._dir / 폴더명
        if not 폴더.exists():
            raise FileNotFoundError(f"아카이브 폴더 없음: {폴더}")
        return self._파일_매핑(폴더)

    def _파일_매핑(self, 폴더: Path) -> dict[str, Path]:
        맵: dict[str, Path] = {}
        po_목록: list[Path] = []

        for f in 폴더.iterdir():
            if f.name == "desktop.ini" or f.name.startswith("~$"):
                continue
            name = f.name.lower()

            if "엔대시_구글시트" in f.name or "엔대시" in f.name and f.suffix == ".xlsx":
                맵["엔대시"] = f

            elif name.startswith("basic_operation_rocket") and f.suffix == ".csv":
                맵["basic"] = f

            elif (name.startswith("a00412786") or "forecast" in name) and f.suffix == ".xlsx":
                맵["forecast"] = f

            elif "nuartcompany" in name and f.suffix == ".xlsx":
                맵["누아트sku"] = f

            elif name.startswith("po_sku_list") and f.suffix == ".csv":
                po_목록.append(f)

            elif ("사방넷" in f.name or "사방넷단품" in f.name) and f.suffix == ".xlsx":
                맵["사방넷"] = f

            elif "주문서확인처리" in f.name and f.suffix == ".xlsx":
                맵["주문서"] = f

        if po_목록:
            # 가장 최신(이름 기준 마지막) PO 파일 선택
            맵["po"] = sorted(po_목록)[-1]
            if len(po_목록) > 1:
                맵["po_목록"] = po_목록  # type: ignore

        return 맵
