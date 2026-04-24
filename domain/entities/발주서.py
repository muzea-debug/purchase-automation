"""
발주서 엔티티 - 원본 발주서 한 장
"""
from dataclasses import dataclass, field
from pathlib import Path
from .제품 import 제품


@dataclass
class 발주서:
    제조사: str
    브랜드: str             # NUART / ILYON 등
    제품명: str             # 발주서 대표 제품명
    패키지원가: float
    부자재원가: float
    리드타임: str           # 예: 25-30일
    시즌구분: str = ""      # 카테고리DB용: 시즌구분
    카테고리1: str = ""     # 카테고리DB용: 대분류
    카테고리2: str = ""     # 카테고리DB용: 소분류
    한국_요청_도착일: str = ""
    중국지사_도착일: str = ""
    제품목록: list[제품] = field(default_factory=list)
    포장_이미지들: list[dict] = field(default_factory=list)  # [{"이름": "로고", "경로": Path}, ...]
    외박스_텍스트: str = ""        # ILYON 전용: 외박스 표기 문구
    외박스_이미지: object = None   # ILYON 전용: 외박스 기본 이미지 경로
    계좌id: str = ""              # PI 계좌 연동 (계좌DB의 id)

    @property
    def 파일명(self) -> str:
        return f"{self.제조사}_{self.브랜드} {self.제품명}({self.브랜드}전용).xlsx"
