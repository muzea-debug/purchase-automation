"""
카탈로그 상품 정보 - 통합발주서원본에서 읽은 단일 제품
"""
from dataclasses import dataclass
from pathlib import Path


@dataclass
class 상품_정보:
    바코드: str
    이름: str           # 한글 상품명
    브랜드: str         # NUART / ILYON 등
    제조사: str         # 富坤电器 등
    원가: float         # 원가(CNY)
    수출_usd: float     # 수출가(USD)
    moq: str
    item: str           # 영문 ITEM명
    model_no: str
    번호: int           # 발주서 내 순번
    원본파일: Path      # 이 상품이 속한 발주서 원본 경로
    계좌id: str = ""   # PI 계좌 연동 id (J5 셀)
