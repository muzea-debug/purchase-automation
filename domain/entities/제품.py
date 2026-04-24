"""
제품 엔티티 - 발주서 한 행
"""
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class 제품:
    번호: int
    item: str               # 영문 ITEM명 (예: Travel Adapter)
    model_no: str           # MODEL NO.
    moq: str                # MOQ (예: 1000)
    이름: str               # 한글 상품명 (예: 누아트 여행용 멀티어댑터, 화이트)
    바코드: str
    원가: float             # 가격(위안)
    수출_usd: float = 0.0   # 수출가 USD
    사진_경로: Path | None = None   # 제품 사진 이미지
    수량: int = 0
    합계: float = 0.0
