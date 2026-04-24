"""
발주서 도메인 서비스 - 순수 비즈니스 규칙
"""
from ..entities.발주서 import 발주서
from ..entities.제품 import 제품


class 발주서_서비스:

    @staticmethod
    def 수출가_계산(원가: float, 환율: float = 195.0) -> float:
        """원가(위안) → 수출가(USD) 변환"""
        if 환율 <= 0:
            return 0.0
        return round(원가 / 환율 * 1000 / 13, 10)

    @staticmethod
    def 합계_계산(원가: float, 수량: int) -> float:
        return round(원가 * 수량, 2)

    @staticmethod
    def 유효성_검사(발주: 발주서) -> list[str]:
        """발주서 저장 전 유효성 체크. 오류 메시지 리스트 반환."""
        errors = []
        if not 발주.제조사.strip():
            errors.append("제조사명을 입력해주세요.")
        if not 발주.브랜드.strip():
            errors.append("브랜드를 입력해주세요.")
        if not 발주.제품명.strip():
            errors.append("제품명을 입력해주세요.")
        if not 발주.제품목록:
            errors.append("제품을 최소 1개 이상 추가해주세요.")
        for i, p in enumerate(발주.제품목록, 1):
            if not p.바코드.strip():
                errors.append(f"{i}번 제품: 바코드를 입력해주세요.")
            if not p.이름.strip():
                errors.append(f"{i}번 제품: 상품명을 입력해주세요.")
        return errors
