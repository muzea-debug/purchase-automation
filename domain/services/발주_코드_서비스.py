"""
발주_코드_서비스 - PO 코드 파싱/포맷/다음번호 계산
"""
import re
import string
from pathlib import Path


class 발주_코드_서비스:

    @staticmethod
    def 파싱(코드: str) -> tuple[str, int]:
        """'NA(45)' → ('NA', 45). 형식 불일치 시 ValueError."""
        m = re.match(r'^([A-Za-z]+)\((\d+)\)$', 코드.strip())
        if not m:
            raise ValueError(f"잘못된 PO 코드 형식: {코드!r}")
        return m.group(1), int(m.group(2))

    @staticmethod
    def 포맷(prefix: str, 번호: int) -> str:
        """('NA', 45) → 'NA(45)'"""
        return f"{prefix}({번호})"

    @staticmethod
    def 유효성_검사(코드: str) -> bool:
        return bool(re.match(r'^[A-Za-z]+\(\d+\)$', 코드.strip()))

    @staticmethod
    def 최대번호_로컬(발주서_디렉토리: Path, prefix: str) -> int:
        """발주서 폴더에서 prefix 최대 사용 번호 반환 (없으면 0)."""
        max_num = 0
        pattern = re.compile(rf'^{re.escape(prefix)}\((\d+)\)\.xlsx$')
        for sub in ("정기발주", "신기종"):
            folder = 발주서_디렉토리 / sub
            if not folder.exists():
                continue
            for f in folder.glob("*.xlsx"):
                m = pattern.match(f.name)
                if m:
                    max_num = max(max_num, int(m.group(1)))
        return max_num

    @staticmethod
    def 다음_번호_로컬(발주서_디렉토리: Path, prefix: str) -> int:
        return 발주_코드_서비스.최대번호_로컬(발주서_디렉토리, prefix) + 1

    @staticmethod
    def 현재_활성_prefix(base: str, 최대번호_조회: "callable[[str], int]",
                         상한: int = 999) -> tuple[str, int]:
        """
        base+A 부터 base+Z 순으로 조회해 상한 미만인 첫 prefix 반환.
        반환: (active_prefix, current_max)
        모두 소진 시 (base+'Z', 상한) 반환 → 호출자가 경보 처리.
        """
        for letter in string.ascii_uppercase:
            prefix = base + letter
            max_num = 최대번호_조회(prefix)
            if max_num < 상한:
                return prefix, max_num
        return base + "Z", 상한

    @staticmethod
    def 브랜드_po_base(브랜드: str, 카테고리1: str = "") -> str:
        """브랜드 + 카테고리로 올바른 base prefix 결정."""
        import config
        cat = (카테고리1 or "").strip()
        if cat in config.전장품_카테고리:
            base = config.BRAND_전장품_PO_BASE.get(브랜드)
            if base:
                return base
        base = config.BRAND_PO_BASE.get(브랜드)
        if base:
            return base
        # fallback: 브랜드 앞 2자 대문자
        return (브랜드[:2].upper() if len(브랜드) >= 2 else 브랜드.upper())
