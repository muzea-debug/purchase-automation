"""
리더 공통 유틸리티
"""


def 컬럼_찾기(헤더_행: list, 키워드들: list[str], require_all: bool = False) -> int | None:
    """헤더 행에서 키워드로 컬럼 인덱스 반환 (0-indexed). 없으면 None."""
    for i, h in enumerate(헤더_행):
        if h is None:
            continue
        h_str = str(h).lower().replace('\n', ' ').replace('\r', ' ')
        if require_all:
            if all(kw.lower() in h_str for kw in 키워드들):
                return i
        else:
            if any(kw.lower() in h_str for kw in 키워드들):
                return i
    return None


def 안전_숫자(v) -> float:
    try:
        if v is None or v == '':
            return 0.0
        return float(str(v).replace(',', '').strip())
    except (ValueError, TypeError):
        return 0.0


def 안전_문자(v) -> str:
    if v is None:
        return ''
    return str(v).strip()
