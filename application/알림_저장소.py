"""
알림_저장소 - 원본 발주서 생성 시 발주 실행 페이지로 전달할 알림 공유 상태
"""
from __future__ import annotations
import threading

_lock = threading.Lock()
_알림_목록: list[dict] = []   # [{"바코드": "...", "이름": "..."}, ...]


def 알림_추가(items: list[dict]) -> None:
    """원본 발주서 생성 라우트에서 호출. items: [{"바코드":..., "이름":...}, ...]"""
    if not items:
        return
    with _lock:
        _알림_목록.extend(items)


def 알림_읽기_후_삭제() -> list[dict]:
    """발주 실행 페이지 폴링 API에서 호출. 읽으면 큐를 비움."""
    with _lock:
        결과 = _알림_목록.copy()
        _알림_목록.clear()
    return 결과
