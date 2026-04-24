"""
PI 유스케이스 - PI 생성, 스케줄 관리, 대시보드 조회
"""
from __future__ import annotations
import json
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import config
from infrastructure.계좌_저장소 import 계좌저장소
from infrastructure.pi_생성기 import PI생성기
from infrastructure.슬랙_저장소 import pi_알림_발송


class PI유스케이스:

    def __init__(self, sheets=None):
        self._계좌 = 계좌저장소(sheets)
        self._생성기 = PI생성기()
        self._sheets = sheets
        self._데이터_파일 = config.PI_데이터_파일
        self._데이터_파일.parent.mkdir(parents=True, exist_ok=True)

    # ── PI 생성 ───────────────────────────────────────────────────

    def PI_생성(
        self,
        po코드: str,
        회사id: str,
        제품목록: list[dict],
        디포짓_목록: list[dict],
        담당자: str,
        담당자_영문: str = "",
        담당자이메일: str = "",
        중국담당자: str = "",
        중국이메일: str = "",
        중국연락처: str = "",
        max_lt: int = 0,
    ) -> dict:
        계좌 = self._계좌.단건_조회(회사id)
        if not 계좌:
            return {"success": False, "error": f"계좌 없음: {회사id}"}

        총액 = sum(int(p.get("qty", 0)) * float(p.get("unit_price", 0)) for p in 제품목록)
        for dep in 디포짓_목록:
            if not dep.get("amount") and dep.get("pct"):
                dep["amount"] = round(총액 * float(dep["pct"]), 2)
            if not dep.get("pct") and dep.get("amount") and 총액:
                dep["pct"] = round(float(dep["amount"]) / 총액, 4)

        # 현재 납부 차수까지만 표시
        표시_목록 = self._현재_표시_디포짓(디포짓_목록)

        _담당자_영문 = 담당자_영문 or 담당자
        파일 = self._생성기.생성(
            po코드, 계좌, 제품목록, 표시_목록, _담당자_영문, 중국담당자,
            담당자이메일=담당자이메일,
            중국이메일=중국이메일,
            중국연락처=중국연락처,
            max_lt=max_lt,
        )

        db = self._db_읽기()
        db[po코드] = {
            "po코드":       po코드,
            "담당자":       담당자,
            "담당자_영문":  _담당자_영문,
            "담당자이메일": 담당자이메일,
            "중국담당자":   중국담당자,
            "중국이메일":   중국이메일,
            "중국연락처":   중국연락처,
            "회사id":       회사id,
            "파일경로":     str(파일),
            "총액_위안":    총액,
            "제품목록":     제품목록,
            "디포짓":       [dict(d, 완료=False) for d in 디포짓_목록],
            "생성일":       date.today().isoformat(),
            "max_lt":       max_lt,
        }
        self._db_저장(db)

        return {"success": True, "파일경로": str(파일), "po코드": po코드}

    # ── PI 수정 ───────────────────────────────────────────────────

    def PI_디포짓_수정(self, po코드: str, 디포짓_목록: list[dict]) -> dict:
        db = self._db_읽기()
        if po코드 not in db:
            return {"success": False, "error": "PI 없음"}
        meta = db[po코드]
        총액 = float(meta["총액_위안"])
        for dep in 디포짓_목록:
            if not dep.get("amount") and dep.get("pct"):
                dep["amount"] = round(총액 * float(dep["pct"]), 2)
            if not dep.get("pct") and dep.get("amount") and 총액:
                dep["pct"] = round(float(dep["amount"]) / 총액, 4)

        계좌 = self._계좌.단건_조회(meta["회사id"])
        표시_목록 = self._현재_표시_디포짓(디포짓_목록)
        파일 = self._생성기.생성(
            po코드, 계좌, meta["제품목록"], 표시_목록,
            meta.get("담당자_영문") or meta["담당자"],
            담당자이메일=meta.get("담당자이메일", ""),
            중국이메일=meta.get("중국이메일", ""),
            중국연락처=meta.get("중국연락처", ""),
            max_lt=int(meta.get("max_lt", 0)),
        )
        meta["디포짓"] = [dict(d, 완료=d.get("완료", False)) for d in 디포짓_목록]
        meta["파일경로"] = str(파일)
        self._db_저장(db)
        return {"success": True, "파일경로": str(파일)}

    def PI_삭제(self, po코드: str) -> dict:
        db = self._db_읽기()
        if po코드 not in db:
            return {"success": False, "error": "PI 없음"}
        del db[po코드]
        self._db_저장(db)
        return {"success": True}

    def PI_완료_처리(self, po코드: str, 디포짓_인덱스: int) -> dict:
        """특정 디포짓 차수 완료 처리 후 xlsx 재생성 (다음 차수 기준)."""
        db = self._db_읽기()
        if po코드 not in db:
            return {"success": False, "error": "PI 없음"}
        meta = db[po코드]
        deps = meta["디포짓"]
        if 디포짓_인덱스 >= len(deps):
            return {"success": False, "error": "인덱스 범위 초과"}

        deps[디포짓_인덱스]["완료"] = True

        # xlsx 재생성: 다음 미완료 차수가 현재 납부 차수로 강조됨
        try:
            계좌 = self._계좌.단건_조회(meta["회사id"])
            표시_목록 = self._현재_표시_디포짓(deps)
            파일 = self._생성기.생성(
                po코드, 계좌, meta["제품목록"], 표시_목록,
                meta.get("담당자_영문") or meta["담당자"],
                담당자이메일=meta.get("담당자이메일", ""),
                중국이메일=meta.get("중국이메일", ""),
                중국연락처=meta.get("중국연락처", ""),
                max_lt=int(meta.get("max_lt", 0)),
            )
            meta["파일경로"] = str(파일)
        except Exception as e:
            print(f"[PI유스케이스] 완료처리 후 재생성 실패: {e}")

        self._db_저장(db)
        return {"success": True}

    # ── 대시보드 조회 ─────────────────────────────────────────────

    def 담당자_PI_목록(self, 담당자: str, sheets=None) -> list[dict]:
        db = self._db_읽기()
        결과 = []
        입고맵 = self._입고일자_조회(sheets or self._sheets)

        for meta in db.values():
            if meta.get("담당자") != 담당자:
                continue
            미완료_deps = [d for d in meta.get("디포짓", []) if not d.get("완료")]
            if not 미완료_deps:
                continue

            입고상태 = self._입고상태_판단(meta.get("po코드", ""), 입고맵)
            결과.append({
                "po코드":    meta["po코드"],
                "담당자":    meta["담당자"],
                "총액_위안": meta["총액_위안"],
                "생성일":    meta["생성일"],
                "파일경로":  meta.get("파일경로", ""),
                "입고상태":  입고상태,
                "제품목록":  meta["제품목록"],
                "디포짓":    meta.get("디포짓", []),
            })

        return sorted(결과, key=lambda x: x["생성일"], reverse=True)

    def 전체_PI_목록(self) -> list[dict]:
        db = self._db_읽기()
        return [
            {
                "po코드":      m["po코드"],
                "담당자":      m["담당자"],
                "총액_위안":   m["총액_위안"],
                "생성일":      m["생성일"],
                "파일경로":    m.get("파일경로", ""),
                "디포짓_요약": [
                    {
                        "label":  d["label"],
                        "pct":    d.get("pct", 0),
                        "amount": d.get("amount", 0),
                        "예정일": d.get("예정일", ""),
                        "완료":   d.get("완료", False),
                    }
                    for d in m.get("디포짓", [])
                ],
            }
            for m in db.values()
        ]

    # ── 슬랙 스케줄러 ─────────────────────────────────────────────

    def 오늘_알림_발송(self) -> int:
        today = date.today().isoformat()
        db = self._db_읽기()
        count = 0
        for meta in db.values():
            for dep in meta.get("디포짓", []):
                if dep.get("완료"):
                    continue
                if dep.get("예정일", "") == today:
                    ok = pi_알림_발송(
                        meta["담당자"],
                        meta["po코드"],
                        dep["label"],
                        float(dep.get("pct", 0)),
                        float(dep.get("amount", 0)),
                        meta.get("제품목록", []),
                    )
                    if ok:
                        count += 1
        return count

    # ── 내부 헬퍼 ─────────────────────────────────────────────────

    @staticmethod
    def _현재_표시_디포짓(디포짓_목록: list[dict]) -> list[dict]:
        """완료 차수 + 첫 번째 미완료 차수까지만 반환.
        미완료 항목에 _현재=True 플래그 추가 → 생성기에서 강조 처리."""
        결과 = []
        for d in 디포짓_목록:
            is_current = not d.get("완료", False)
            결과.append({**d, "_현재": is_current})
            if is_current:
                break  # 첫 미완료에서 멈춤 (미래 차수 제외)
        return 결과

    def _db_읽기(self) -> dict:
        if not self._데이터_파일.exists():
            return {}
        return json.loads(self._데이터_파일.read_text(encoding="utf-8"))

    def _db_저장(self, db: dict) -> None:
        self._데이터_파일.write_text(
            json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def _입고일자_조회(self, sheets) -> dict[str, str]:
        if not sheets:
            return {}
        try:
            result = sheets._service.spreadsheets().values().get(
                spreadsheetId=sheets._스프레드시트_id,
                range=f"'{sheets._시트명}'!B:K",
            ).execute()
            rows = result.get("values", [])
            맵 = {}
            for row in rows[1:]:
                if len(row) < 3:
                    continue
                po = str(row[1]).strip() if len(row) > 1 else ""
                입고일 = str(row[9]).strip() if len(row) > 9 else ""
                if po and 입고일:
                    맵[po] = 입고일
            return 맵
        except Exception as e:
            print(f"[PI유스케이스] 입고일자 조회 실패: {e}")
            return {}

    def _입고상태_판단(self, po코드: str, 입고맵: dict) -> str:
        입고일 = 입고맵.get(po코드, "")
        if not 입고일:
            return "입고 미정"
        try:
            clean = 입고일.replace(". ", "-").replace(".", "").strip()
            d = datetime.strptime(clean, "%Y-%m-%d").date()
            if d <= date.today():
                return f"입고 완료 ({입고일})"
            return f"입고 예정 {입고일}"
        except ValueError:
            return f"입고 예정 {입고일}"
