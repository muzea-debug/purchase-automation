"""
계좌 DB 저장소 - 제조사별 계좌정보 CRUD + 구글시트 숨김탭 백업
"""
from __future__ import annotations
import json
import re
import uuid
from pathlib import Path
from typing import Optional

import config


class 계좌저장소:

    _SHEETS_탭 = "계좌DB_백업"

    def __init__(self, sheets=None):
        self._파일 = config.계좌DB_파일
        self._sheets = sheets
        self._파일.parent.mkdir(parents=True, exist_ok=True)
        config.업체도장_DIR.mkdir(parents=True, exist_ok=True)

    # ── CRUD ──────────────────────────────────────────────────────

    def 전체_조회(self) -> dict:
        if not self._파일.exists():
            return {}
        return json.loads(self._파일.read_text(encoding="utf-8"))

    def 단건_조회(self, 회사id: str) -> Optional[dict]:
        return self.전체_조회().get(회사id)

    def 저장(self, 계좌: dict) -> str:
        """계좌 추가/수정. 반환: 회사id"""
        db = self.전체_조회()
        회사id = 계좌.get("id") or self._id_생성(계좌.get("회사명_영문", ""))
        계좌["id"] = 회사id
        db[회사id] = 계좌
        self._파일.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
        self._구글시트_백업(db)
        return 회사id

    def 삭제(self, 회사id: str) -> bool:
        db = self.전체_조회()
        if 회사id not in db:
            return False
        도장 = db[회사id].get("도장_파일", "")
        if 도장:
            p = config.BASE_DIR / "assets" / 도장
            if p.exists():
                p.unlink()
        del db[회사id]
        self._파일.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
        self._구글시트_백업(db)
        return True

    def 도장_저장(self, 회사id: str, 이미지_바이트: bytes, 확장자: str = ".png") -> str:
        """업체 도장 이미지 저장. 반환: 상대경로"""
        경로 = config.업체도장_DIR / f"{회사id}{확장자}"
        경로.write_bytes(이미지_바이트)
        상대경로 = f"업체도장/{회사id}{확장자}"
        db = self.전체_조회()
        if 회사id in db:
            db[회사id]["도장_파일"] = 상대경로
            self._파일.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
        return 상대경로

    # ── 텍스트 파싱 ───────────────────────────────────────────────

    @staticmethod
    def 텍스트_파싱(text: str) -> dict:
        """PI 하단 Bank information 블록 텍스트에서 계좌 정보 추출."""
        def _find(pattern: str) -> str:
            m = re.search(pattern, text, re.IGNORECASE)
            return m.group(1).strip() if m else ""

        결과 = {
            "account_number":      _find(r"Beneficiary (?:Account )?Number\s*:\s*(.+)"),
            "beneficiary_name":    _find(r"Beneficiary Name\s*:\s*(.+)"),
            "beneficiary_address": _find(r"Beneficiary Address\s*:\s*(.+)"),
            "swift":               _find(r"SWIFT/BIC\s*(?:code)?\s*:\s*(\S+)"),
            "bank_name":           _find(r"Bank name\s*:\s*(.+)"),
            "bank_address":        _find(r"Bank Address\s*:\s*(.+)"),
            "sort_code":           _find(r"Sort code\s*:\s*(\S+)"),
            "branch_code":         _find(r"Branch Code\s*:\s*(\S+)"),
        }
        cc = _find(r"County\s*/\s*City\s*:\s*(.+)")
        if "/" in cc:
            parts = [p.strip() for p in cc.split("/", 1)]
            결과["country"] = parts[0]
            결과["city"] = parts[1]
        결과["회사명_영문"] = 결과.get("beneficiary_name", "")
        return 결과


    # ── 구글시트 백업 ─────────────────────────────────────────────

    def _구글시트_백업(self, db: dict) -> None:
        if not self._sheets:
            return
        try:
            rows = [["id","회사명_영문","회사명_중문","country","city",
                     "account_number","beneficiary_name","beneficiary_address",
                     "swift","bank_name","bank_address","sort_code","branch_code",
                     "seller_name","seller_email","seller_phone","도장_파일"]]
            for 계좌 in db.values():
                rows.append([
                    계좌.get("id",""), 계좌.get("회사명_영문",""), 계좌.get("회사명_중문",""),
                    계좌.get("country",""), 계좌.get("city",""), 계좌.get("account_number",""),
                    계좌.get("beneficiary_name",""), 계좌.get("beneficiary_address",""),
                    계좌.get("swift",""), 계좌.get("bank_name",""), 계좌.get("bank_address",""),
                    계좌.get("sort_code",""), 계좌.get("branch_code",""),
                    계좌.get("seller_name",""), 계좌.get("seller_email",""), 계좌.get("seller_phone",""),
                    계좌.get("도장_파일",""),
                ])
            self._sheets._구글시트_탭_덮어쓰기(self._SHEETS_탭, rows)
        except Exception as e:
            print(f"[계좌저장소] 구글시트 백업 실패: {e}")

    @staticmethod
    def _id_생성(회사명: str) -> str:
        slug = re.sub(r'[^a-zA-Z0-9]+', '-', 회사명.lower()).strip('-')
        return slug[:30] or str(uuid.uuid4())[:8]
