"""
발주서 생성 유스케이스 - 흐름 조율
"""
from pathlib import Path
from werkzeug.datastructures import ImmutableMultiDict, FileStorage

from domain.entities.발주서 import 발주서
from domain.entities.제품 import 제품
from domain.services.발주서_서비스 import 발주서_서비스
from infrastructure.xlsx_저장소 import Xlsx저장소
from infrastructure.이미지_저장소 import 이미지_저장소
import config


class 발주서_생성_유스케이스:

    def __init__(self, 저장_경로: Path, upload_dir: Path, sheets_저장소=None):
        self.저장_경로 = 저장_경로
        self.xlsx_저장소 = Xlsx저장소()
        self.이미지_저장소 = 이미지_저장소(upload_dir)
        self.서비스 = 발주서_서비스()
        self._sheets = sheets_저장소

    def 실행(self, form: ImmutableMultiDict, files: dict) -> dict:
        """
        Returns:
            {"success": True, "파일명": "...", "경로": "..."}
            {"success": False, "errors": [...]}
        """
        # ── 1. 이미지 저장 ──────────────────────────────────────
        제품_사진들 = {}
        for key, file in files.items():
            if key.startswith("제품_사진_") and file.filename:
                idx = int(key.split("_")[-1])
                제품_사진들[idx] = self.이미지_저장소.저장(file)

        포장_이미지들 = []
        idx = 1
        while True:
            # 구분명이 없으면 해당 인덱스 이후는 없는 것으로 간주
            이름 = form.get(f"포장_구분_{idx}", "").strip()
            텍스트 = form.get(f"포장_텍스트_{idx}", "").strip()
            file_list = files.getlist(f"포장_이미지_{idx}")

            # 구분명도 없고 파일도 없으면 중단
            if not 이름 and not file_list:
                break

            path = None
            if file_list:
                file = file_list[0]
                if file.filename:
                    path = self.이미지_저장소.저장(file)

            # 구분명 또는 이미지 중 하나라도 있으면 항목 추가
            if 이름 or path:
                포장_이미지들.append({"이름": 이름, "경로": path, "텍스트": 텍스트})
            idx += 1

        # ── 2. 제품 목록 파싱 ───────────────────────────────────
        제품목록 = []
        idx = 1
        while f"item_{idx}" in form:
            p = 제품(
                번호=idx,
                item=form.get(f"item_{idx}", ""),
                model_no=form.get(f"model_no_{idx}", ""),
                moq=form.get(f"moq_{idx}", ""),
                이름=form.get(f"이름_{idx}", ""),
                바코드=form.get(f"바코드_{idx}", ""),
                원가=float(form.get(f"원가_{idx}", 0) or 0),
                수출_usd=float(form.get(f"수출_usd_{idx}", 0) or 0),
                사진_경로=제품_사진들.get(idx),
            )
            제품목록.append(p)
            idx += 1

        # ── 3. 외박스 이미지 저장 (ILYON 전용) ─────────────────────
        # 업로드하면 그것 사용, 없으면 기본 이미지 자동 적용
        외박스_이미지_경로 = None
        외박스_파일 = files.get("외박스_이미지")
        if 외박스_파일 and 외박스_파일.filename:
            외박스_이미지_경로 = self.이미지_저장소.저장(외박스_파일)
        elif form.get("브랜드") in ("ILYON", "일리온"):
            if config.ILYON_외박스_기본이미지.exists():
                외박스_이미지_경로 = config.ILYON_외박스_기본이미지
        elif form.get("브랜드") == "누아트스튜디오":
            if config.누아트스튜디오_외박스_기본이미지.exists():
                외박스_이미지_경로 = config.누아트스튜디오_외박스_기본이미지

        # ── 4. 발주서 엔티티 생성 ────────────────────────────────
        발주 = 발주서(
            제조사=form.get("제조사", ""),
            브랜드=form.get("브랜드", ""),
            제품명=form.get("제품명", ""),
            패키지원가=float(form.get("패키지원가", 0) or 0),
            부자재원가=float(form.get("부자재원가", 0) or 0),
            리드타임=form.get("리드타임", ""),
            시즌구분=form.get("시즌구분", ""),
            카테고리1=form.get("카테고리1", ""),
            카테고리2=form.get("카테고리2", ""),
            한국_요청_도착일=form.get("한국_요청_도착일", ""),
            중국지사_도착일=form.get("중국지사_도착일", ""),
            제품목록=제품목록,
            포장_이미지들=포장_이미지들,
            외박스_텍스트=form.get("외박스_텍스트", "").strip(),
            외박스_이미지=외박스_이미지_경로,
            계좌id=form.get("계좌id", "").strip(),
        )

        # ── 5. 유효성 검사 ───────────────────────────────────────
        errors = self.서비스.유효성_검사(발주)
        if errors:
            return {"success": False, "errors": errors}

        # ── 6. xlsx 메모리 생성 ──────────────────────────────────
        버퍼, 파일명 = self.xlsx_저장소.저장_메모리(발주)

        # ── 7. 카테고리DB 업데이트 ───────────────────────────────
        if self._sheets and hasattr(self._sheets, "카테고리DB_행추가"):
            try:
                from datetime import date
                오늘 = date.today()
                등록일 = f"{오늘.year}. {오늘.month}. {오늘.day}"

                db_목록 = []
                for p in 발주.제품목록:
                    원가 = p.원가
                    최종원가 = round(원가 + 발주.패키지원가 + 발주.부자재원가, 4)
                    db_목록.append({
                        "브랜드":    발주.브랜드,
                        "시즌구분":  발주.시즌구분,
                        "카테고리1": 발주.카테고리1,
                        "카테고리2": 발주.카테고리2,
                        "바코드":    p.바코드,
                        "상품명":    p.이름,
                        "발주서명":  발주.제품명,
                        "moq":       p.moq,
                        "최종원가":  최종원가,
                        "제품원가":  원가,
                        "패키지원가": 발주.패키지원가,
                        "부자재원가": 발주.부자재원가,
                        "리드타임":  발주.리드타임,
                        "등록일":    등록일,
                    })
                self._sheets.카테고리DB_행추가(db_목록)
            except Exception as e:
                print(f"[유스케이스] 카테고리DB 업데이트 실패: {e}")

        return {"success": True, "버퍼": 버퍼, "파일명": 파일명}
