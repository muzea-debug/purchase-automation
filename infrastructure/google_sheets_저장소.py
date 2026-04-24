"""
Google Sheets 저장소 - 엔대시 시트 PO번호 조회 + 발주 이력 기록 + 발주요청 읽기
"""
import re
from datetime import date, datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build

_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class GoogleSheets저장소:

    def __init__(self, 서비스계정: str, 스프레드시트_id: str, 시트명: str):
        self._스프레드시트_id = 스프레드시트_id
        self._시트명 = 시트명
        # JSON 문자열이면 from_service_account_info, 파일 경로면 from_service_account_file
        import json as _json
        try:
            info = _json.loads(서비스계정)
            creds = service_account.Credentials.from_service_account_info(info, scopes=_SCOPES)
        except (_json.JSONDecodeError, ValueError):
            creds = service_account.Credentials.from_service_account_file(서비스계정, scopes=_SCOPES)
        self._service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    def 마지막_po번호_조회(self, prefix: str) -> int:
        """
        C열(발주번호)에서 prefix(xxx) 패턴 중 최대 번호 반환.
        없거나 오류 시 0 반환.
        """
        try:
            result = (
                self._service.spreadsheets()
                .values()
                .get(
                    spreadsheetId=self._스프레드시트_id,
                    range=f"{self._시트명}!C:C",
                )
                .execute()
            )
            rows = result.get("values", [])
        except Exception as e:
            print(f"[GoogleSheets] 조회 실패: {e}")
            return 0

        pattern = re.compile(rf'^{re.escape(prefix)}\((\d+)\)$')
        max_num = 0
        for row in rows:
            if row:
                m = pattern.match(str(row[0]).strip())
                if m:
                    max_num = max(max_num, int(m.group(1)))
        return max_num

    def 발주_기록(
        self,
        바코드: str,
        po코드: str,
        수량: int,
        제품명: str,
        제조사: str,
    ) -> None:
        """
        엔대시 시트에 한 행 추가.
        컬럼 순서: 발주일자 | 바코드 | 발주번호 | 쿠팡수량 | 일반수량 | 제품명 | 비고 | 중국출발 | 누아트입고 | 14일경과 | 제조사
        """
        오늘 = date.today()
        발주일자 = f"{오늘.year}. {오늘.month}. {오늘.day}"
        row = [발주일자, 바코드, po코드, 수량, "", 제품명, "", "", "", "", 제조사]
        self._service.spreadsheets().values().append(
            spreadsheetId=self._스프레드시트_id,
            range=f"{self._시트명}!A:K",
            valueInputOption="USER_ENTERED",
            body={"values": [row]},
        ).execute()

    # ── 발주요청 읽기 (재고현황 자동화용) ─────────────────────

    def 발주요청_읽기(self, 시트명: str) -> tuple[list[str], list[list]]:
        """
        발주요청 시트 전체 읽기.
        Returns: (헤더_리스트, 데이터_행_리스트)
        """
        result = (
            self._service.spreadsheets()
            .values()
            .get(
                spreadsheetId=self._스프레드시트_id,
                range=f"{시트명}!A:T",
            )
            .execute()
        )
        values = result.get("values", [])
        if not values:
            return [], []
        헤더 = values[0]
        데이터 = [r for r in values[1:] if any(v for v in r)]
        return 헤더, 데이터

    def 제품DB_읽기(self) -> dict[str, dict]:
        """
        '카테고리 포함 db' 시트 → 바코드별 제품 메타데이터 dict.
        엑셀시트저장소.제품DB_읽기()와 동일한 반환 구조.
        """
        시트명 = "카테고리 포함 db"
        try:
            result = self._service.spreadsheets().values().get(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{시트명}'!A:AZ",
            ).execute()
            rows = result.get("values", [])
        except Exception as e:
            print(f"[GoogleSheets] 제품DB 읽기 실패: {e}")
            return {}

        if len(rows) < 2:
            return {}

        # 공백 보존(소문자만) — "최종원가" vs "최종 원가" 구분을 위해
        헤더 = [str(v).strip().lower().replace("\n", "") for v in rows[0]]

        def _col(*keywords):
            # 키워드 우선순위 순으로 검색
            for k in keywords:
                kn = k.lower()
                for i, h in enumerate(헤더):
                    if kn == h or kn in h:
                        return i
            return None

        바코드_c   = _col("바코드")
        브랜드_c   = _col("브랜드")
        발주서명_c = _col("발주서명")
        상품명_c   = _col("상품명")
        moq_c      = _col("moq")
        원가위안_c = _col("제품원가", "소싱원가", "원가")
        최종원가_c = _col("최종 원가", "최종원가")
        시즌_c     = _col("시즌구분", "시즌")
        대분류_c   = _col("카테고리_1", "대분류")
        소분류_c   = _col("카테고리_2", "소분류")
        아이템_c   = _col("아이템")
        lt_c       = _col("리드타임", "l/t", "lt", "기본lt")
        주의사항_c = _col("주의사항")

        if 바코드_c is None:
            print("[GoogleSheets] 제품DB: 바코드 컬럼 없음")
            return {}

        def _v(row, col, default=""):
            if col is None or col >= len(row):
                return default
            return row[col] if row[col] else default

        def _n(row, col):
            v = _v(row, col, 0)
            try:
                return float(str(v).replace(",", ""))
            except (TypeError, ValueError):
                return 0.0

        def _lt(row, col):
            """'20-25일', '10~15일', '30일' → 최대값 숫자 반환."""
            v = str(_v(row, col, "0"))
            nums = re.findall(r'\d+', v)
            if not nums:
                return 0.0
            return float(max(int(n) for n in nums))

        def _moq_max(row, col):
            """'15~20', '35-45', '100' → 최대값 숫자 반환."""
            v = str(_v(row, col, "0"))
            nums = re.findall(r'\d+', v)
            if not nums:
                return 0.0
            return float(max(int(n) for n in nums))

        결과 = {}
        for row in rows[1:]:
            바코드 = str(_v(row, 바코드_c, "")).strip()
            if not 바코드:
                continue
            결과[바코드] = {
                "바코드":    바코드,
                "브랜드":    str(_v(row, 브랜드_c)),
                "발주서명":  str(_v(row, 발주서명_c)),
                "상품명":    str(_v(row, 상품명_c)),
                "moq":           str(_v(row, moq_c)),
                "발주확정용_moq": _moq_max(row, moq_c),
                "원가_위안": _n(row, 원가위안_c),
                "최종원가":  _n(row, 최종원가_c),
                "시즌구분":  str(_v(row, 시즌_c)),
                "대분류":    str(_v(row, 대분류_c)),
                "소분류":    str(_v(row, 소분류_c)),
                "아이템":    str(_v(row, 아이템_c)),
                "안전일수":  0.0,
                "기본_lt":   _lt(row, lt_c),
                "주의사항":  str(_v(row, 주의사항_c)),
            }
        print(f"[GoogleSheets] 제품DB {len(결과)}개 로드")
        return 결과

    def 제품DB_수정(self, 바코드: str, 변경_필드: dict) -> bool:
        """Google Sheets 제품DB 시트에서 바코드 행의 특정 필드를 업데이트."""
        시트명 = "카테고리 포함 db"
        _필드_키워드 = {
            "moq":       ["moq"],
            "기본_lt":   ["리드타임", "l/t", "lt", "기본lt"],
            "원가_위안": ["제품원가", "소싱원가", "원가"],
            "최종원가":  ["최종 원가", "최종원가"],
            "발주서명":  ["발주서명"],
            "안전일수":  ["안전일수"],
            "주의사항":  ["주의사항"],
            "브랜드":    ["브랜드"],
        }
        try:
            result = self._service.spreadsheets().values().get(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{시트명}'!A:AZ",
            ).execute()
            rows = result.get("values", [])
            if len(rows) < 2:
                return False

            헤더 = [str(v).strip().lower().replace("\n", "") for v in rows[0]]

            def _find_col(*keywords):
                for kw in keywords:
                    kw_n = kw.lower()
                    for i, h in enumerate(헤더):
                        if kw_n == h or kw_n in h:
                            return i
                return None

            바코드_c = _find_col("바코드")
            if 바코드_c is None:
                return False
            col_map = {field: _find_col(*kws) for field, kws in _필드_키워드.items()}

            # 행 번호 찾기 (1-based for Sheets API)
            row_idx = None
            for i, row in enumerate(rows[1:], start=2):
                val = row[바코드_c] if 바코드_c < len(row) else ""
                if str(val).strip() == 바코드:
                    row_idx = i
                    break

            if row_idx is None:
                return False

            # batchUpdate로 각 셀 수정
            def col_letter(n):
                s = ""
                while n >= 0:
                    s = chr(65 + n % 26) + s
                    n = n // 26 - 1
                return s

            data = []
            for field, val in 변경_필드.items():
                c = col_map.get(field)
                if c is not None:
                    data.append({
                        "range": f"'{시트명}'!{col_letter(c)}{row_idx}",
                        "values": [[val]],
                    })

            if data:
                self._service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self._스프레드시트_id,
                    body={"valueInputOption": "USER_ENTERED", "data": data},
                ).execute()
            return True
        except Exception as e:
            print(f"[GoogleSheets] 제품DB 수정 실패: {e}")
            return False

    # ── 중국발주요청리스트 추가 ───────────────────────────────────

    _발주요청_시트 = "[엔대시]중국 발주요청리스트(交期管理表)"

    def 중국발주요청리스트_추가(self, 행들: list[dict]) -> None:
        """
        중국발주요청리스트 하단에 행 추가.
        행 dict 키: 바코드, po코드, 쿠팡수량, 일반수량, 제품명, 입고예정일
        컬럼 순서: A=쿠팡중국재고, B=일반중국재고, C=발주일자, D=바코드,
                   E=발주번호, F=쿠팡발주수량, G=일반건수량, H=제품명, I=비고/입고예정일
        """
        if not 행들:
            return
        오늘 = date.today()
        발주일자 = f"{오늘.year}. {오늘.month}. {오늘.day}"
        rows = []
        for h in 행들:
            po = str(h.get("po코드", ""))
            is_fa = po.upper().startswith("FA")
            쿠팡수량 = "" if is_fa else (h.get("쿠팡수량") or h.get("수량") or "")
            일반수량 = (h.get("일반수량") or h.get("수량") or "") if is_fa else ""
            rows.append([
                "",                          # A 쿠팡중국재고
                "",                          # B 일반중국재고
                발주일자,                     # C 발주일자
                str(h.get("바코드", "")),     # D 바코드
                po,                          # E 발주번호
                쿠팡수량,                     # F 쿠팡발주수량
                일반수량,                     # G 일반건수량
                str(h.get("제품명", "")),     # H 제품명
                "",                          # I 비고 (공란)
                "",                          # J 중국직원 사용
                "",                          # K 중국직원 사용
                "",                          # L 중국직원 사용
                str(h.get("제조사", "")),     # M 제조사
            ])
        self._service.spreadsheets().values().append(
            spreadsheetId=self._스프레드시트_id,
            range=f"'{self._발주요청_시트}'!A:M",
            valueInputOption="USER_ENTERED",
            body={"values": rows},
        ).execute()
        print(f"[GoogleSheets] 중국발주요청리스트 {len(rows)}행 추가 완료")

    # ── 카테고리DB 관련 ───────────────────────────────────────────

    _카테고리DB_시트 = "카테고리 포함 db"

    def 바코드_중복확인(self, 바코드: str) -> str | None:
        """
        카테고리 포함 db M열(바코드)에서 중복 확인.
        중복이면 I열(상품명) 반환, 없으면 None.
        """
        try:
            result = self._service.spreadsheets().values().get(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{self._카테고리DB_시트}'!M:M",
            ).execute()
            rows = result.get("values", [])
        except Exception as e:
            print(f"[GoogleSheets] 바코드 중복확인 실패: {e}")
            return None

        바코드 = str(바코드).strip()
        for i, row in enumerate(rows):
            if row and str(row[0]).strip() == 바코드:
                # 상품명(I열) 가져오기 — 같은 행 번호로 조회
                try:
                    r = self._service.spreadsheets().values().get(
                        spreadsheetId=self._스프레드시트_id,
                        range=f"'{self._카테고리DB_시트}'!I{i+1}",
                    ).execute()
                    v = r.get("values", [[""]])
                    return v[0][0] if v and v[0] else 바코드
                except Exception:
                    return 바코드
        return None

    def _다음_아이템코드들(self, 개수: int) -> list[str]:
        """HA-1 ~ HA-9999 → HB-1 ... 순서로 개수만큼 순차 코드 생성."""
        try:
            result = self._service.spreadsheets().values().get(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{self._카테고리DB_시트}'!AG:AG",
            ).execute()
            rows = result.get("values", [])
        except Exception:
            rows = []

        pattern = re.compile(r'^H([A-Z])-(\d+)$', re.IGNORECASE)
        max_letter = 'A'
        max_num = 0
        for row in rows:
            if row:
                m = pattern.match(str(row[0]).strip().upper())
                if m:
                    letter = m.group(1).upper()
                    num = int(m.group(2))
                    if (letter, num) > (max_letter, max_num):
                        max_letter, max_num = letter, num

        코드들 = []
        letter, num = max_letter, max_num
        for _ in range(개수):
            num += 1
            if num > 9999:
                letter = chr(ord(letter) + 1)
                num = 1
            코드들.append(f"H{letter}-{num}")
        return 코드들

    def 카테고리DB_행추가(self, 제품목록: list[dict]) -> None:
        """
        카테고리 포함 db에 제품 행 추가.
        제품목록 각 dict 키: 브랜드, 시즌구분, 카테고리1, 카테고리2,
          바코드, 상품명, 발주서명, moq, 최종원가, 제품원가,
          패키지원가, 부자재원가, 리드타임, 등록일
        """
        if not 제품목록:
            return
        아이템코드들 = self._다음_아이템코드들(len(제품목록))
        행들 = []
        for p, 아이템 in zip(제품목록, 아이템코드들):
            bc = p.get("바코드", "")
            # 열 순서: A~AO (39열)
            # A=구분, B=브랜드, C=시즌구분, D=카테고리_1, E=카테고리_2,
            # F=LOCATION, G=상품코드, H=대표코드, I=상품명, J=사이즈체계,
            # K=주의사항, L=사이즈코드, M=바코드, N=단일로케이션,
            # O=최종원가, P=소싱원가, Q=수입원가, R=원가, S=주의메세지,
            # T=확인대상, U=등록자, V=등록일, W=수정자, X=수정일,
            # Y=발주서명, Z=색상코드, AA=MOQ, AB=최종 원가, AC=제품원가,
            # AD=패키지 원가, AE=부자재 원가, AF=리드타임, AG=아이템
            행 = [
                "",                         # A 구분
                p.get("브랜드", ""),         # B 브랜드
                p.get("시즌구분", ""),        # C 시즌구분
                p.get("카테고리1", ""),       # D 카테고리_1
                p.get("카테고리2", ""),       # E 카테고리_2
                "",                         # F LOCATION
                bc,                         # G 상품코드
                bc,                         # H 대표코드
                p.get("상품명", ""),          # I 상품명
                "",                         # J 사이즈체계
                "",                         # K 주의사항
                "",                         # L 사이즈코드
                bc,                         # M 바코드
                "",                         # N 단일로케이션
                "",                         # O 최종원가(구)
                "",                         # P 소싱원가
                "",                         # Q 수입원가
                "",                         # R 원가
                "",                         # S 주의메세지
                "",                         # T 확인대상
                "",                         # U 등록자
                p.get("등록일", ""),          # V 등록일
                "",                         # W 수정자
                "",                         # X 수정일
                p.get("발주서명", ""),        # Y 발주서명
                "",                         # Z 색상코드
                p.get("moq", ""),            # AA MOQ
                p.get("최종원가", ""),        # AB 최종 원가
                p.get("제품원가", ""),        # AC 제품원가
                p.get("패키지원가", ""),      # AD 패키지 원가
                p.get("부자재원가", ""),      # AE 부자재 원가
                p.get("리드타임", ""),        # AF 리드타임
                아이템,                      # AG 아이템
            ]
            행들.append(행)

        self._service.spreadsheets().values().append(
            spreadsheetId=self._스프레드시트_id,
            range=f"'{self._카테고리DB_시트}'!A:AG",
            valueInputOption="USER_ENTERED",
            body={"values": 행들},
        ).execute()
        print(f"[GoogleSheets] 카테고리DB {len(행들)}행 추가 완료")

    # ── N월 발주현황 탭 기록 ─────────────────────────────────────

    # 발주현황 필드 → 시트 헤더 키워드 매핑 (소문자 부분일치)
    _발주현황_필드맵 = {
        "발주일":    ["발주일", "발주일자", "date"],
        "바코드":    ["바코드", "barcode"],
        "주문번호":  ["주문번호", "발주번호", "po번호", "po코드", "po"],
        "발주수량":  ["발주수량", "수량", "qty", "quantity"],
        "제품명":    ["제품명", "상품명", "품명", "이름"],
        "브랜드명":  ["브랜드명", "브랜드", "brand"],
        "업체명":    ["업체명", "제조사", "업체", "공급사"],
        "비고":      ["비고", "note", "메모", "remark"],
        "색상코드":  ["색상코드", "색상", "color"],
        "시즌구분":  ["시즌구분", "시즌", "season"],
        "카테고리_1":["카테고리_1", "카테고리1", "대분류", "category1", "cat1"],
        "카테고리_2":["카테고리_2", "카테고리2", "소분류", "category2", "cat2"],
        "단가":      ["단가(최종단가)", "단가", "최종단가", "최종 단가", "price", "단가("],
    }

    def 발주현황_기록(self, 행들: list[dict]) -> None:
        """
        N월 발주현황 탭에 발주 내역 추가.
        탭이 없으면 자동 생성, 탭이 있으면 1행 헤더를 읽어 컬럼 위치 자동 매핑.
        """
        if not 행들:
            return
        from datetime import datetime as _dt
        탭명 = f"{_dt.now().month}월 발주현황"
        기본_헤더 = ["발주일", "바코드", "주문번호", "발주수량", "제품명", "브랜드명",
                    "업체명", "비고", "색상코드", "시즌구분", "카테고리_1", "카테고리_2", "단가(최종단가)"]

        # 탭 존재 확인
        meta = self._service.spreadsheets().get(
            spreadsheetId=self._스프레드시트_id
        ).execute()
        existing = {s["properties"]["title"] for s in meta.get("sheets", [])}

        if 탭명 not in existing:
            # 새 탭 생성 + 기본 헤더 삽입
            self._service.spreadsheets().batchUpdate(
                spreadsheetId=self._스프레드시트_id,
                body={"requests": [{"addSheet": {"properties": {"title": 탭명}}}]},
            ).execute()
            self._service.spreadsheets().values().update(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{탭명}'!A1",
                valueInputOption="USER_ENTERED",
                body={"values": [기본_헤더]},
            ).execute()
            print(f"[GoogleSheets] 새 탭 생성: {탭명}", flush=True)
            col_맵 = {field: i for i, field in enumerate(기본_헤더_키 for 기본_헤더_키 in [
                "발주일","바코드","주문번호","발주수량","제품명","브랜드명",
                "업체명","비고","색상코드","시즌구분","카테고리_1","카테고리_2","단가"
            ])}
        else:
            # 기존 탭 1행 헤더 읽기
            result = self._service.spreadsheets().values().get(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{탭명}'!1:1",
            ).execute()
            헤더_행 = result.get("values", [[]])[0] if result.get("values") else []
            헤더_lower = [str(h).strip().lower() for h in 헤더_행]
            col_맵 = self._발주현황_컬럼_매핑(헤더_lower)
            print(f"[GoogleSheets] {탭명} 헤더 매핑: {col_맵}", flush=True)

        # 최대 열 수 계산
        max_col = max(col_맵.values(), default=12) + 1

        rows = []
        for h in 행들:
            row = [""] * max_col
            for field, col_idx in col_맵.items():
                if col_idx < max_col:
                    row[col_idx] = h.get(field, "")
            rows.append(row)

        self._service.spreadsheets().values().append(
            spreadsheetId=self._스프레드시트_id,
            range=f"'{탭명}'!A:A",
            valueInputOption="USER_ENTERED",
            body={"values": rows},
        ).execute()
        print(f"[GoogleSheets] {탭명} {len(rows)}행 추가 완료", flush=True)

    # 항상 공란이라 시트에 없어도 무시해도 되는 필드
    _발주현황_선택필드 = {"비고", "색상코드"}

    def _발주현황_컬럼_매핑(self, 헤더_lower: list[str]) -> dict[str, int]:
        """헤더 리스트(소문자)를 읽어 필드명 → 열 인덱스 dict 반환.
        시트에 없는 필드는 선택필드면 제외, 필수필드면 마지막 열 뒤에 추가."""
        col_맵 = {}
        for field, keywords in self._발주현황_필드맵.items():
            for i, h in enumerate(헤더_lower):
                if any(kw in h for kw in keywords):
                    col_맵[field] = i
                    break
        # 매핑 안 된 필드: 선택필드는 건너뛰고, 필수필드만 뒤에 추가
        기본순서 = ["발주일","바코드","주문번호","발주수량","제품명","브랜드명",
                    "업체명","시즌구분","카테고리_1","카테고리_2","단가"]
        next_col = len(헤더_lower)
        for field in 기본순서:
            if field not in col_맵:
                col_맵[field] = next_col
                next_col += 1
        return col_맵

    def 아카이브_추가(self, 아카이브_시트명: str, 행들: list[list]) -> None:
        """아카이브 시트에 행 추가 (바코드, 제품명, 날짜, 발주번호 등)."""
        if not 행들:
            return
        self._service.spreadsheets().values().append(
            spreadsheetId=self._스프레드시트_id,
            range=f"{아카이브_시트명}!A:E",
            valueInputOption="USER_ENTERED",
            body={"values": 행들},
        ).execute()

    def _구글시트_탭_덮어쓰기(self, 탭명: str, rows: list[list]) -> None:
        """탭 전체를 rows로 덮어씀. 탭 없으면 생성 후 숨김 처리."""
        meta = self._service.spreadsheets().get(
            spreadsheetId=self._스프레드시트_id
        ).execute()
        existing = {s["properties"]["title"]: s["properties"]["sheetId"]
                    for s in meta.get("sheets", [])}

        if 탭명 not in existing:
            res = self._service.spreadsheets().batchUpdate(
                spreadsheetId=self._스프레드시트_id,
                body={"requests": [{"addSheet": {"properties": {
                    "title": 탭명, "hidden": True
                }}}]},
            ).execute()
        else:
            sheet_id = existing[탭명]
            self._service.spreadsheets().batchUpdate(
                spreadsheetId=self._스프레드시트_id,
                body={"requests": [{"updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "hidden": True},
                    "fields": "hidden",
                }}]},
            ).execute()
            self._service.spreadsheets().values().clear(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{탭명}'!A:ZZ",
            ).execute()

        self._service.spreadsheets().values().update(
            spreadsheetId=self._스프레드시트_id,
            range=f"'{탭명}'!A1",
            valueInputOption="USER_ENTERED",
            body={"values": rows},
        ).execute()

    # ── 재고현황 스냅샷 누적 ─────────────────────────────────────

    # Claude API가 읽기 쉽도록 순서 고정
    _스냅샷_헤더 = [
        "날짜", "바코드", "상품명", "브랜드", "대분류", "소분류", "기본_lt",
        "중국재고_쿠팡", "중국재고_일반", "사방넷재고", "쿠팡fc재고", "재고합계",
        "중국쿠팡_미입고", "중국일반_미입고", "쿠팡발주수량",
        "입고예정_코드", "입고일자",
        "쿠팡주간판매", "일반채널판매",
        "재고소진일", "안전재고일수", "예상판매_lt", "발주량산출", "재고상태",
    ]

    def 스냅샷_저장(self, rows: list[dict], 날짜: str = None) -> bool:
        """
        재고현황 전체 SKU를 아카이브 시트에 저장.
        - 같은 날짜 기존 행 삭제 후 덮어씀 (하루 1회 스냅샷)
        - 숫자는 숫자 타입으로 저장 (Claude API 분석용)
        """
        if not rows:
            return False
        날짜 = 날짜 or date.today().isoformat()
        시트명 = "재고현황_아카이브"

        try:
            # 1. 헤더 확인 — 없으면 생성
            existing = self._service.spreadsheets().values().get(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{시트명}'!A1:A",
            ).execute().get("values", [])

            if not existing or existing[0] != ["날짜"]:
                self._service.spreadsheets().values().update(
                    spreadsheetId=self._스프레드시트_id,
                    range=f"'{시트명}'!A1",
                    valueInputOption="RAW",
                    body={"values": [self._스냅샷_헤더]},
                ).execute()
                existing = [["날짜"]]  # 헤더만 있는 상태로 초기화

            # 2. 오늘 날짜 행 인덱스 수집 (1-based, 헤더=1행)
            삭제_인덱스 = [
                i  # 0-based for API
                for i, row in enumerate(existing)
                if row and row[0] == 날짜
            ]

            if 삭제_인덱스:
                시트_id = self._시트_id_조회(시트명)
                requests = [
                    {"deleteDimension": {"range": {
                        "sheetId": 시트_id,
                        "dimension": "ROWS",
                        "startIndex": idx,
                        "endIndex": idx + 1,
                    }}}
                    for idx in sorted(삭제_인덱스, reverse=True)
                ]
                self._service.spreadsheets().batchUpdate(
                    spreadsheetId=self._스프레드시트_id,
                    body={"requests": requests},
                ).execute()

            # 3. 전체 SKU 행 append
            def _n(v, default=0):
                try: return float(v) if v not in (None, "", "-") else default
                except (ValueError, TypeError): return default

            def _s(v):
                return str(v) if v not in (None,) else ""

            새_행들 = []
            for r in rows:
                새_행들.append([
                    날짜,
                    _s(r.get("바코드")),
                    _s(r.get("상품명") or r.get("제품명") or r.get("이름")),
                    _s(r.get("브랜드")),
                    _s(r.get("대분류")),
                    _s(r.get("소분류")),
                    _n(r.get("기본_lt")),
                    _n(r.get("중국재고_쿠팡")),
                    _n(r.get("중국재고_일반")),
                    _n(r.get("사방넷재고")),
                    _n(r.get("쿠팡fc재고")),
                    _n(r.get("재고합계")),
                    _n(r.get("중국쿠팡_미입고")),
                    _n(r.get("중국일반_미입고")),
                    _n(r.get("쿠팡발주수량")),
                    _s(r.get("입고예정_코드")),
                    _s(r.get("입고일자")),
                    _n(r.get("쿠팡주간판매")),
                    _n(r.get("일반채널판매")),
                    _n(r.get("재고소진일")) if r.get("재고소진일") not in (None, "") else "",
                    _n(r.get("안전재고일수")),
                    _n(r.get("예상판매_lt")),
                    _n(r.get("발주량산출")),
                    _s(r.get("재고상태")),
                ])

            self._service.spreadsheets().values().append(
                spreadsheetId=self._스프레드시트_id,
                range=f"'{시트명}'!A1",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": 새_행들},
            ).execute()

            print(f"[스냅샷] {날짜} {len(새_행들)}개 SKU 저장 완료")
            return True

        except Exception as e:
            print(f"[스냅샷] 저장 실패: {e}")
            return False

    def _시트_id_조회(self, 시트명: str) -> int:
        """시트명 → sheetId(숫자) 반환."""
        meta = self._service.spreadsheets().get(
            spreadsheetId=self._스프레드시트_id,
        ).execute()
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["title"] == 시트명:
                return sheet["properties"]["sheetId"]
        raise ValueError(f"시트 없음: {시트명}")
