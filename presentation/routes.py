"""
Flask 라우트
"""
from flask import Blueprint, render_template, request, jsonify, send_file
from pathlib import Path
import time
import config
from application.발주서_생성_유스케이스 import 발주서_생성_유스케이스

bp = Blueprint("main", __name__)

_CACHE_TTL = 600  # 10분

# ── 캐시 ─────────────────────────────────────────────────────────
_cache = {
    "바코드맵":   {"data": {}, "ts": 0},   # {바코드: 상품명}
    "카테고리맵": {"data": {}, "ts": 0},   # {대분류: [소분류...]}
}


def _sheets_연결():
    if config.GOOGLE_SERVICE_ACCOUNT_JSON and config.GOOGLE_SHEET_ID:
        try:
            from infrastructure.google_sheets_저장소 import GoogleSheets저장소
            return GoogleSheets저장소(
                config.GOOGLE_SERVICE_ACCOUNT_JSON,
                config.GOOGLE_SHEET_ID,
                config.GOOGLE_SHEET_NAME,
            )
        except Exception as e:
            print(f"[routes] Sheets 연결 실패: {e}")
    return None


_sheets = _sheets_연결()
_유스케이스 = 발주서_생성_유스케이스(config.발주서원본_DIR, config.UPLOAD_DIR, sheets_저장소=_sheets)


def _카테고리DB_로드():
    """카테고리 포함 db M/I/D/E 열을 한 번에 읽어 캐시 갱신."""
    if _sheets is None:
        return
    try:
        result = _sheets._service.spreadsheets().values().get(
            spreadsheetId=_sheets._스프레드시트_id,
            range="'카테고리 포함 db'!D:M",  # D=카테고리1, E=카테고리2, M=바코드 (col 10)
        ).execute()
        rows = result.get("values", [])
        바코드맵 = {}
        카테고리맵 = {}
        for row in rows[1:]:
            # D열(0), E열(1), ... M열(9)
            cat1 = str(row[0]).strip() if len(row) > 0 else ""
            cat2 = str(row[1]).strip() if len(row) > 1 else ""
            bc   = str(row[9]).strip() if len(row) > 9 else ""
            # 상품명은 I열 = D기준 5번째(idx5)
            상품명 = str(row[5]).strip() if len(row) > 5 else ""
            if bc:
                바코드맵[bc] = 상품명 or bc
            if cat1:
                if cat1 not in 카테고리맵:
                    카테고리맵[cat1] = []
                if cat2 and cat2 not in 카테고리맵[cat1]:
                    카테고리맵[cat1].append(cat2)
        now = time.time()
        _cache["바코드맵"]["data"] = 바코드맵
        _cache["바코드맵"]["ts"]   = now
        _cache["카테고리맵"]["data"] = 카테고리맵
        _cache["카테고리맵"]["ts"]   = now
        print(f"[cache] 카테고리DB 로드: 바코드 {len(바코드맵)}개, 카테고리 {len(카테고리맵)}개")
    except Exception as e:
        print(f"[cache] 카테고리DB 로드 실패: {e}")


def _캐시_갱신_필요(키: str) -> bool:
    return time.time() - _cache[키]["ts"] > _CACHE_TTL


@bp.route("/")
def index():
    return render_template("발주서_생성.html")


@bp.route("/api/카테고리목록", methods=["GET"])
def 카테고리목록():
    if _캐시_갱신_필요("카테고리맵"):
        _카테고리DB_로드()
    return jsonify({"카테고리맵": _cache["카테고리맵"]["data"]})


@bp.route("/api/발주서/바코드중복확인", methods=["GET"])
def 바코드중복확인():
    바코드 = request.args.get("바코드", "").strip()
    if not 바코드:
        return jsonify({"중복": False})
    if _캐시_갱신_필요("바코드맵"):
        _카테고리DB_로드()
    맵 = _cache["바코드맵"]["data"]
    if 바코드 in 맵:
        return jsonify({"중복": True, "상품명": 맵[바코드]})
    return jsonify({"중복": False})


@bp.route("/api/발주서/제품DB조회", methods=["GET"])
def 제품DB조회():
    """바코드로 제품DB에서 MOQ/원가/LT 등 조회."""
    바코드 = request.args.get("바코드", "").strip()
    if not 바코드:
        return jsonify({"found": False})
    try:
        from infrastructure.엑셀_시트_저장소 import 엑셀시트저장소
        if not config.로컬시트_파일.exists():
            return jsonify({"found": False})
        저장소 = 엑셀시트저장소(config.로컬시트_파일)
        db = 저장소.제품DB_읽기()
        info = db.get(바코드)
        if not info:
            return jsonify({"found": False})
        return jsonify({
            "found": True,
            "moq":      info.get("moq", ""),
            "원가_위안": info.get("원가_위안", ""),
            "상품명":   info.get("상품명", ""),
            "발주서명": info.get("발주서명", ""),
        })
    except Exception as e:
        return jsonify({"found": False, "error": str(e)})


@bp.route("/api/발주서/생성", methods=["POST"])
def 발주서_생성():
    result = _유스케이스.실행(request.form, request.files)

    if not result["success"]:
        return jsonify(result), 400

    # 카테고리DB에 추가됐으면 캐시 무효화
    _cache["바코드맵"]["ts"] = 0
    _cache["카테고리맵"]["ts"] = 0

    # 발주 실행 페이지에 신규 DB 알림 전달
    try:
        from application.알림_저장소 import 알림_추가
        알림_항목 = []
        idx = 1
        while f"item_{idx}" in request.form:
            bc = request.form.get(f"바코드_{idx}", "").strip()
            name = request.form.get(f"이름_{idx}", "").strip()
            if bc:
                알림_항목.append({"바코드": bc, "이름": name})
            idx += 1
        알림_추가(알림_항목)
    except Exception as _e:
        print(f"[routes] 알림 추가 실패: {_e}")

    # NAS 통합발주서원본 폴더에 자동 저장
    try:
        config.통합발주서원본_DIR.mkdir(parents=True, exist_ok=True)
        저장_경로 = config.통합발주서원본_DIR / result["파일명"]
        result["버퍼"].seek(0)
        저장_경로.write_bytes(result["버퍼"].read())
        result["버퍼"].seek(0)
    except Exception as e:
        print(f"[routes] NAS 저장 실패: {e}")

    return send_file(
        result["버퍼"],
        as_attachment=True,
        download_name=result["파일명"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
