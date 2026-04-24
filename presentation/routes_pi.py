"""
PI 관련 API 라우트
- 계좌 DB CRUD
- PI 생성 / 수정 / 다운로드
- PI 대시보드
- 스크린샷/텍스트 파싱
"""
from pathlib import Path
from flask import Blueprint, render_template, request, jsonify, send_file, abort
import config
from infrastructure.계좌_저장소 import 계좌저장소
from application.pi_유스케이스 import PI유스케이스

bp_pi = Blueprint("pi", __name__)

try:
    from presentation.routes_발주 import _유스케이스 as _발주_uc
    _sheets = _발주_uc._sheets if _발주_uc else None
except Exception:
    _sheets = None

_계좌저장소 = 계좌저장소(_sheets)
_pi_uc = PI유스케이스(_sheets)


# ── 페이지 ──────────────────────────────────────────────────────

@bp_pi.route("/pi")
def pi_페이지():
    return render_template("pi_관리.html")


# ── 계좌 DB API ─────────────────────────────────────────────────

@bp_pi.route("/api/계좌/목록", methods=["GET"])
def 계좌_목록():
    return jsonify({"success": True, "data": list(_계좌저장소.전체_조회().values())})


@bp_pi.route("/api/계좌/저장", methods=["POST"])
def 계좌_저장():
    data = request.get_json(force=True)
    회사id = _계좌저장소.저장(data)
    return jsonify({"success": True, "id": 회사id})


@bp_pi.route("/api/계좌/삭제/<cid>", methods=["DELETE"])
def 계좌_삭제(cid):
    ok = _계좌저장소.삭제(cid)
    return jsonify({"success": ok})


@bp_pi.route("/api/계좌/누아트_도장_업로드", methods=["POST"])
def 누아트_도장_업로드():
    f = request.files.get("file")
    if not f:
        return jsonify({"success": False, "error": "파일 없음"}), 400
    ext = Path(f.filename).suffix or ".png"
    경로 = config.BASE_DIR / "assets" / f"nuart_도장{ext}"
    경로.parent.mkdir(parents=True, exist_ok=True)
    경로.write_bytes(f.read())
    # config.누아트_도장 경로가 다를 경우 덮어쓰기
    if 경로 != config.누아트_도장:
        import shutil as _shutil
        _shutil.copy2(경로, config.누아트_도장)
    return jsonify({"success": True, "경로": str(경로)})


@bp_pi.route("/api/계좌/도장_업로드/<cid>", methods=["POST"])
def 도장_업로드(cid):
    f = request.files.get("file")
    if not f:
        return jsonify({"success": False, "error": "파일 없음"}), 400
    ext = Path(f.filename).suffix or ".png"
    경로 = _계좌저장소.도장_저장(cid, f.read(), ext)
    return jsonify({"success": True, "경로": 경로})


@bp_pi.route("/api/계좌/텍스트_파싱", methods=["POST"])
def 텍스트_파싱():
    data = request.get_json(force=True)
    text = data.get("text", "")
    if not text:
        return jsonify({"success": False, "error": "text 없음"}), 400
    결과 = 계좌저장소.텍스트_파싱(text)
    return jsonify({"success": True, "data": 결과})




# ── PI API ──────────────────────────────────────────────────────

@bp_pi.route("/api/pi/계좌_자동매칭", methods=["POST"])
def pi_계좌_자동매칭():
    """바코드 목록 → 카탈로그 제조사명 → 계좌DB 매칭 → 계좌id + 회사명_중문 반환."""
    from datetime import date as _date
    data = request.get_json(force=True)
    barcodes = data.get("barcodes", [])
    제조사명 = data.get("제조사명", "")

    # 카탈로그로 제조사명 + 계좌id 확인
    계좌id_직접 = None
    try:
        if _발주_uc:
            카탈로그 = _발주_uc._파서.카탈로그_로드()
            for bc in barcodes:
                상품 = 카탈로그.get(str(bc))
                if 상품:
                    if not 제조사명 and 상품.제조사:
                        제조사명 = 상품.제조사
                    if 상품.계좌id:
                        계좌id_직접 = 상품.계좌id
                        break
    except Exception as e:
        print(f"[계좌_자동매칭] 카탈로그 조회 실패: {e}")

    db = _계좌저장소.전체_조회()

    # 0단계: 발주서에 계좌id가 직접 저장된 경우
    if 계좌id_직접 and str(계좌id_직접) in db:
        계좌 = db[str(계좌id_직접)]
        return jsonify({"success": True, "계좌id": 계좌["id"],
                        "회사명_중문": 계좌.get("회사명_중문", "")})

    if not 제조사명:
        return jsonify({"success": False, "error": "제조사 없음"})

    def _n(s): return (s or "").strip().lower()
    key = _n(제조사명)

    # 1단계: 회사명_발주서 정확 매칭
    for 계좌 in db.values():
        if _n(계좌.get("회사명_발주서")) == key:
            return jsonify({"success": True, "계좌id": 계좌["id"],
                            "회사명_중문": 계좌.get("회사명_중문", "")})

    # 2단계: 부분 매칭 (발주서명 ↔ 제조사명 포함 관계)
    for 계좌 in db.values():
        발주 = _n(계좌.get("회사명_발주서"))
        영문 = _n(계좌.get("회사명_영문"))
        if (발주 and (key in 발주 or 발주 in key)) or \
           (영문 and (key in 영문 or 영문 in key)):
            return jsonify({"success": True, "계좌id": 계좌["id"],
                            "회사명_중문": 계좌.get("회사명_중문", "")})

    return jsonify({"success": False, "error": f"계좌 매칭 없음: {제조사명}"})


@bp_pi.route("/api/pi/생성", methods=["POST"])
def pi_생성():
    from datetime import date as _date
    from infrastructure.슬랙_저장소 import pi_알림_발송
    data = request.get_json(force=True)
    결과 = _pi_uc.PI_생성(
        po코드=data["po코드"],
        회사id=data["회사id"],
        제품목록=data["제품목록"],
        디포짓_목록=data["디포짓"],
        담당자=data["담당자"],
        담당자_영문=data.get("담당자_영문", data.get("담당자", "")),
        담당자이메일=data.get("담당자이메일", ""),
        중국담당자=data.get("중국담당자", ""),
        중국이메일=data.get("중국이메일", ""),
        중국연락처=data.get("중국연락처", ""),
        max_lt=int(data.get("max_lt", 0)),
    )

    # 1차 디포짓 당일이면 즉시 슬랙 알림
    if 결과.get("success"):
        오늘 = _date.today().isoformat()
        디포짓 = data.get("디포짓", [])
        if 디포짓:
            dep = 디포짓[0]
            if dep.get("예정일") == 오늘:
                총액 = sum(
                    int(p.get("qty", 0)) * float(p.get("unit_price", 0))
                    for p in data.get("제품목록", [])
                )
                pct = float(dep.get("pct", 0))
                amt = float(dep.get("amount") or (총액 * pct))
                pi_알림_발송(
                    data.get("담당자", ""),
                    data["po코드"],
                    dep.get("label", "1st Deposit"),
                    pct,
                    amt,
                    data.get("제품목록", []),
                )

    return jsonify(결과), 200 if 결과["success"] else 400


@bp_pi.route("/api/pi/디포짓_수정", methods=["POST"])
def pi_디포짓_수정():
    data = request.get_json(force=True)
    결과 = _pi_uc.PI_디포짓_수정(data["po코드"], data["디포짓"])
    return jsonify(결과), 200 if 결과["success"] else 400


@bp_pi.route("/api/pi/삭제", methods=["DELETE"])
def pi_삭제():
    data = request.get_json(force=True)
    결과 = _pi_uc.PI_삭제(data["po코드"])
    return jsonify(결과)


@bp_pi.route("/api/pi/완료_처리", methods=["POST"])
def pi_완료_처리():
    data = request.get_json(force=True)
    결과 = _pi_uc.PI_완료_처리(data["po코드"], int(data["인덱스"]))
    return jsonify(결과)


@bp_pi.route("/api/pi/목록", methods=["GET"])
def pi_목록():
    return jsonify({"success": True, "data": _pi_uc.전체_PI_목록()})


@bp_pi.route("/api/pi/담당자조회", methods=["GET"])
def pi_담당자_조회():
    담당자 = request.args.get("name", "").strip()
    if not 담당자:
        return jsonify({"success": False, "error": "name 필요"}), 400
    return jsonify({"success": True, "data": _pi_uc.담당자_PI_목록(담당자, _sheets)})


@bp_pi.route("/api/pi/다운로드", methods=["GET"])
def pi_다운로드():
    po코드 = request.args.get("po", "").strip()
    if not po코드:
        abort(400)
    db = _pi_uc._db_읽기()
    meta = db.get(po코드)
    if not meta or not meta.get("파일경로"):
        abort(404)
    경로 = Path(meta["파일경로"])
    if not 경로.exists():
        abort(404)
    return send_file(
        경로,
        as_attachment=True,
        download_name=경로.name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@bp_pi.route("/api/pi/템플릿_다운로드", methods=["GET"])
def pi_템플릿_다운로드():
    if not config.PI_템플릿.exists():
        abort(404)
    return send_file(
        config.PI_템플릿,
        as_attachment=True,
        download_name="pi_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@bp_pi.route("/api/pi/제조사_매칭_현황", methods=["GET"])
def 제조사_매칭_현황():
    """카탈로그 제조사명 ↔ 계좌DB 매칭 결과 반환."""
    try:
        if not _발주_uc:
            return jsonify({"success": False, "error": "발주 UC 없음"})
        카탈로그 = _발주_uc._파서.카탈로그_로드()
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

    def _n(s): return (s or "").strip().lower()

    # 카탈로그 제조사 유니크 목록
    제조사_set = {str(p.제조사).strip() for p in 카탈로그.values() if p.제조사}
    계좌_db = _계좌저장소.전체_조회()

    결과 = []
    for 제조사명 in sorted(제조사_set):
        key = _n(제조사명)
        매칭 = None
        for 계좌 in 계좌_db.values():
            발주 = _n(계좌.get("회사명_발주서"))
            영문 = _n(계좌.get("회사명_영문"))
            if 발주 == key or 영문 == key:
                매칭 = 계좌
                break
        if not 매칭:
            for 계좌 in 계좌_db.values():
                발주 = _n(계좌.get("회사명_발주서"))
                영문 = _n(계좌.get("회사명_영문"))
                if (발주 and (key in 발주 or 발주 in key)) or \
                   (영문 and (key in 영문 or 영문 in key)):
                    매칭 = 계좌
                    break
        결과.append({
            "제조사명": 제조사명,
            "매칭여부": 매칭 is not None,
            "계좌id": 매칭["id"] if 매칭 else None,
            "계좌명_중문": 매칭.get("회사명_중문", "") if 매칭 else "",
            "계좌명_영문": 매칭.get("회사명_영문", "") if 매칭 else "",
        })

    매칭됨 = [r for r in 결과 if r["매칭여부"]]
    미매칭 = [r for r in 결과 if not r["매칭여부"]]
    return jsonify({"success": True, "매칭됨": 매칭됨, "미매칭": 미매칭})


@bp_pi.route("/api/pi/슬랙_테스트", methods=["POST"])
def 슬랙_테스트():
    from infrastructure.슬랙_저장소 import 슬랙_발송
    ok = 슬랙_발송("발주 자동화 슬랙 연결 테스트")
    return jsonify({"success": ok})
