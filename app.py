"""
발주 자동화 - Flask 진입점
실행: python app.py
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from dotenv import load_dotenv
load_dotenv()

import threading
from flask import Flask
from presentation.routes import bp
from presentation.routes_발주 import bp_발주
from presentation.routes_재고현황 import bp_재고현황
from presentation.routes_pi import bp_pi
import config

app = Flask(
    __name__,
    template_folder="presentation/templates",
    static_folder="static",
)
app.register_blueprint(bp)
app.register_blueprint(bp_발주)
app.register_blueprint(bp_재고현황)
app.register_blueprint(bp_pi)


@app.errorhandler(Exception)
def _json_error(e):
    import traceback
    from flask import request as req, jsonify as jfy
    if req.path.startswith("/api/"):
        print(f"[앱 오류] {req.path}: {e}\n{traceback.format_exc()}", flush=True)
        return jfy({"success": False, "error": str(e)}), 500
    raise e


def _캐시_워밍업():
    """서버 시작 직후 백그라운드에서 발주서원본 캐시 미리 로딩."""
    try:
        from presentation.routes_발주 import _유스케이스
        _유스케이스._파서.카탈로그_로드()
    except Exception as e:
        print(f"[캐시 워밍업] 실패: {e}")


def _재고현황_워밍업():
    """서버 시작 직후 재고현황 JSON을 메모리에 미리 로딩."""
    try:
        from presentation.routes_재고현황 import 재고현황_캐시_로드
        재고현황_캐시_로드()
    except Exception as e:
        print(f"[재고현황 워밍업] 실패: {e}")


def _슬랙_스케줄러():
    """매일 09:00에 오늘 예정된 PI 알림 발송."""
    import time
    import datetime as _datetime
    from datetime import datetime as _dt
    while True:
        now = _dt.now()
        next_9am = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if now >= next_9am:
            next_9am = next_9am + _datetime.timedelta(days=1)
        time.sleep((next_9am - now).total_seconds())
        try:
            from presentation.routes_발주 import _유스케이스 as _uc
            _sh = _uc._sheets if _uc else None
            from application.pi_유스케이스 import PI유스케이스
            n = PI유스케이스(_sh).오늘_알림_발송()
            print(f"[슬랙 스케줄러] {n}건 발송 완료", flush=True)
        except Exception as e:
            print(f"[슬랙 스케줄러] 오류: {e}", flush=True)


if __name__ == "__main__":
    # 추가된 PO 시리즈 로드
    _추가시리즈_파일 = config.발주서_DIR / "po_추가시리즈.json"
    if _추가시리즈_파일.exists():
        import json as _json
        try:
            for s in _json.loads(_추가시리즈_파일.read_text(encoding="utf-8")):
                if not any(x["base"] == s["base"] for x in config.PO_시리즈):
                    config.PO_시리즈.append(s)
        except Exception:
            pass

    # 저장 폴더 없으면 생성
    config.발주서원본_DIR.mkdir(parents=True, exist_ok=True)
    config.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    config.통합발주서원본_DIR.mkdir(parents=True, exist_ok=True)
    config.발주서_DIR.mkdir(parents=True, exist_ok=True)
    config.재고현황_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    config.PI_DIR.mkdir(parents=True, exist_ok=True)
    config.업체도장_DIR.mkdir(parents=True, exist_ok=True)

    # 백그라운드에서 발주서원본 캐시 워밍업 (1.4GB 첫 파싱을 서버 시작 시점에 처리)
    # debug=True + reloader일 때 자식 프로세스에서만 실행
    import os
    if not config.DEBUG or os.environ.get("WERKZEUG_RUN_MAIN") == "true":
        threading.Thread(target=_캐시_워밍업, daemon=True).start()
        threading.Thread(target=_재고현황_워밍업, daemon=True).start()
        threading.Thread(target=_슬랙_스케줄러, daemon=True).start()

    print(f"서버 시작: http://localhost:{config.PORT}")
    print(f"저장 경로: {config.발주서원본_DIR}")
    print(f"[SMTP] FROM={config.SMTP_발신자!r} TO={config.SMTP_수신자!r} CC={config.SMTP_참조!r} PW={'설정됨' if config.SMTP_비밀번호 else '없음'}")
    app.run(host=config.HOST, port=config.PORT, debug=config.DEBUG)
