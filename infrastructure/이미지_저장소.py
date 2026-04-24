"""
이미지 저장소 - 업로드 이미지 임시 관리
"""
import uuid
from pathlib import Path
from werkzeug.datastructures import FileStorage

_허용_확장자 = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}


class 이미지_저장소:

    def __init__(self, upload_dir: Path):
        self.upload_dir = upload_dir
        upload_dir.mkdir(parents=True, exist_ok=True)

    def 저장(self, file: FileStorage) -> Path | None:
        if not file or not file.filename:
            return None
        ext = Path(file.filename).suffix.lower()
        if ext not in _허용_확장자:
            return None
        filename = f"{uuid.uuid4().hex}{ext}"
        path = self.upload_dir / filename
        file.save(str(path))
        return path

    def 삭제(self, path: Path):
        if path and path.exists():
            path.unlink()
