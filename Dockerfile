FROM python:3.11-slim

# 시스템 의존성 (Pillow 빌드용)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# pip 캐시 레이어 분리 (코드 변경 시 재설치 방지)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 앱 소스 복사
COPY . .

# 런타임에 필요한 폴더 생성
RUN mkdir -p uploads/재고현황 output "구글 시트 대체용" 발주서_output

EXPOSE 5006

ENV PYTHONUNBUFFERED=1

CMD ["python", "app.py"]
