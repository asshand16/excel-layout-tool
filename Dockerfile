FROM python:3.10-slim

# 安裝雲端轉檔必備的系統工具與繁體中文字型（避免報價單中文字漏字或跑版）
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fonts-noto-cjk \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]