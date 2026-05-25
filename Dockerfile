FROM python:3.10-slim

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

# 配合 Zeabur 標準，改為露出 8080 埠
EXPOSE 8080

# 啟動伺服器時，綁定在 8080 埠
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]
