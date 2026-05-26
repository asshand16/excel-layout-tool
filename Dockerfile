FROM python:3.10-slim

# 安裝必要的系統依賴，但移除不必要的包
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fonts-noto-cjk \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* /tmp/* /var/tmp/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 設置環境變數以優化 LibreOffice 和 Python
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1
ENV SAL_NO_DIALOGS=1
ENV SAL_NO_SHOW_DIALOGS=1

EXPOSE 8080

# 優化的 Gunicorn 配置：
# --workers 1: 只用 1 個工作進程（節省記憶體）
# --threads 2: 每個進程 2 個線程（支持並發但不過度）
# --worker-class gthread: 使用線程工作類
# --timeout 120: 120 秒超時（防止卡住）
# --max-requests 100: 每 100 個請求後重啟工作進程（清理記憶體洩漏）
CMD ["gunicorn", \
     "--bind", "0.0.0.0:8080", \
     "--workers", "1", \
     "--threads", "2", \
     "--worker-class", "gthread", \
     "--timeout", "120", \
     "--max-requests", "100", \
     "--max-requests-jitter", "10", \
     "app:app"]
