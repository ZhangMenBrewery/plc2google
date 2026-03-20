FROM python:3.11-slim

WORKDIR /app

COPY . /app

ENV TZ=Asia/Taipei
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# 安裝系統依賴
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /app/
RUN python3 -m pip install -r requirements.txt

# 暴露 Flask 網頁端口
EXPOSE 8001

# 預設運行 Flask 網頁應用程式
CMD ["python3", "-u", "app.py"]
