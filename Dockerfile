FROM python:3.12.7-slim

WORKDIR /app

# 安装curl用于健康检查
RUN apt-get update && apt-get install -y curl && rm -rf /var/lib/apt/lists/*

# 复制项目文件
COPY requirements.txt .
COPY app.py .
COPY converters/ ./converters/
COPY storage/ ./storage/
COPY services/ ./services/

# 安装依赖
RUN pip install --no-cache-dir -r requirements.txt

# 创建存储目录
RUN mkdir -p tmp/storage

# 设置环境变量
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# 暴露端口
EXPOSE 5000

# 启动命令
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"] 