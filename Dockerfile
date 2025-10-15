# 汇享易报告自助生成智能体 - Dockerfile
FROM python:3.12-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    DEBIAN_FRONTEND=noninteractive

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖文件
COPY requirements.txt .

# 安装Python依赖
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# 复制应用代码
COPY app.py .
COPY excel_parser.py .
COPY date_calculator.py .
COPY word_generator.py .
COPY template.docx .

# 创建必要的目录
RUN mkdir -p upload output

# 暴露端口
EXPOSE 7861

# 健康检查
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import requests; requests.get('http://localhost:7861')" || exit 1

# 启动应用
CMD ["python", "app.py"]
