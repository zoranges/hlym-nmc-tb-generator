FROM python:3.12-slim

WORKDIR /app

# 安装依赖
COPY poc/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY poc/server.py .
COPY poc/ymc_app.py .
COPY poc/ymac_app.py .
COPY poc/lcp_app.py .
COPY poc/templates/ templates/

# 内置数据文件 (可通过 volume 挂载覆盖)
COPY ["NMC Project/data.xlsx", "/data/data.xlsx"]
ENV STRUCT_PATH="/data/data.xlsx"

# 输出目录
RUN mkdir -p /app/output
VOLUME ["/app/output", "/data"]

EXPOSE 8080

CMD ["python", "server.py"]
