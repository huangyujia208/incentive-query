# Hugging Face Spaces (Docker) 兜底部署用
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# HF Spaces 默认通过 PORT 环境变量指定端口（通常为 7860）
ENV PORT=7860
EXPOSE 7860

CMD ["python", "app.py"]
