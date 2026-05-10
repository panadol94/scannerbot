FROM python:3.11-slim
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
CMD exec gunicorn -b :$PORT --workers 2 --threads 8 --timeout 60 --graceful-timeout 30 --max-requests 1000 --max-requests-jitter 100 main:app
