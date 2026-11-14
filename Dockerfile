FROM python:3.11-slim
ENV PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
# Servicio (dashboard). El Job usa la misma imagen con el comando: python nube.py
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 120 app:app
