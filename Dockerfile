FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-java-common \
    default-jre \
    poppler-utils \
    fonts-dejavu \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .

CMD gunicorn app:app --bind 0.0.0.0:8080 --timeout 120
