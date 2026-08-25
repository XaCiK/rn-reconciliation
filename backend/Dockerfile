FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY ./app /app

ENV LOG_DIR=/logs
ENV DATA_DIR=/data

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
