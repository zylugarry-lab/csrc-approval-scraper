FROM python:3.11-slim

WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8787

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN mkdir -p output

EXPOSE 8787
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT}"]
