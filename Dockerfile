FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY create_word.py .

EXPOSE 8001
CMD ["python", "create_word.py"]
