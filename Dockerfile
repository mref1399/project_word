FROM python:3.11-slim

WORKDIR /app

# کپی requirements
COPY requirements.txt .

# نصب کتابخانه‌ها
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# کپی کد
COPY create_word.py .

# تنظیمات
ENV PYTHONUNBUFFERED=1

EXPOSE 8001

CMD ["python", "create_word.py"]
