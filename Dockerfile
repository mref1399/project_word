FROM python:3.11-slim

# نصب وابستگی‌های سیستمی
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# کپی و نصب requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# کپی کد اصلی
COPY create_word.py .

# تنظیم متغیر محیطی
ENV PYTHONUNBUFFERED=1

# باز کردن پورت
EXPOSE 8001

# اجرای برنامه
CMD ["python", "create_word.py"]
