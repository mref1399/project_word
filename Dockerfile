FROM python:3.11-slim

# نصب وابستگی‌های سیستمی ضروری
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    build-essential \
    libffi-dev \
    libssl-dev \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# کپی requirements
COPY requirements.txt .

# نصب کتابخانه‌ها به‌صورت مرحله‌ای
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir flask==3.0.0 && \
    pip install --no-cache-dir python-docx==1.1.0 && \
    pip install --no-cache-dir hazm==0.7.0 && \
    pip install --no-cache-dir sympy==1.12 && \
    pip install --no-cache-dir Werkzeug==3.0.1

# کپی کد
COPY create_word.py .

# تنظیمات
ENV PYTHONUNBUFFERED=1

EXPOSE 8001

CMD ["python", "create_word.py"]
