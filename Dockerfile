FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libzbar0 \
    libzxing-dev \
    zbar-tools \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

CMD ["python", "ticket_processor.py"]
