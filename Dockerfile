FROM mcr.microsoft.com/azure-functions/python:4-python3.11

RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libzbar0 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . /home/site/wwwroot
