#!/bin/bash
apt-get update -y
apt-get install -y tesseract-ocr libzbar0 2>/dev/null || true
pip install -r requirements.txt
