FROM python:3.11-slim

WORKDIR /app

# System deps:
#   tesseract-ocr  → pytesseract (OCR fallback in excel-service)
#   fonts-liberation → PIL text rendering in image-service
#   libgl1         → OpenCV/PIL on headless systems
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    fonts-liberation \
    libgl1 \
    psmisc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Writable dirs created at build time; docker-compose volume mounts
# will overlay generated_images at runtime so output persists.
RUN mkdir -p generated_images \
             uploaded_product_images \
             uploaded_excel_templates \
             uploaded_catalogs

EXPOSE 7860 7861
