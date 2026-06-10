FROM python:3.11-slim

# Install system deps: Tesseract OCR + poppler (pdftotext/pdftoppm)
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-fra \
    tesseract-ocr-deu \
    tesseract-ocr-spa \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY pyproject.toml .
RUN pip install --no-cache-dir .

COPY cv_screener.py criteria.yaml ./

ENTRYPOINT ["python", "cv_screener.py"]
