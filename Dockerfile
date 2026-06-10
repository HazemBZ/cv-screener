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

COPY cv_screener.py webapp.py criteria.yaml ./
COPY templates/ ./templates/

# Expose the web port
EXPOSE 8080

ENV PORT=8080

ENTRYPOINT ["python", "webapp.py"]
