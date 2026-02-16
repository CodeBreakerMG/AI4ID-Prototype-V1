# Use a small Python base image
FROM python:3.11-slim

# Prevents Python from writing pyc files and forces stdout/stderr to be unbuffered
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Install OS dependencies:
# - poppler-utils: provides pdfinfo/pdftoppm used by pdf2image
# - tesseract-ocr: OCR engine used by pytesseract
# - libglib2.0-0, libsm6, libxext6, libxrender1: common runtime deps for imaging stacks
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    tesseract-ocr \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender1 \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Install Python deps first (better layer caching)
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt \
    && pip install --no-cache-dir gunicorn

# Copy the rest of the project
COPY . /app

# Create runtime directories (safe even if your app also creates them)
RUN mkdir -p /app/uploads /app/data /app/data/sessions

# Azure Web Apps uses PORT env var; default to 8000 for local
ENV PORT=8000
EXPOSE 8000

# Run Flask via Gunicorn
# app:app means "app.py" module, "app" Flask instance
CMD ["sh", "-c", "gunicorn -b 0.0.0.0:${PORT} --workers 2 --threads 4 --timeout 180 app:app"]