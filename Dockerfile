# Image de base légère avec Python.
FROM python:3.11-slim

# LibreOffice Writer suffit pour le Word -> PDF (pas besoin de la suite complète,
# ce qui garde l'image plus petite). fonts-liberation évite que les polices
# soient remplacées par des équivalents visuellement différents.
# Ghostscript sert à la compression PDF (endpoint /compress/pdf).
# Tesseract (+ le pack de langue française) sert à la détection de tableaux
# et à l'OCR des scans (endpoint /ocr/images-to-docx, via img2table).
# libgl1 est une dépendance runtime d'OpenCV (utilisé par img2table pour la
# détection des lignes de grille), même en usage "headless" côté serveur.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-liberation \
    ghostscript \
    tesseract-ocr \
    tesseract-ocr-fra \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
