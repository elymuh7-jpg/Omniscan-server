# Image de base légère avec Python.
FROM python:3.11-slim

# LibreOffice Writer suffit pour le Word -> PDF (pas besoin de la suite complète,
# ce qui garde l'image plus petite). fonts-liberation évite que les polices
# soient remplacées par des équivalents visuellement différents.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
