"""
OmniScan Server — API de conversion de documents.

Phase 1 : Word (.docx) -> PDF via LibreOffice en mode headless.
Phase 2 : PDF -> Word (.docx) via pdf2docx (texte, tableaux et images
reconstruits par une vraie analyse de mise en page, contrairement à nos
heuristiques OpenCV faites à la main côté Android).

Lancer en local pour tester :
    uvicorn main:app --host 0.0.0.0 --port 8000 --reload
"""

import os
import shutil
import subprocess
import tempfile
import uuid

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pdf2docx import Converter

app = FastAPI(title="OmniScan Server", version="0.2.0")

# Dossier de travail temporaire pour les fichiers reçus/générés.
WORK_DIR = os.path.join(tempfile.gettempdir(), "omniscan_server")
os.makedirs(WORK_DIR, exist_ok=True)

# Durée max qu'on laisse à LibreOffice pour convertir un document (secondes).
# Au-delà, on considère que ça a planté plutôt que de bloquer indéfiniment.
CONVERSION_TIMEOUT_SECONDS = 90


@app.get("/health")
def health_check():
    """Endpoint simple pour vérifier que le serveur répond (utile pour Render/monitoring)."""
    return {"status": "ok"}


@app.post("/convert/docx-to-pdf")
async def convert_docx_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Reçoit un fichier .docx, le convertit en PDF via LibreOffice headless,
    et renvoie le PDF résultant. Préserve fidèlement logos, tableaux et
    mise en page — contrairement à la conversion faite à la main côté app.
    """
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Le fichier doit être un .docx")

    # Dossier isolé par requête pour éviter les collisions entre utilisateurs simultanés.
    request_id = str(uuid.uuid4())
    request_dir = os.path.join(WORK_DIR, request_id)
    os.makedirs(request_dir, exist_ok=True)

    input_path = os.path.join(request_dir, "input.docx")
    expected_output_path = os.path.join(request_dir, "input.pdf")
    # Profil LibreOffice isolé par requête : sans ça, deux conversions simultanées
    # sur le même serveur peuvent se bloquer mutuellement (verrou de profil).
    lo_profile_dir = os.path.join(request_dir, "lo_profile")

    # Nettoyage programmé pour APRÈS l'envoi de la réponse (le fichier doit encore
    # exister pendant que FileResponse le transmet).
    background_tasks.add_task(shutil.rmtree, request_dir, ignore_errors=True)

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # LibreOffice headless : convertit et dépose le résultat dans --outdir.
        result = subprocess.run(
            [
                "soffice",
                "--headless",
                "--norestore",
                f"-env:UserInstallation=file://{lo_profile_dir}",
                "--convert-to", "pdf",
                "--outdir", request_dir,
                input_path,
            ],
            capture_output=True,
            text=True,
            timeout=CONVERSION_TIMEOUT_SECONDS,
        )

        if result.returncode != 0 or not os.path.exists(expected_output_path):
            raise HTTPException(
                status_code=500,
                detail=f"Échec de la conversion LibreOffice : {result.stderr.strip()}",
            )

        return FileResponse(
            expected_output_path,
            media_type="application/pdf",
            filename=os.path.splitext(file.filename)[0] + ".pdf",
        )

    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=504, detail="La conversion a dépassé le délai autorisé")


@app.post("/convert/pdf-to-docx")
def convert_pdf_to_docx(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Reçoit un fichier .pdf (par exemple un document scanné), le convertit en
    .docx via pdf2docx, et renvoie le résultat. Texte réel modifiable, vrais
    tableaux Word, images reconstruites à leur place — une vraie analyse de
    mise en page, contrairement aux heuristiques OpenCV (une seule grille de
    tableau détectée par page, seuils à calibrer) qu'on a dû bricoler côté
    Android.

    Définie en fonction normale (pas "async def") : pdf2docx est une
    bibliothèque bloquante — FastAPI l'exécute alors automatiquement dans un
    thread séparé, pour ne pas geler les autres requêtes en attente.
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Le fichier doit être un .pdf")

    request_id = str(uuid.uuid4())
    request_dir = os.path.join(WORK_DIR, request_id)
    os.makedirs(request_dir, exist_ok=True)

    input_path = os.path.join(request_dir, "input.pdf")
    output_path = os.path.join(request_dir, "output.docx")

    background_tasks.add_task(shutil.rmtree, request_dir, ignore_errors=True)

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        converter = Converter(input_path)
        try:
            converter.convert(output_path, start=0, end=None)
        finally:
            converter.close()

        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="La conversion PDF vers Word a échoué")

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.splitext(file.filename)[0] + ".docx",
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur de conversion pdf2docx : {str(e)}")
