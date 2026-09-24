"""
OmniScan Server — API de conversion de documents.

Phase 1 : Word (.docx) -> PDF via LibreOffice en mode headless.
Phase 2 : PDF -> Word (.docx) via pdf2docx (texte, tableaux et images
reconstruits par une vraie analyse de mise en page), avec un post-traitement
qui force des bordures visibles sur tous les tableaux — pdf2docx ne
reproduit pas toujours fidèlement les lignes de grille de l'original.

pdf2docx s'appuie sur PyMuPDF, qui n'est pas thread-safe : deux conversions
PDF->Word lancées en même temps dans des threads différents peuvent se
corrompre silencieusement l'une l'autre. FastAPI exécute chaque endpoint
"def" (non "async def") dans un thread séparé, donc deux requêtes qui
arrivent proches dans le temps (ou une requête de conversion qui croise une
requête de monitoring) peuvent se chevaucher. On sérialise donc toutes les
conversions PDF->Word avec un verrou (pdf2docx_lock) : une seule à la fois.

Lancer en local pour tester :
    uvicorn main:app --host 0.0.0.0 --port 8000 --reload
"""

import os
import shutil
import subprocess
import tempfile
import threading
import uuid

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pdf2docx import Converter

app = FastAPI(title="OmniScan Server", version="0.2.2")

# Dossier de travail temporaire pour les fichiers reçus/générés.
WORK_DIR = os.path.join(tempfile.gettempdir(), "omniscan_server")
os.makedirs(WORK_DIR, exist_ok=True)

# Durée max qu'on laisse à LibreOffice pour convertir un document (secondes).
# Au-delà, on considère que ça a planté plutôt que de bloquer indéfiniment.
CONVERSION_TIMEOUT_SECONDS = 90

# Verrou global : garantit qu'une seule conversion pdf2docx (PyMuPDF) tourne
# à la fois dans tout le processus, quel que soit le nombre de requêtes
# reçues en parallèle.
pdf2docx_lock = threading.Lock()


def force_table_borders(docx_path: str) -> None:
    """
    Force des bordures visibles (grille complète) sur tous les tableaux d'un
    document. pdf2docx crée parfois la bonne structure de tableau (lignes,
    colonnes) mais sans reproduire fidèlement les lignes de grille visibles
    de l'original — on les rajoute systématiquement, quitte à ne pas
    correspondre exactement au style d'origine, pour garantir un tableau
    lisible et utilisable.
    """
    document = Document(docx_path)
    print(f"[DEBUG] Nombre de tableaux détectés : {len(document.tables)}")
    if not document.tables:
        return

    for table in document.tables:
        tbl_pr = table._tbl.tblPr

        # pdf2docx insère déjà un élément w:tblBorders (souvent réglé sur
        # "aucune bordure") — on doit le retirer avant d'ajouter le nôtre,
        # sinon Word se retrouve avec deux w:tblBorders dans le même
        # tableau, ce qui est invalide, et applique silencieusement le
        # premier (celui sans bordures) en ignorant le second.
        existing_borders = tbl_pr.find(qn("w:tblBorders"))
        if existing_borders is not None:
            tbl_pr.remove(existing_borders)

        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            edge_element = OxmlElement(f"w:{edge}")
            edge_element.set(qn("w:val"), "single")
            edge_element.set(qn("w:sz"), "4")
            edge_element.set(qn("w:space"), "0")
            edge_element.set(qn("w:color"), "000000")
            borders.append(edge_element)
        tbl_pr.append(borders)

    document.save(docx_path)


@app.get("/health")
def health_check():
    """Endpoint simple pour vérifier que le serveur répond (utile pour monitoring)."""
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

    request_id = str(uuid.uuid4())
    request_dir = os.path.join(WORK_DIR, request_id)
    os.makedirs(request_dir, exist_ok=True)

    input_path = os.path.join(request_dir, "input.docx")
    expected_output_path = os.path.join(request_dir, "input.pdf")
    lo_profile_dir = os.path.join(request_dir, "lo_profile")

    background_tasks.add_task(shutil.rmtree, request_dir, ignore_errors=True)

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

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


@app.post("/compress/pdf")
def compress_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...), quality: int = 50):
    """
    Reçoit un fichier .pdf et le recompresse via Ghostscript (ré-encodage des
    images internes, sous-échantillonnage des résolutions, nettoyage de la
    structure) — bien plus efficace qu'une compression "maison", car
    Ghostscript retravaille vraiment le contenu du PDF plutôt que de le
    manipuler en surface.

    [quality] (0-100, reçu du slider de l'appli) est mappé vers les
    préréglages standards de Ghostscript :
        <= 40  -> /screen   (~72 dpi, le plus compact)
        <= 70  -> /ebook    (~150 dpi, bon compromis)
        > 70   -> /printer  (~300 dpi, la meilleure qualité)
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Le fichier doit être un .pdf")

    if quality <= 40:
        gs_preset = "/screen"
    elif quality <= 70:
        gs_preset = "/ebook"
    else:
        gs_preset = "/printer"

    request_id = str(uuid.uuid4())
    request_dir = os.path.join(WORK_DIR, request_id)
    os.makedirs(request_dir, exist_ok=True)

    input_path = os.path.join(request_dir, "input.pdf")
    output_path = os.path.join(request_dir, "compressed.pdf")

    background_tasks.add_task(shutil.rmtree, request_dir, ignore_errors=True)

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        result = subprocess.run(
            [
                "gs",
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={gs_preset}",
                "-dNOPAUSE",
                "-dBATCH",
                "-dQUIET",
                f"-sOutputFile={output_path}",
                input_path,
            ],
            capture_output=True,
            text=True,
            timeout=CONVERSION_TIMEOUT_SECONDS,
        )

        if result.returncode != 0 or not os.path.exists(output_path):
            raise HTTPException(
                status_code=500,
                detail=f"Échec de la compression Ghostscript : {result.stderr.strip()}",
            )

        # Filet de sécurité : sur un PDF déjà très optimisé, Ghostscript peut
        # ressortir un fichier légèrement PLUS gros (ré-encodage moins
        # efficace que l'original). Dans ce cas, autant renvoyer l'original.
        if os.path.getsize(output_path) >= os.path.getsize(input_path):
            return FileResponse(
                input_path,
                media_type="application/pdf",
                filename=os.path.splitext(file.filename)[0] + "_compressed.pdf",
            )

        return FileResponse(
            output_path,
            media_type="application/pdf",
            filename=os.path.splitext(file.filename)[0] + "_compressed.pdf",
        )

    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=504, detail="La compression a dépassé le délai autorisé")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur de compression : {str(e)}")


@app.post("/convert/pdf-to-docx")
def convert_pdf_to_docx(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Reçoit un fichier .pdf, le convertit en .docx via pdf2docx, force des
    bordures visibles sur tous les tableaux détectés, et renvoie le résultat.

    Définie en fonction normale (pas "async def") : pdf2docx est bloquant —
    FastAPI l'exécute alors automatiquement dans un thread séparé. La
    conversion elle-même est protégée par pdf2docx_lock pour éviter que
    deux threads touchent PyMuPDF en même temps (voir note en haut du
    fichier).
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

        with pdf2docx_lock:
            converter = Converter(input_path)
            try:
                converter.convert(output_path, start=0, end=None)
            finally:
                converter.close()

        if not os.path.exists(output_path):
            raise HTTPException(status_code=500, detail="La conversion PDF vers Word a échoué")

        try:
            force_table_borders(output_path)
        except Exception as e:
            # Un échec du post-traitement des bordures ne doit pas faire perdre
            # toute la conversion : on renvoie quand même le document.
            print(f"Avertissement : échec du forçage des bordures ({e})")

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.splitext(file.filename)[0] + ".docx",
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur de conversion pdf2docx : {str(e)}")
