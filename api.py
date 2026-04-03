from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
import tempfile
from pathlib import Path
import uvicorn

# ← LIGNE 8 CORRIGÉE
from engine import convertpdf

app = FastAPI(title="OFX API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/convert/ofx")
async def convert_to_ofx(
    file: UploadFile = File(...),
    target: str = Form("quadra")
):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(await file.read())
        pdf_path