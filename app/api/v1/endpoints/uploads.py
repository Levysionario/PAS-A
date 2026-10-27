import os
import shutil
from fastapi import APIRouter, UploadFile, File, HTTPException
from app.core.parser.docx_extractor import DocxParser
from app.schemas.prova import ProvaSchema

router = APIRouter()

UPLOAD_DIR = "storage/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@router.post("/processar-docx", response_model=ProvaSchema)
async def processar_documento(file: UploadFile = File(...)):
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Formato inválido. Envie um arquivo .docx")

    caminho_arquivo = os.path.join(UPLOAD_DIR, file.filename)
    
    # Salva o arquivo enviado no storage local
    with open(caminho_arquivo, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Executa o parser semi-automático
    parser = DocxParser(docx_path=caminho_arquivo)
    imagens = parser.extrair_imagens()
    prova_parseada = parser.processar()

    return prova_parseada