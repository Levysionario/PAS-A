import os
import re
import uuid
import shutil
from fastapi import APIRouter, HTTPException, Request, UploadFile, File
from fastapi.responses import FileResponse
from app.schemas.prova import ProvaSchema
from app.core.pdf_engine.docx_builder import ProvaDocxBuilder

router = APIRouter()

OUTPUT_DIR = "storage/output"
BANNERS_CUSTOM_DIR = "storage/banners_custom"
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(BANNERS_CUSTOM_DIR, exist_ok=True)


def _sanitizar_nome_arquivo(nome: str) -> str:
    base = os.path.splitext(nome or "avaliacao")[0]
    limpo = re.sub(r"[^\w\-]+", "_", base, flags=re.UNICODE).strip("_")
    return limpo or "avaliacao"


@router.post("/upload-banner")
def upload_banner_customizado(file: UploadFile = File(...)):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in {".png", ".jpg", ".jpeg", ".webp"}:
        raise HTTPException(status_code=400, detail="Formato inválido. Envie PNG ou JPG.")

    nome_unico = f"banner_{uuid.uuid4().hex[:10]}{ext}"
    caminho_destino = os.path.join(BANNERS_CUSTOM_DIR, nome_unico)

    with open(caminho_destino, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "status": "sucesso",
        "banner_path": caminho_destino,
        "filename": file.filename,
    }


@router.post("/gerar-docx-formatado")
async def gerar_docx_formatado(request: Request, prova: ProvaSchema):
    try:
        try:
            data = await request.json()
        except Exception:
            data = {}

        nome_bruto = (
            data.get("nome_arquivo_original")
            or prova.nome_arquivo_original
            or f"{prova.disciplina}_{prova.turma}"
        )
        nome_limpo = _sanitizar_nome_arquivo(nome_bruto)
        nome_saida = f"{nome_limpo}_PASA.docx"
        caminho_saida = os.path.join(OUTPUT_DIR, nome_saida)

        tipo_cab = data.get("tipo_cabecalho") or prova.tipo_avaliacao or "mensal"
        builder = ProvaDocxBuilder(prova, caminho_saida, tipo_cabecalho=tipo_cab)
        builder.build()

        return {
            "status": "sucesso",
            "download_url": f"/api/v1/pdf/download/{nome_saida}",
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar o arquivo Word: {str(e)}")


@router.get("/download/{nome_arquivo}")
async def download_arquivo(nome_arquivo: str):
    caminho_arquivo = os.path.join(OUTPUT_DIR, nome_arquivo)
    if not os.path.exists(caminho_arquivo):
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")

    return FileResponse(
        path=caminho_arquivo,
        filename=nome_arquivo,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )