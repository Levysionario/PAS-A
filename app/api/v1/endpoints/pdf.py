import os
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse
from app.schemas.prova import ProvaSchema
from app.core.pdf_engine.docx_builder import ProvaDocxBuilder

router = APIRouter()

OUTPUT_DIR = "storage/output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

@router.post("/gerar-docx-formatado")
async def gerar_docx_formatado(request: Request, prova: ProvaSchema):
    try:
        data = await request.json()
        nome_original = data.get("nome_arquivo_original", "avaliacao")

        nome_saida = f"{nome_original}_PASA.docx"
        caminho_saida = os.path.join(OUTPUT_DIR, nome_saida)

        builder = ProvaDocxBuilder(prova, caminho_saida, tipo_cabecalho="mensal")
        builder.build()

        return {
            "status": "sucesso",
            "download_url": f"/api/v1/pdf/download/{nome_saida}"
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
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )