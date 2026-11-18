import os
import shutil
from docx import Document
from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from app.core.parser.docx_extractor import DocxParser
from app.core.ai.rescue_engine import AIRescueEngine
from app.schemas.prova import ProvaSchema

router = APIRouter()

UPLOAD_DIR = "storage/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)


def _extrair_texto_bruto_seguro(caminho_arquivo: str) -> str:
    """Lê todos os parágrafos e tabelas diretamente via python-docx sem depender do parser."""
    doc = Document(caminho_arquivo)
    linhas = []

    for p in doc.paragraphs:
        if p.text and p.text.strip():
            linhas.append(p.text.strip())

    for table in doc.tables:
        for row in table.rows:
            linha_celulas = []
            for cell in row.cells:
                texto_celula = " ".join(
                    p.text.strip() for p in cell.paragraphs if p.text and p.text.strip()
                )
                if texto_celula and texto_celula not in linha_celulas:
                    linha_celulas.append(texto_celula)
            if linha_celulas:
                linhas.append(" | ".join(linha_celulas))

    return "\n".join(linhas)


@router.post("/processar-docx", response_model=ProvaSchema)
async def processar_documento(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Formato inválido. Envie um arquivo .docx")

    caminho_arquivo = os.path.join(UPLOAD_DIR, file.filename)

    with open(caminho_arquivo, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    prova_parseada = None
    eh_prova_anormal = False

    try:
        parser = DocxParser(docx_path=caminho_arquivo)
        prova_parseada = parser.processar()
        if not prova_parseada or len(prova_parseada.questoes) == 0:
            eh_prova_anormal = True
    except Exception as e:
        print(f"[PAS/A Engine] Erro no DocxParser tradicional ({e}). Acionando AI-Rescue...")
        eh_prova_anormal = True

    if eh_prova_anormal:
        print("[PAS/A Engine] Ativando AI-Rescue automático...")
        texto_completo = _extrair_texto_bruto_seguro(caminho_arquivo)
        ai_engine = AIRescueEngine()
        prova_resgatada = ai_engine.estruturar_prova_com_ia(texto_completo)

        if len(prova_resgatada.questoes) > 0:
            prova_resgatada.nome_arquivo_original = file.filename
            return prova_resgatada

        motivo = ai_engine.ultimo_erro or "Não foi possível extrair questões deste arquivo."
        raise HTTPException(status_code=422, detail=motivo)

    prova_parseada.nome_arquivo_original = file.filename
    return prova_parseada


@router.post("/refazer-com-ia", response_model=ProvaSchema)
async def refazer_com_ia(filename: str = Form(...)):
    caminho_arquivo = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(caminho_arquivo):
        raise HTTPException(
            status_code=404,
            detail="Arquivo original não encontrado no servidor (o servidor pode ter reiniciado). Envie o .docx novamente.",
        )

    texto_completo = _extrair_texto_bruto_seguro(caminho_arquivo)

    print(f"[PAS/A Engine] Reestruturação por IA solicitada manualmente para: {filename}")
    ai_engine = AIRescueEngine()
    prova_resgatada = ai_engine.estruturar_prova_com_ia(texto_completo)

    if len(prova_resgatada.questoes) == 0:
        motivo = ai_engine.ultimo_erro or "A IA não conseguiu estruturar as questões deste arquivo."
        raise HTTPException(status_code=422, detail=motivo)

    prova_resgatada.nome_arquivo_original = filename
    return prova_resgatada