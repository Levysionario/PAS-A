import os
import shutil
from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from app.core.parser.docx_extractor import DocxParser
from app.core.ai.rescue_engine import AIRescueEngine
from app.schemas.prova import ProvaSchema

router = APIRouter()

UPLOAD_DIR = "storage/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def _extrair_texto_bruto_completo(doc_parser: DocxParser) -> str:
    linhas = []
    for p in doc_parser.doc.paragraphs:
        if p.text and p.text.strip():
            linhas.append(p.text.strip())

    for table in doc_parser.doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text and p.text.strip():
                        linhas.append(p.text.strip())

    return "\n".join(linhas)

@router.post("/processar-docx", response_model=ProvaSchema)
async def processar_documento(file: UploadFile = File(...)):
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Formato inválido. Envie um arquivo .docx")

    caminho_arquivo = os.path.join(UPLOAD_DIR, file.filename)
    
    with open(caminho_arquivo, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    parser = DocxParser(docx_path=caminho_arquivo)
    prova_parseada = parser.processar()

    eh_prova_anormal = False
    if len(prova_parseada.questoes) == 0:
        eh_prova_anormal = True

    if eh_prova_anormal:
        print("[PAS/A Engine] Anomalia detectada. Ativando AI-Rescue automático...")
        texto_completo = _extrair_texto_bruto_completo(parser)
        ai_engine = AIRescueEngine()
        prova_resgatada = ai_engine.estruturar_prova_com_ia(texto_completo)

        if len(prova_resgatada.questoes) > 0:
            return prova_resgatada

    return prova_parseada

@router.post("/refazer-com-ia", response_model=ProvaSchema)
async def refazer_com_ia(filename: str = Form(...)):
    caminho_arquivo = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(caminho_arquivo):
        raise HTTPException(status_code=404, detail="Arquivo original não encontrado no servidor.")

    parser = DocxParser(docx_path=caminho_arquivo)
    texto_completo = _extrair_texto_bruto_completo(parser)

    print(f"[PAS/A Engine] Reestruturação por IA solicitada manualmente para: {filename}")
    ai_engine = AIRescueEngine()
    prova_resgatada = ai_engine.estruturar_prova_com_ia(texto_completo)

    if len(prova_resgatada.questoes) == 0:
        raise HTTPException(status_code=422, detail="A IA não conseguiu estruturar as questões deste arquivo.")

    return prova_resgatada