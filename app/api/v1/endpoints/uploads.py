import os
import shutil
from fastapi import APIRouter, UploadFile, File, HTTPException
from app.core.parser.docx_extractor import DocxParser
from app.core.ai.rescue_engine import AIRescueEngine
from app.schemas.prova import ProvaSchema

router = APIRouter()

UPLOAD_DIR = "storage/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

def _extrair_texto_bruto_completo(doc_parser: DocxParser) -> str:
    """Extrai todo o texto corrido do documento (parágrafos e tabelas) para enviar ao AI-Rescue"""
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
    
    # Salva o arquivo enviado no storage local
    with open(caminho_arquivo, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # 1. Executa o parser local semi-automático
    parser = DocxParser(docx_path=caminho_arquivo)
    imagens = parser.extrair_imagens()
    prova_parseada = parser.processar()

    # 2. Heurística de Validação de Anomalias
    # Ativa a IA se o parser local não encontrar questões ou se as questões extraídas não tiverem alternativas
    eh_prova_anormal = False
    if len(prova_parseada.questoes) == 0:
        eh_prova_anormal = True
    else:
        # Se mais de 50% das questões extraídas estiverem sem alternativas, considera prova exótica
        questoes_sem_alt = sum(1 for q in prova_parseada.questoes if len(q.alternativas) == 0)
        if questoes_sem_alt / len(prova_parseada.questoes) >= 0.5:
            eh_prova_anormal = True

    # 3. Fallback inteligente via AI-Rescue
    if eh_prova_anormal:
        print("[PAS/A Engine] Anomalia estrutural detectada pelo parser local. Ativando AI-Rescue...")
        texto_completo = _extrair_texto_bruto_completo(parser)
        
        ai_engine = AIRescueEngine()
        prova_resgatada = ai_engine.estruturar_prova_com_ia(texto_completo)

        # Se a IA conseguir estruturar as questões, assume o resultado e preserva as imagens mapeadas
        if len(prova_resgatada.questoes) > 0:
            print(f"[PAS/A Engine] AI-Rescue resgatou {len(prova_resgatada.questoes)} questões com sucesso.")
            # Reassocia as imagens extraídas pelo parser local às questões resgatadas, se aplicável
            for idx, q in enumerate(prova_resgatada.questoes):
                if idx < len(prova_parseada.questoes) and prova_parseada.questoes[idx].imagens:
                    q.imagens = prova_parseada.questoes[idx].imagens
            return prova_resgatada

    return prova_parseada