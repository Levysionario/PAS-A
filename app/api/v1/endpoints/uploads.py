import os
import shutil
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from app.core.ai.rescue_engine import AIRescueEngine
from app.schemas.prova import ProvaSchema

router = APIRouter()

UPLOAD_DIR = "storage/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)


def _iterar_blocos_documento(doc: Document):
    """Itera sobre parágrafos e tabelas na ordem exata em que aparecem no corpo do Word."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("tbl"):
            yield Table(child, doc)


def _extrair_conteudo_estruturado_bruto(caminho_arquivo: str) -> str:
    """
    Extrai texto e tabelas na posição exata do documento.
    Ignora tabelas de cabeçalho/gabarito antigo e marca tabelas de questões com [TABELA]...[/TABELA].
    """
    doc = Document(caminho_arquivo)
    blocos = []

    for bloco in _iterar_blocos_documento(doc):
        if isinstance(bloco, Paragraph):
            texto = bloco.text.strip()
            if texto:
                blocos.append(texto)

        elif isinstance(bloco, Table):
            linhas_tabela = []
            for row in bloco.rows:
                celulas_unicas = []
                for cell in row.cells:
                    txt_cel = " ".join(p.text.strip() for p in cell.paragraphs if p.text and p.text.strip())
                    if txt_cel and txt_cel not in celulas_unicas:
                        celulas_unicas.append(txt_cel)
                if celulas_unicas:
                    linhas_tabela.append(" | ".join(celulas_unicas))

            if not linhas_tabela:
                continue

            texto_tabela_completo = "\n".join(linhas_tabela)
            upper_tab = texto_tabela_completo.upper()

            if "ALUNO (A):" in upper_tab or "PROF." in upper_tab:
                blocos.append(texto_tabela_completo)
            elif "A | B | C | D | E" in upper_tab or "NÚMERO DE REGISTRO" in upper_tab:
                continue
            else:
                blocos.append("[TABELA]\n" + texto_tabela_completo + "\n[/TABELA]")

    return "\n".join(blocos)


@router.post("/processar-docx", response_model=ProvaSchema)
def processar_documento(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Formato inválido. Envie um arquivo .docx")

    caminho_arquivo = os.path.join(UPLOAD_DIR, file.filename)

    with open(caminho_arquivo, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    texto_completo = _extrair_conteudo_estruturado_bruto(caminho_arquivo)
    if not texto_completo.strip():
        raise HTTPException(
            status_code=422,
            detail="O arquivo .docx enviado não possui texto legível (pode conter apenas imagens ou caixas de texto).",
        )

    print(f"[PAS/A Engine] Processando arquivo com IA Semântica: {file.filename}")
    ai_engine = AIRescueEngine()
    prova_estruturada = ai_engine.estruturar_prova_com_ia(texto_completo)

    if len(prova_estruturada.questoes) == 0:
        motivo = ai_engine.ultimo_erro or "Não foi possível identificar questões neste arquivo."
        raise HTTPException(status_code=422, detail=f"Falha na extração: {motivo}")

    prova_estruturada.nome_arquivo_original = file.filename
    return prova_estruturada


@router.post("/refazer-com-ia", response_model=ProvaSchema)
def refazer_com_ia(filename: str = Form(...)):
    caminho_arquivo = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(caminho_arquivo):
        raise HTTPException(
            status_code=404,
            detail="Arquivo original não encontrado no servidor. Envie o .docx novamente.",
        )

    texto_completo = _extrair_conteudo_estruturado_bruto(caminho_arquivo)

    print(f"[PAS/A Engine] Reestruturação manual por IA solicitada para: {filename}")
    ai_engine = AIRescueEngine()
    prova_resgatada = ai_engine.estruturar_prova_com_ia(texto_completo)

    if len(prova_resgatada.questoes) == 0:
        motivo = ai_engine.ultimo_erro or "A IA não conseguiu estruturar as questões deste arquivo."
        raise HTTPException(status_code=422, detail=f"Falha na extração: {motivo}")

    prova_resgatada.nome_arquivo_original = filename
    return prova_resgatada