import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from app.schemas.prova import ProvaSchema

MAPA_BANNERS = {
    "acn": "assets/cabecalhos/ACN.png",
    "atividade": "assets/cabecalhos/atividade.png",
    "bimestral": "assets/cabecalhos/bimestral.png",
    "diagnostica": "assets/cabecalhos/diagnostica.png",
    "extra": "assets/cabecalhos/extra.png",
    "final": "assets/cabecalhos/final.png",
    "recuperacao": "assets/cabecalhos/final.png",
    "guia": "assets/cabecalhos/guia.png",
    "mensal": "assets/cabecalhos/mensal.png",
    "ta": "assets/cabecalhos/TA.png",
    "td": "assets/cabecalhos/TD.png",
}

TIPOS_OFICIAIS_COM_NOTA = {
    "mensal",
    "bimestral",
    "acn",
    "final",
    "recuperacao",
    "diagnostica",
}


def _aplicar_fonte_padrao(run, negrito: bool = False, tamanho: int = 12, cor_rgb: tuple = None):
    run.font.name = "Arial Narrow"
    run.font.size = Pt(tamanho)
    run.bold = negrito
    if cor_rgb:
        run.font.color.rgb = RGBColor(*cor_rgb)
    # Garante Arial Narrow no XML (eastAsia / cs)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Arial Narrow")
    rFonts.set(qn("w:hAnsi"), "Arial Narrow")
    rFonts.set(qn("w:cs"), "Arial Narrow")


def _adicionar_borda_inferior_paragrafo(paragrafo, espessura: str = "12"):
    """Adiciona linha divisória sólida de 1.5pt (12 eighths of a point)."""
    pPr = paragrafo._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{espessura}" w:space="2" w:color="000000"/>'
        f"</w:pBdr>"
    )
    pPr.append(pBdr)


def _aplicar_fundo_vermelho_paragrafo(paragrafo, cor_hex: str = "FF3333"):
    """Aplica destaque marca-texto vermelho (#FF3333) para Prova Adaptada."""
    pPr = paragrafo._p.get_or_add_pPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{cor_hex}"/>')
    pPr.append(shd)


def _configurar_bordas_tabela(table, borda_externa: bool = False):
    tblPr = table._tbl.tblPr
    val = "single" if borda_externa else "none"
    sz = "12" if borda_externa else "0"
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def _configurar_borda_celula_nota(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'<w:left w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f'<w:right w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        f"</w:tcBorders>"
    )
    tcPr.append(borders)


def _inserir_cartao_resposta(doc: Document, total_questoes_objetivas: int):
    if total_questoes_objetivas <= 0:
        return

    colunas = 1 if total_questoes_objetivas <= 10 else (2 if total_questoes_objetivas <= 20 else 3)
    linhas_por_col = (total_questoes_objetivas + colunas - 1) // colunas

    table = doc.add_table(rows=linhas_por_col, cols=colunas)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _configurar_bordas_tabela(table, borda_externa=False)

    for idx in range(total_questoes_objetivas):
        col_idx = idx // linhas_por_col
        row_idx = idx % linhas_por_col
        cell = table.cell(row_idx, col_idx)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)

        num_str = f"{idx + 1}- "
        run = p.add_run(f"{num_str}A ○   B ○   C ○   D ○   E ○")
        _aplicar_fonte_padrao(run, negrito=True, tamanho=11)


def gerar_documento_oficial(prova: ProvaSchema, output_path: str) -> str:
    doc = Document()

    # Margens A4 padrão (1.2 cm laterais, 1.0 cm superior/inferior)
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.0)
        section.bottom_margin = Cm(1.0)
        section.left_margin = Cm(1.2)
        section.right_margin = Cm(1.2)

    tipo_chave = (prova.tipo_avaliacao or "mensal").strip().lower()
    eh_prova_oficial = tipo_chave in TIPOS_OFICIAIS_COM_NOTA

    # 1. Banner Institucional (assets/cabecalhos/*.png)
    caminho_banner = MAPA_BANNERS.get(tipo_chave, "assets/cabecalhos/mensal.png")
    p_banner = doc.add_paragraph()
    p_banner.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_banner.paragraph_format.space_before = Pt(0)
    p_banner.paragraph_format.space_after = Pt(2)
    if os.path.exists(caminho_banner):
        run_img = p_banner.add_run()
        run_img.add_picture(caminho_banner, width=Cm(18.6))
    else:
        run_fallback = p_banner.add_run("COLÉGIO SANTO ANTÔNIO")
        _aplicar_fonte_padrao(run_fallback, negrito=True, tamanho=14)

    _adicionar_borda_inferior_paragrafo(p_banner)

    # 2. Disciplina + Destaque Vermelho (#FF3333) se Prova Adaptada
    p_disc = doc.add_paragraph()
    p_disc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_disc.paragraph_format.space_before = Pt(3)
    p_disc.paragraph_format.space_after = Pt(4)

    texto_disc = (prova.disciplina or "DISCIPLINA").strip().upper()
    if prova.is_adaptada:
        _aplicar_fundo_vermelho_paragrafo(p_disc, "FF3333")
        run_disc = p_disc.add_run(f"  {texto_disc} (ADAPTADA)  ")
        _aplicar_fonte_padrao(run_disc, negrito=True, tamanho=13, cor_rgb=(255, 255, 255))
    else:
        run_disc = p_disc.add_run(texto_disc)
        _aplicar_fonte_padrao(run_disc, negrito=True, tamanho=13)

    # 3. Bloco de Identificação (Cabeçalho em Negrito + Caixa de Nota Ampliada condicional)
    cols_cabecalho = 2 if eh_prova_oficial else 1
    tabela_cab = doc.add_table(rows=1, cols=cols_cabecalho)
    tabela_cab.alignment = WD_TABLE_ALIGNMENT.CENTER
    _configurar_bordas_tabela(tabela_cab, borda_externa=False)

    cell_dados = tabela_cab.cell(0, 0)
    cell_dados.width = Cm(15.8 if eh_prova_oficial else 18.6)
    cell_dados.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p_aluno = cell_dados.paragraphs[0]
    p_aluno.paragraph_format.space_before = Pt(1)
    p_aluno.paragraph_format.space_after = Pt(4)
    linha_underline = "_" * 66 if eh_prova_oficial else "_" * 80
    r_aluno = p_aluno.add_run(f"Aluno (a): {linha_underline}   Nº _____")
    _aplicar_fonte_padrao(r_aluno, negrito=True, tamanho=12)

    p_meta = cell_dados.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(0)
    p_meta.paragraph_format.space_after = Pt(2)

    partes_meta = [
        (prova.turma or "1º Ano").strip(),
        (prova.bimestre or "4º Bimestre").strip(),
    ]
    if eh_prova_oficial:
        partes_meta.append((prova.chamada or "1ª Chamada").strip())
    partes_meta.append(f"Prof. {(prova.professor or '').strip()}")
    partes_meta.append("DATA: ____/____/2026")

    r_meta = p_meta.add_run("      ".join(partes_meta))
    _aplicar_fonte_padrao(r_meta, negrito=True, tamanho=12)

    if eh_prova_oficial:
        cell_nota = tabela_cab.cell(0, 1)
        cell_nota.width = Cm(2.8)
        cell_nota.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _configurar_borda_celula_nota(cell_nota)

        p_nota = cell_nota.paragraphs[0]
        p_nota.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_nota.paragraph_format.space_before = Pt(2)
        p_nota.paragraph_format.space_after = Pt(18)  # Altura ampliada conforme pedido da gráfica
        r_nota = p_nota.add_run("NOTA:")
        _aplicar_fonte_padrao(r_nota, negrito=True, tamanho=11)

    p_div_bottom = doc.add_paragraph()
    p_div_bottom.paragraph_format.space_before = Pt(0)
    p_div_bottom.paragraph_format.space_after = Pt(4)
    _adicionar_borda_inferior_paragrafo(p_div_bottom)

    # 4. Cartão-Resposta (Apenas para Provas Oficiais com questões objetivas)
    if eh_prova_oficial and prova.exibir_gabarito:
        qtd_objetivas = sum(1 for q in prova.questoes if len(q.alternativas) > 0)
        _inserir_cartao_resposta(doc, qtd_objetivas)

    # 5. Corpo das Questões (Regras Estritas: Arial Narrow 12, linha em branco, Questão X isolado em negrito, alternativas minúsculas)
    for q in prova.questoes:
        # Linha em branco obrigatória antes de todos os "Questão X"
        p_vazio = doc.add_paragraph()
        p_vazio.paragraph_format.space_before = Pt(0)
        p_vazio.paragraph_format.space_after = Pt(0)
        p_vazio.paragraph_format.line_spacing = 1.15
        _aplicar_fonte_padrao(p_vazio.add_run(""))

        # "Questão X" sozinho numa linha dedicada, antes do enunciado, sempre em negrito
        p_q_titulo = doc.add_paragraph()
        p_q_titulo.paragraph_format.space_before = Pt(0)
        p_q_titulo.paragraph_format.space_after = Pt(0)
        p_q_titulo.paragraph_format.line_spacing = 1.15
        p_q_titulo.paragraph_format.keep_with_next = True
        _aplicar_fonte_padrao(p_q_titulo.add_run(f"Questão {q.numero}"), negrito=True, tamanho=12)

        # Enunciado fiel ao original (preservando quebras de linha e sem recuo)
        for linha in q.enunciado.split("\n"):
            p_enunc = doc.add_paragraph()
            p_enunc.paragraph_format.space_before = Pt(0)
            p_enunc.paragraph_format.space_after = Pt(0)
            p_enunc.paragraph_format.line_spacing = 1.15
            p_enunc.paragraph_format.first_line_indent = Pt(0)
            _aplicar_fonte_padrao(p_enunc.add_run(linha), negrito=False, tamanho=12)

        # Alternativas sempre em letras minúsculas, texto puro sem marcadores automáticos
        for alt in q.alternativas:
            letra_min = alt.letra.strip().lower().replace(")", "").replace(".", "")
            p_alt = doc.add_paragraph()
            p_alt.paragraph_format.space_before = Pt(0)
            p_alt.paragraph_format.space_after = Pt(0)
            p_alt.paragraph_format.line_spacing = 1.15
            p_alt.paragraph_format.first_line_indent = Pt(0)
            _aplicar_fonte_padrao(p_alt.add_run(f"{letra_min}) {alt.texto}"), negrito=False, tamanho=12)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path

class ProvaDocxBuilder:
    def __init__(self, prova: ProvaSchema = None):
        self.prova = prova

    def build(self, output_path: str, prova: ProvaSchema = None) -> str:
        alvo = prova or self.prova
        return gerar_documento_oficial(alvo, output_path)

    def gerar(self, output_path: str, prova: ProvaSchema = None) -> str:
        alvo = prova or self.prova
        return gerar_documento_oficial(alvo, output_path)