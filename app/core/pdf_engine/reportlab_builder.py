import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Frame, 
    PageTemplate, FrameBreak, HRFlowable, Table, TableStyle
)
from reportlab.lib.units import inch
from app.schemas.prova import ProvaSchema

class ProvaPDFBuilder:
    def __init__(self, prova: ProvaSchema, output_pdf_path: str):
        self.prova = prova
        self.output_pdf_path = output_pdf_path

    def build(self) -> str:
        # Configuração da página Letter (215.9 x 279.4 mm) com margens reduzidas
        doc = SimpleDocTemplate(
            self.output_pdf_path,
            pagesize=letter,
            leftMargin=0.4 * inch,
            rightMargin=0.4 * inch,
            topMargin=0.4 * inch,
            bottomMargin=0.4 * inch
        )

        styles = getSampleStyleSheet()

        # Estilos do cabeçalho e questões
        title_style = ParagraphStyle(
            'HeaderTitle',
            parent=styles['Heading1'],
            fontSize=11,
            leading=13,
            fontName='Helvetica-Bold',
            textColor=colors.HexColor("#1A365D"),
            alignment=1 # Centralizado
        )

        subtitle_style = ParagraphStyle(
            'HeaderSub',
            parent=styles['Normal'],
            fontSize=8.5,
            leading=11,
            alignment=1
        )

        enunciado_style = ParagraphStyle(
            'EnunciadoStyle',
            parent=styles['Normal'],
            fontSize=9,
            leading=12,
            fontName='Helvetica-Bold',
            spaceAfter=4,
            textColor=colors.HexColor("#1A202C")
        )

        alt_style = ParagraphStyle(
            'AlternativaStyle',
            parent=styles['Normal'],
            fontSize=8.5,
            leading=11,
            leftIndent=8,
            spaceAfter=2,
            textColor=colors.HexColor("#2D3748")
        )

        story = []

        # --- CABEÇALHO INSTITUCIONAL ---
        header_data = [
            [
                Paragraph("<b>COLÉGIO SANTO ANTÔNIO</b><br/><font size=8>Educação e Tradição</font>", title_style),
            ],
            [
                Paragraph(f"<b>AVALIAÇÃO DE {self.prova.disciplina.upper()}</b> | Prof(a): {self.prova.professor}", subtitle_style)
            ],
            [
                Paragraph("<b>Aluno(a):</b> ____________________________________________________  <b>Data:</b> ___/___/______", subtitle_style)
            ]
        ]

        header_table = Table(header_data, colWidths=[7.2 * inch])
        header_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#EDF2F7")),
            ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E0")),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))

        story.append(header_table)
        story.append(Spacer(1, 8))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#2B6CB0"), spaceAfter=8))

        # --- CORPO DA PROVA (QUESTÕES) ---
        for q in self.prova.questoes:
            # Enunciado
            texto_enunciado = f"<b>{q.numero}.</b> {q.enunciado.replace('\n', '<br/>')}"
            story.append(Paragraph(texto_enunciado, enunciado_style))

            # Alternativas
            for alt in q.alternativas:
                texto_alt = f"<b>({alt.letra})</b> {alt.texto}"
                story.append(Paragraph(texto_alt, alt_style))

            story.append(Spacer(1, 6))

        # Compilação do PDF
        doc.build(story)
        return self.output_pdf_path