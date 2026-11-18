from typing import List, Optional
from pydantic import BaseModel, Field

class AlternativaSchema(BaseModel):
    letra: str
    texto: str

class QuestaoSchema(BaseModel):
    numero: int
    enunciado: str
    alternativas: List[AlternativaSchema] = Field(default_factory=list)

class ProvaSchema(BaseModel):
    titulo: str = "Avaliação do Colégio Santo Antônio"
    tipo_avaliacao: str = "mensal"  # mensal, bimestral, acn, atividade, diagnostica, extra, final, guia, ta, td
    disciplina: str = "DISCIPLINA"
    professor: str = "PROFESSOR"
    turma: str = "Xº ANO"
    bimestre: str = "Xº Bimestre"
    chamada: str = "Xª Chamada"
    is_adaptada: bool = False
    exibir_gabarito: bool = True
    nome_arquivo_original: Optional[str] = None
    questoes: List[QuestaoSchema] = Field(default_factory=list)