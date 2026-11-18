from typing import List, Optional
from pydantic import BaseModel, Field


class AlternativaSchema(BaseModel):
    letra: str
    texto: str


class TabelaQuestaoSchema(BaseModel):
    linhas: List[List[str]] = Field(default_factory=list)


class QuestaoSchema(BaseModel):
    numero: int
    enunciado: str
    tabela: Optional[TabelaQuestaoSchema] = None
    alternativas: List[AlternativaSchema] = Field(default_factory=list)


class ProvaSchema(BaseModel):
    titulo: str = "Avaliação do Colégio Santo Antônio"
    tipo_avaliacao: str = "mensal"
    disciplina: str = "DISCIPLINA"
    professor: str = "PROFESSOR"
    turma: str = "1º Ano"
    bimestre: str = "4º Bimestre"
    chamada: str = "1ª Chamada"
    is_adaptada: bool = False
    exibir_gabarito: bool = True
    nome_arquivo_original: Optional[str] = None
    questoes: List[QuestaoSchema] = Field(default_factory=list)