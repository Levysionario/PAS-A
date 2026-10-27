from pydantic import BaseModel
from typing import List, Optional

class AlternativaSchema(BaseModel):
    letra: str
    texto: str

class QuestaoSchema(BaseModel):
    numero: int
    enunciado: str
    imagens: List[str] = []
    alternativas: List[AlternativaSchema] = []

class ProvaSchema(BaseModel):
    titulo: str
    disciplina: Optional[str] = "Não informada"
    professor: Optional[str] = "Não informado"
    questoes: List[QuestaoSchema] = []