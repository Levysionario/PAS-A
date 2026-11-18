import os
import json
from google import genai
from google.genai import types
from app.schemas.prova import ProvaSchema, QuestaoSchema, AlternativaSchema

class AIRescueEngine:
    def __init__(self):
        self.api_key = os.getenv("GEMINI_API_KEY")
        self.client = genai.Client(api_key=self.api_key) if self.api_key else None
        self.ultimo_erro: str | None = None

    def estruturar_prova_com_ia(self, texto_bruto_prova: str) -> ProvaSchema:
        self.ultimo_erro = None
        if not self.client:
            self.ultimo_erro = "GEMINI_API_KEY não configurada no servidor."
            print(f"[AI-Rescue] {self.ultimo_erro}")
            return ProvaSchema(questoes=[])

        prompt = f"""
        Você é um extrator estrutural estrito de provas do Colégio Santo Antônio.
        Sua única tarefa é separar o texto abaixo em metadados, questões e alternativas no formato JSON.

        REGRAS ABSOLUTAS E INVIOLÁVEIS:
        1. FIDELIDADE 1:1 ABSOLUTA: NUNCA escreva, invente, resuma, remova ou sequer corrija erros de digitação, ortografia ou pontuação da prova original. Transcreva os enunciados e alternativas ESTRITAMENTE IGUAIS ao original.
        2. SEM MARCADORES AUTOMÁTICOS: Se houver tópicos ou listas dentro do enunciado, mantenha o caractere literal original (ex: "· ") como texto simples. NUNCA converta para markdown (* ou -).
        3. ALTERNATIVAS SEMPRE EM MINÚSCULO: As letras das alternativas devem ser SEMPRE minúsculas ("a", "b", "c", "d", "e"), mesmo que no original estejam em maiúsculo.
        4. QUESTÕES DISCURSIVAS/ABERTAS: Se a questão não tiver alternativas, retorne "alternativas": [].
        5. METADADOS: Extraia disciplina, professor, turma e tipo_avaliacao (valores válidos: "mensal", "bimestral", "acn", "atividade", "diagnostica", "extra", "final", "guia", "ta", "td") se constarem no topo do arquivo. Caso não constem, use os valores padrão.

        Retorne EXATAMENTE um JSON neste formato:
        {{
            "tipo_avaliacao": "mensal",
            "disciplina": "FÍSICA",
            "professor": "NOME DO PROFESSOR",
            "turma": "1º Ano",
            "bimestre": "4º Bimestre",
            "chamada": "1ª Chamada",
            "is_adaptada": false,
            "questoes": [
                {{
                    "numero": 1,
                    "enunciado": "Texto exatamente igual ao original",
                    "alternativas": [
                        {{"letra": "a", "texto": "texto exatamente igual ao original"}},
                        {{"letra": "b", "texto": "texto exatamente igual ao original"}}
                    ]
                }}
            ]
        }}

        Texto original:
        ---
        {texto_bruto_prova}
        ---
        """

        try:
            response = self.client.models.generate_content(
                model="gemini-3.8-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )

            raw_json = response.text.strip()
            if raw_json.startswith("```"):
                raw_json = raw_json.split("\n", 1)[1]
            if raw_json.endswith("```"):
                raw_json = raw_json.rsplit("\n", 1)[0]
            raw_json = raw_json.strip()

            dados = json.loads(raw_json)

            questoes_convertidas = []
            for item in dados.get("questoes", []):
                alts = [
                    AlternativaSchema(
                        letra=str(alt.get("letra", "")).strip().lower().replace(")", "").replace(".", ""),
                        texto=str(alt.get("texto", "")).strip(),
                    )
                    for alt in item.get("alternativas", [])
                ]
                questoes_convertidas.append(
                    QuestaoSchema(
                        numero=item.get("numero", len(questoes_convertidas) + 1),
                        enunciado=str(item.get("enunciado", "")).strip(),
                        alternativas=alts,
                    )
                )

            return ProvaSchema(
                tipo_avaliacao=str(dados.get("tipo_avaliacao", "mensal")).strip().lower(),
                disciplina=str(dados.get("disciplina", "DISCIPLINA")).strip().upper(),
                professor=str(dados.get("professor", "PROFESSOR")).strip(),
                turma=str(dados.get("turma", "1º Ano")).strip(),
                bimestre=str(dados.get("bimestre", "4º Bimestre")).strip(),
                chamada=str(dados.get("chamada", "1ª Chamada")).strip(),
                is_adaptada=bool(dados.get("is_adaptada", False)),
                questoes=questoes_convertidas,
            )

        except Exception as e:
            self.ultimo_erro = str(e)
            print(f"[AI-Rescue] Erro ao processar resposta do Gemini: {e}")
            return ProvaSchema(questoes=[])