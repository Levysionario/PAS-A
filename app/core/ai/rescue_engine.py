import os
import json
from google import genai
from google.genai import types
from app.schemas.prova import ProvaSchema

class AIRescueEngine:
    def __init__(self):
        self.api_key = os.getenv("GEMINI_API_KEY")
        self.client = genai.Client(api_key=self.api_key) if self.api_key else None

    def estruturar_prova_com_ia(self, texto_bruto_prova: str) -> ProvaSchema:
        if not self.client:
            print("[AI-Rescue] API Key não configurada no ambiente (GEMINI_API_KEY).")
            return ProvaSchema(titulo="Avaliação do Colégio Santo Antônio", questoes=[])

        prompt = f"""
        Você é um assistente especializado em diagramação e estruturação de provas escolares.
        Abaixo está o texto extraído de um documento Word enviado por um professor. Esse documento pode conter tabelas desalinhadas, quebras de linha forçadas ou falta de numeração explícita.

        Sua missão:
        1. Identificar todas as questões, construindo seus respectivos enunciados e alternativas (A, B, C, D, E).
        2. Corrigir erros de digitação óbvios no texto das questões/alternativas se necessário.
        3. Ignorar cabeçalhos de escola, instruções de preenchimento de cartão e avisos genéricos.

        Texto da Prova:
        ---
        {texto_bruto_prova}
        ---
        """

        try:
            response = self.client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ProvaSchema,
                    temperature=0.1
                ),
            )
            
            dados_json = json.loads(response.text)
            return ProvaSchema(**dados_json)

        except Exception as e:
            print(f"[AI-Rescue] Erro durante o processamento com a API: {e}")
            return ProvaSchema(titulo="Avaliação do Colégio Santo Antônio", questoes=[])