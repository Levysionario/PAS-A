import os
import json
from google import genai
from google.genai import types
from app.schemas.prova import ProvaSchema, QuestaoSchema, AlternativaSchema

class AIRescueEngine:
    def __init__(self):
        self.api_key = os.getenv("GEMINI_API_KEY")
        self.client = genai.Client(api_key=self.api_key) if self.api_key else None

    def estruturar_prova_com_ia(self, texto_bruto_prova: str) -> ProvaSchema:
        if not self.client:
            print("[AI-Rescue] API Key não configurada no ambiente (GEMINI_API_KEY).")
            return ProvaSchema(titulo="Avaliação do Colégio Santo Antônio", questoes=[])

        prompt = f"""
        Você é um Extrator e Estruturador Estrito de Provas do Colégio Santo Antônio.
        Sua única função é extrair e organizar o conteúdo de um documento Word (.docx) enviado por um professor para a estrutura JSON definida.

        REGRA DE OURO IMPRESCINDÍVEL:
        - REGRA DE FIDELIDADE ZERO ALTERAÇÃO: Você NUNCA deve alterar, resumir, reescrever, corrigir, adicionar ou remover qualquer palavra, frase ou pontuação do texto original da prova. Transcreva o enunciado e as alternativas EXATAMENTE como foram escritos pelo professor.

        REGRAS DE ESTRUTURAÇÃO DO COLÉGIO SANTO ANTÔNIO:
        1. IGNORAR CABEÇALHOS E LIXO DE PÁGINA: Ignore campos de identificação como "Aluno (a):", "Professor:", "DATA:", "NOTA:", títulos como "Avaliação Mensal", "Atividade" e tabelas de gabarito/bolhas de resposta.
        2. IDENTIFICAÇÃO DE QUESTÕES:
           - Localize os identificadores de questão (ex.: "Questão 1", "Questão 2.", "1-", "2-").
           - Mantenha a numeração original da questão no campo 'numero'.
        3. QUESTÕES OBJETIVAS (MÚLTIPLA ESCOLHA):
           - Se a questão possui opções (sejam 3, 4 ou 5 alternativas precedidas por a), b), c), d), e) ou A), B), C)...), cada opção deve ser mapeada individualmente na lista 'alternativas' com a 'letra' e o 'texto' EXATO da alternativa.
        4. QUESTÕES DISCURSIVAS / ABERTAS / PRODUÇÃO TEXTUAL:
           - Se a questão for de resposta escrita, cálculos ou redação (sem opções de a/b/c/d/e), coloque todo o comando/orientações no campo 'enunciado' e deixe a lista 'alternativas' estritamente vazia ([]).
        5. TEXTOS DE APOIO / CONTEXTUALIZADORES:
           - Se houver um texto motivador antes das questões (ex.: "Texto para as questões de 2 a 11: ..."), inclua o texto motivador integralmente no início do 'enunciado' da primeira questão do bloco correspondente.

        FORMATO DE SAÍDA EXIGIDO (JSON Estrito sem marcadores adicionais):
        {{
            "titulo": "Avaliação do Colégio Santo Antônio",
            "questoes": [
                {{
                    "numero": 1,
                    "enunciado": "Texto integral da pergunta 1...",
                    "alternativas": [
                        {{"letra": "a", "texto": "Texto exato da alternativa A"}},
                        {{"letra": "b", "texto": "Texto exato da alternativa B"}}
                    ]
                }},
                {{
                    "numero": 2,
                    "enunciado": "Explique com suas palavras o conceito de seleção natural...",
                    "alternativas": []
                }}
            ]
        }}

        TEXTO DA PROVA PARA EXTRAÇÃO:
        ---
        {texto_bruto_prova}
        ---
        """

        try:
            response = self.client.models.generate_content(
                model='gemini-1.5-flash',
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0
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
                        texto=str(alt.get("texto", "")).strip()
                    )
                    for alt in item.get("alternativas", [])
                ]
                questoes_convertidas.append(
                    QuestaoSchema(
                        numero=item.get("numero", len(questoes_convertidas) + 1),
                        enunciado=str(item.get("enunciado", "")).strip(),
                        alternativas=alts
                    )
                )

            titulo = dados.get("titulo", "Avaliação do Colégio Santo Antônio")
            return ProvaSchema(titulo=titulo, questoes=questoes_convertidas)

        except Exception as e:
            print(f"[AI-Rescue] Erro durante o processamento da IA: {e}")
            return ProvaSchema(titulo="Avaliação do Colégio Santo Antônio", questoes=[])