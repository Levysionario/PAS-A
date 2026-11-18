import os
import re
import json
import time
from google import genai
from google.genai import types
from app.schemas.prova import ProvaSchema, QuestaoSchema, AlternativaSchema, TabelaQuestaoSchema


class AIRescueEngine:
    def __init__(self):
        self.api_key = os.getenv("GEMINI_API_KEY")
        self.client = (
            genai.Client(
                api_key=self.api_key,
                http_options=types.HttpOptions(timeout=25_000),
            )
            if self.api_key
            else None
        )
        self.ultimo_erro: str | None = None
        self.modelos_fallback = [
            "gemini-3.8-flash",
            "gemini-3.8-flash-lite",
        ]

    def _extracao_contingencia_local(self, texto_bruto: str) -> ProvaSchema:
        """
        Contingência local determinística caso a API do Gemini esteja indisponível ou lenta.
        Extrai questões com ou sem numeração, objetivas ou discursivas, e tabelas [TABELA]...[/TABELA].
        """
        linhas_raw = [l.strip() for l in texto_bruto.splitlines() if l.strip()]

        disciplina = "DISCIPLINA"
        professor = "PROFESSOR"
        turma = "1º Ano"
        tipo_avaliacao = "mensal"

        for l in linhas_raw[:15]:
            l_upper = l.upper()
            if "PROF." in l_upper:
                m_prof = re.search(r"Prof\.\s*([A-ZÀ-Ú\s]+?)(?:\s+DATA:|\||$)", l, re.IGNORECASE)
                if m_prof:
                    professor = m_prof.group(1).strip()
            if any(t in l_upper for t in ["1º ANO", "2º ANO", "3º ANO", "6º ANO", "7º ANO", "8º ANO", "9º ANO"]):
                m_turma = re.search(r"([1-9]º\s*Ano(?:\s*[–-]\s*[A-ZÀ-Ú]+)?)", l, re.IGNORECASE)
                if m_turma:
                    turma = m_turma.group(1).strip()
            if "ACN" in l_upper:
                tipo_avaliacao = "acn"
            elif "BIMESTRAL" in l_upper:
                tipo_avaliacao = "bimestral"

            for disc_nome in [
                "BIOLOGIA 1", "BIOLOGIA 2", "BIOLOGIA", "FÍSICA 1", "FÍSICA 2", "FÍSICA",
                "QUÍMICA 1", "QUÍMICA 2", "QUÍMICA", "MATEMÁTICA 1", "MATEMÁTICA 2", "MATEMÁTICA",
                "HISTÓRIA", "GEOGRAFIA", "PORTUGUÊS", "LÍNGUA PORTUGUESA", "INGLÊS", "ESPANHOL",
                "FILOSOFIA", "SOCIOLOGIA", "LITERATURA", "REDAÇÃO", "ARTES", "EDUCAÇÃO FÍSICA",
            ]:
                if disc_nome in l_upper and len(l) <= 45:
                    limpa = re.sub(r"NOTA\s*:?", "", l, flags=re.IGNORECASE).replace("|", "").strip()
                    if limpa:
                        disciplina = limpa.upper()
                    break

        padrao_q_num = re.compile(r"^(?:Quest[ãa]o\s*)?0?(\d{1,2})[\.\)\-–:]\s*(.+)$", re.IGNORECASE)
        padrao_alt = re.compile(r"^\(?([a-eA-E])[\)\.\-:]\s*(.+)$")

        questoes: list[QuestaoSchema] = []
        buffer_enunciado: list[str] = []
        alternativas_atuais: list[AlternativaSchema] = []
        tabela_atual: TabelaQuestaoSchema | None = None
        dentro_tabela = False
        linhas_tabela_temp: list[list[str]] = []
        numero_atual = 1

        def _fechar_questao():
            nonlocal numero_atual, buffer_enunciado, alternativas_atuais, tabela_atual
            if buffer_enunciado or alternativas_atuais:
                enunc_texto = "\n".join(buffer_enunciado).strip()
                if enunc_texto:
                    questoes.append(
                        QuestaoSchema(
                            numero=numero_atual,
                            enunciado=enunc_texto,
                            tabela=tabela_atual,
                            alternativas=alternativas_atuais,
                        )
                    )
                    numero_atual += 1
            buffer_enunciado = []
            alternativas_atuais = []
            tabela_atual = None

        for linha in linhas_raw:
            if linha == "[TABELA]":
                dentro_tabela = True
                linhas_tabela_temp = []
                continue
            if linha == "[/TABELA]":
                dentro_tabela = False
                if linhas_tabela_temp:
                    tabela_atual = TabelaQuestaoSchema(linhas=linhas_tabela_temp)
                continue
            if dentro_tabela:
                linhas_tabela_temp.append([c.strip() for c in linha.split("|")])
                continue

            l_upper = linha.upper()
            if any(
                ign in l_upper
                for ign in [
                    "ALUNO (A):",
                    "ALUNO(A):",
                    "NÚMERO DE REGISTRO",
                    "NUMERO DE REGISTRO",
                    "A B C D E",
                    "COLÉGIO SANTO ANTÔNIO",
                    "DATA: ____",
                ]
            ):
                continue
            if re.match(r"^(NOTA\s*:?\s*\|?\s*)+", l_upper) and len(linha) < 50:
                continue

            m_q = padrao_q_num.match(linha)
            if m_q and not re.match(r"^\d{1,2}\s+A\s+B\s+C", linha, re.IGNORECASE):
                if buffer_enunciado or alternativas_atuais:
                    _fechar_questao()
                numero_atual = int(m_q.group(1))
                buffer_enunciado.append(m_q.group(2).strip())
                continue

            m_alt = padrao_alt.match(linha)
            if m_alt:
                letra = m_alt.group(1).lower()
                if letra == "a" and len(alternativas_atuais) >= 2:
                    novo_enunc = []
                    if buffer_enunciado and len(buffer_enunciado) > 1:
                        novo_enunc = [buffer_enunciado.pop()]
                    _fechar_questao()
                    buffer_enunciado = novo_enunc
                alternativas_atuais.append(AlternativaSchema(letra=letra, texto=m_alt.group(2).strip()))
                continue

            if len(alternativas_atuais) >= 4:
                _fechar_questao()
                buffer_enunciado.append(linha)
            elif alternativas_atuais:
                alternativas_atuais[-1].texto += " " + linha
            else:
                buffer_enunciado.append(linha)

        _fechar_questao()

        return ProvaSchema(
            tipo_avaliacao=tipo_avaliacao,
            disciplina=disciplina,
            professor=professor,
            turma=turma,
            questoes=questoes,
        )

    def estruturar_prova_com_ia(self, texto_bruto_prova: str) -> ProvaSchema:
        self.ultimo_erro = None
        if not self.client:
            self.ultimo_erro = "GEMINI_API_KEY não configurada no servidor."
            print(f"[AI-Rescue] {self.ultimo_erro}")
            return self._extracao_contingencia_local(texto_bruto_prova)

        prompt = f"""
        Você é o motor de extração estrutural do sistema gráfico PAS/A do Colégio Santo Antônio.
        Os professores enviam arquivos .docx totalmente fora do padrão: sem numeração nas questões (sem 01, 1, 1-), com numeração automática do Word removida, com alternativas em maiúsculo (A, B, C, D, E), com cabeçalhos antigos colados errado, com gabarito pintado de vermelho, com listas marcadas e com tabelas dentro dos enunciados.

        SUA MISSÃO É SEGMENTAR E EXTRAIR TODO O CONTEÚDO EM JSON SEGUINDO ESTAS REGRAS INVIOLÁVEIS:

        1. IDENTIFICAÇÃO SEMÂNTICA DE QUESTÕES SEM NÚMERO:
           - Mesmo que o texto NÃO tenha "01.", "1)", "1 -" ou "Questão 1" antes de uma pergunta, identifique cada questão pela estrutura semântica (texto-base/comando seguido de alternativas, ou bloco de comando discursivo separado) e numere sequencialmente: 1, 2, 3, 4...
           - Remova qualquer prefixo numérico antigo do início do enunciado (ex: se vier "01. Em um ecossistema...", o campo "enunciado" deve começar direto em "Em um ecossistema...").

        2. FIDELIDADE 1:1 ABSOLUTA NO TEXTO:
           - NUNCA reescreva, resuma, corrija erros de português, altere palavras ou mude a pontuação do enunciado ou das alternativas. Transcreva o texto do professor exatamente como foi escrito.
           - Se houver tópicos dentro do enunciado (ex: "• Situação I:"), preserve o texto exato sem transformar em markdown (* ou -).

        3. ALTERNATIVAS SEMPRE EM MINÚSCULO E SEM PREFIXO DUPLICADO:
           - Extraia cada alternativa com "letra" sempre em minúsculo ("a", "b", "c", "d", "e") e "texto" contendo apenas o conteúdo da alternativa (sem repetir "A)" ou "a)" dentro de "texto").
           - Se a questão for discursiva/aberta (sem alternativas), retorne "alternativas": [].

        4. TABELAS DENTRO DE QUESTÕES:
           - No texto abaixo, tabelas reais do Word aparecem delimitadas por [TABELA] e [/TABELA], com colunas separadas por " | ".
           - Quando uma questão contiver uma tabela, extraia os dados dela no campo "tabela": {{"linhas": [["col1", "col2"], ["val1", "val2"]]}} e mantenha no "enunciado" o texto antes e depois da tabela (não repita os dados brutos da tabela dentro do texto corrido do enunciado). Se a questão não tiver tabela, retorne "tabela": null.

        5. IGNORAR LIXO DE CABEÇALHO ANTIGO:
           - Ignore linhas de cartão-resposta antigo ("Número de Registro", "Gabarito", "A B C D E"), "Aluno (a):", "NOTA:" e nomes de escola.
           - Extraia os metadados do topo se existirem: "disciplina", "professor", "turma", "bimestre", "chamada" e "tipo_avaliacao" ("mensal", "bimestral", "acn", "atividade", "diagnostica", "extra", "final", "guia", "ta", "td").

        Retorne EXATAMENTE um JSON válido neste formato:
        {{
            "tipo_avaliacao": "acn",
            "disciplina": "BIOLOGIA 1",
            "professor": "LUCA DE LUNA",
            "turma": "3º Ano - MÉDIO",
            "bimestre": "4º Bimestre",
            "chamada": "1ª Chamada",
            "is_adaptada": false,
            "questoes": [
                {{
                    "numero": 1,
                    "enunciado": "Texto exatamente igual ao original sem o número inicial",
                    "tabela": null,
                    "alternativas": [
                        {{"letra": "a", "texto": "texto exato da alternativa a"}},
                        {{"letra": "b", "texto": "texto exato da alternativa b"}}
                    ]
                }}
            ]
        }}

        Conteúdo extraído do arquivo do professor:
        ---
        {texto_bruto_prova}
        ---
        """

        response = None
        erros_coletados = []

        for modelo in self.modelos_fallback:
            for tentativa in range(1, 3):
                try:
                    response = self.client.models.generate_content(
                        model=modelo,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.0,
                        ),
                    )
                    if response and response.text:
                        break
                except Exception as e:
                    erro_str = str(e)
                    erros_coletados.append(f"[{modelo} t{tentativa}]: {erro_str}")
                    print(f"[AI-Rescue] Tentativa {tentativa} falhou no modelo {modelo}: {erro_str}")

                    if any(cod in erro_str for cod in ["503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"]):
                        time.sleep(1.0)
                        continue
                    else:
                        break

            if response and response.text:
                break

        if not response or not response.text:
            self.ultimo_erro = " | ".join(erros_coletados) if erros_coletados else "Sem resposta da API Gemini."
            print("[AI-Rescue] API indisponível ou timeout. Usando extrator de contingência local...")
            return self._extracao_contingencia_local(texto_bruto_prova)

        try:
            raw_json = response.text.strip()
            if raw_json.startswith("```"):
                raw_json = raw_json.split("\n", 1)[1]
            if raw_json.endswith("```"):
                raw_json = raw_json.rsplit("\n", 1)[0]
            raw_json = raw_json.strip()

            dados = json.loads(raw_json)

            questoes_convertidas = []
            for idx, item in enumerate(dados.get("questoes", []), start=1):
                alts = [
                    AlternativaSchema(
                        letra=str(alt.get("letra", "")).strip().lower().replace(")", "").replace(".", ""),
                        texto=str(alt.get("texto", "")).strip(),
                    )
                    for alt in item.get("alternativas", [])
                ]

                tabela_obj = None
                raw_tabela = item.get("tabela")
                if isinstance(raw_tabela, dict) and raw_tabela.get("linhas"):
                    linhas_limpas = [
                        [str(cel).strip() for cel in linha]
                        for linha in raw_tabela["linhas"]
                        if isinstance(linha, list) and any(str(c).strip() for c in linha)
                    ]
                    if linhas_limpas:
                        tabela_obj = TabelaQuestaoSchema(linhas=linhas_limpas)

                questoes_convertidas.append(
                    QuestaoSchema(
                        numero=idx,
                        enunciado=str(item.get("enunciado", "")).strip(),
                        tabela=tabela_obj,
                        alternativas=alts,
                    )
                )

            if not questoes_convertidas:
                self.ultimo_erro = "IA retornou lista de questões vazia."
                return self._extracao_contingencia_local(texto_bruto_prova)

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
            self.ultimo_erro = f"Erro ao interpretar JSON da IA: {e}"
            print(f"[AI-Rescue] {self.ultimo_erro}")
            return self._extracao_contingencia_local(texto_bruto_prova)