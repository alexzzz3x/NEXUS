import os
import re
import json
import threading
import tkinter as tk
import customtkinter as ctk
import requests
from utils_os import obter_caminho_config, obter_diretorio_dados
from utils_skills import obter_gerenciador_skills

CAMINHO_CONFIG = obter_caminho_config("config_chat.txt")

MODELOS_OPENROUTER = [
    {"id": "local/motor-integrado", "nome": "⚡ Motor Inteligente Local (Sem API / Gratuito)"},
    {"id": "deepseek/deepseek-chat:free", "nome": "DeepSeek V3 (Free)"},
    {"id": "meta-llama/llama-3.3-70b-instruct:free", "nome": "Llama 3.3 70B (Free)"},
    {"id": "google/gemini-2.0-flash-exp:free", "nome": "Gemini 2.0 Flash (Free)"},
    {"id": "mistralai/mistral-small-24b-instruct-2501:free", "nome": "Mistral Small 24B (Free)"},
    {"id": "google/gemma-3-27b-it:free", "nome": "Gemma 3 27B (Free)"},
    {"id": "deepseek/deepseek-r1:free", "nome": "DeepSeek R1 (Free)"},
    {"id": "qwen/qwen-2.5-72b-instruct", "nome": "Qwen 2.5 72B"},
    {"id": "xiaomi/mimo-v2.5", "nome": "MiMo V2.5"},
]

SYSTEM_PROMPT = """Você é o Hub Assistant AI, o assistente virtual e agente orquestrador do Conecta Hub / N.E.X.U.S, sistema de gerenciamento de estoque, etiquetas e automação SAP da Conecta Empreendimentos.
Responda sempre em português brasileiro de forma amigável, clara e objetiva.
Você é um agente com Skills/Ferramentas conectadas diretamente aos módulos do aplicativo:
1. BANDEIRADA: Quando o usuário pedir para 'bandeirar' ou fizer menção a bandeirada de equipes/recursos (ex: 'bandeira o csc300', 'bandeirar csc300, csc301'), chame a ferramenta 'executar_bandeirada'.
2. EXTRAÇÃO OSME: Quando o usuário pedir para extrair OSME de equipes (ex: 'extrair osme csc300', 'osme do csc300'), chame a ferramenta 'extrair_osme'.
3. CONSULTAR MIGO / CLANDESTINAS: Quando o usuário pedir consulta de MIGO ou IW39 de clandestinas (ex: 'consultar migo rcl001', 'migo rcl002'), chame a ferramenta 'consultar_migo'.
4. TRANSFERÊNCIA OSME: Quando o usuário pedir transferência de OSME (Z62/Z61), chame a ferramenta 'transferir_osme'.
5. MUDANÇA DE RECURSO: Quando o usuário pedir para mudar recurso ou transferir de uma equipe para outra (ex: 'mudar recurso csc395 para cvc396', 'saldo origem csc395'), chame 'mudar_recurso'.
6. CONFERÊNCIA: Quando o usuário pedir para conferir ou comparar seriais, chame 'abrir_conferencia'.
7. FICHAS: Quando o usuário pedir fichas ou dados de RE, chame 'abrir_fichas'.
8. SAÍDA PARA OSME: Quando o usuário pedir saída para OSME, baixa de materiais ou processar materiais/seriais na OSME de uma equipe (ex: 'Processa na osme do cvc395 330702 50m', 'saída osme csc300 seriais 18398041 a 18398045'), chame a ferramenta 'saida_osme'. A ferramenta prepara a tela e dados para conferência prévia.
10. NAVEGAÇÃO: Se o usuário pedir para abrir ou ir para qualquer aba, chame 'abrir_modulo'.
11. MEDIDOR / IQ09: Quando o usuário pedir para pesquisar, puxar, buscar ou consultar medidores ou seriais no SAP / IQ09 (ex: 'pesquise o medidor 18398041', 'puxar seriais 18400025 a 18400028', 'iq09 18398041 a 18398045', 'consultar seriais ...'), chame a ferramenta 'consultar_iq09' passando os seriais ou o intervalo como parâmetro.
Sempre priorize chamar as ferramentas para executar ações diretamente na interface do usuário."""

TOOLS = [
    {
        "functionDeclarations": [
            {
                "name": "saida_osme",
                "description": "Prepara a Saída para OSME no módulo de Saída OSME para uma equipe, validando seriais e calculando quantidades consolidadas para conferência prévia do usuário antes de processar no SAP",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "equipe": {"type": "string", "description": "Código da equipe/recurso (ex: CVC395, CSC300)"},
                        "seriais": {"type": "string", "description": "Seriais ou intervalo de seriais (ex: '18398041 a 18398050')"},
                        "material": {"type": "string", "description": "Código do material não-serializado (ex: '330702', '510388')"},
                        "quantidade": {"type": "string", "description": "Quantidade do material não-serializado (ex: '50m', '50', '20')"}
                    },
                    "required": ["equipe"]
                }
            },
            {
                "name": "contagem_carro",
                "description": "Executa ou abre a Contagem de Carro (balanço de medidores no veículo da equipe com auditoria ZMCSE01, localização da ordem mais recente na IW39, estorno MIGO Z62 e saída MIGO Z61)",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "equipe": {"type": "string", "description": "Código da equipe/recurso (ex: CSC300)"},
                        "seriais": {"type": "string", "description": "Lista de seriais ou intervalo opcional"},
                        "acao": {"type": "string", "enum": ["abrir", "osme", "auditar", "executar"], "description": "Ação a realizar: 'abrir', 'osme' (puxar da OSME), 'auditar' (ZMCSE01), 'executar' (ciclo completo)"}
                    },
                    "required": ["equipe"]
                }
            },
            {
                "name": "executar_saida_osme",
                "description": "Executa o lançamento automático de Saída OSME no SAP MIGO Z61 após conferência ou confirmação do usuário",
                "parameters": {
                    "type": "object",
                    "properties": {}
                }
            },
            {
                "name": "executar_bandeirada",
                "description": "Executa o processo de Bandeirada no SAP para uma ou mais equipes/recursos (ex: CSC300, CDC410)",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "recursos": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Lista com códigos de recursos/equipes (ex: ['CSC300', 'CDC410'])"
                        }
                    },
                    "required": ["recursos"]
                }
            },
            {
                "name": "extrair_osme",
                "description": "Executa a extração de OSME no SAP para uma ou mais equipes/recursos (ex: CSC300)",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "recursos": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Lista com códigos de recursos/equipes para extrair OSME (ex: ['CSC300'])"
                        }
                    },
                    "required": ["recursos"]
                }
            },
            {
                "name": "consultar_migo",
                "description": "Consulta ordens ativas na IW39 e materiais/reservas na MIGO no SAP para Equipes Clandestinas (RCL001 a RCL008) ou outras equipes",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "equipe": {"type": "string", "description": "Código da equipe (ex: RCL001, RCL002)"},
                        "ordem": {"type": "string", "description": "Número da ordem IW39 opcional"}
                    },
                    "required": ["equipe"]
                }
            },
            {
                "name": "transferir_osme",
                "description": "Prepara e inicia o processo de Transferência de OSME (movimentações Z62 e Z61 no SAP) para equipes",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "equipes": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Lista com códigos de equipes para transferência de OSME"
                        }
                    },
                    "required": ["equipes"]
                }
            },
            {
                "name": "mudar_recurso",
                "description": "Executa consulta de saldo ou transferência integral de materiais entre duas equipes (Origem -> Destino) no SAP",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "origem": {"type": "string", "description": "Equipe de origem (ex: CSC395)"},
                        "destino": {"type": "string", "description": "Equipe de destino opcional (ex: CVC396)"},
                        "acao": {"type": "string", "enum": ["consultar", "executar"], "description": "'consultar' saldo ou 'executar' transferência"}
                    },
                    "required": ["origem"]
                }
            },
            {
                "name": "abrir_conferencia",
                "description": "Abre o módulo de Conferência de seriais/materiais com bipagem ou comparação",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "base": {"type": "string", "description": "Seriais recebidos (base)"},
                        "fisico": {"type": "string", "description": "Seriais coletados (físico)"}
                    }
                }
            },
            {
                "name": "abrir_fichas",
                "description": "Abre o módulo de Fichas WorkForce e opcionalmente preenche dados do funcionário por RE e recurso",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "re": {"type": "string", "description": "Número do RE do funcionário (ex: 19595)"},
                        "recurso": {"type": "string", "description": "Código do recurso/equipe (ex: CDC424)"}
                    }
                }
            },
            {
                "name": "abrir_modulo",
                "description": "Alterna e navega diretamente para qualquer aba/módulo do Conecta Hub / N.E.X.U.S",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "modulo": {
                            "type": "string",
                            "description": "Nome do módulo/aba: bandeirada, osme, migo, etiquetas, extrator, sap, conferencia, fichas, transfosme, mudanca, zm02, plataforma, home"
                        }
                    },
                    "required": ["modulo"]
                }
            },
            {
                "name": "consultar_iq09",
                "description": "Consulta um ou mais medidores/números de série ou intervalos diretamente na transação IQ09 do SAP (ex: '18398041', '18400025 a 18400028', '18400025, 18400026')",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "seriais": {
                            "type": "string",
                            "description": "Número do medidor, serial, intervalo ou lista com múltiplos seriais (ex: '18398041', '18400025 a 18400028', '18400025, 18400026')"
                        }
                    },
                    "required": ["seriais"]
                }
            },
            {
                "name": "buscar_lote",
                "description": "Busca informações completas de um lote (material, descrição, depósito, quantidade e lista de seriais)",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "lote": {"type": "string", "description": "Número do lote (ex: CL03-002)"}
                    },
                    "required": ["lote"]
                }
            },
            {
                "name": "buscar_multiplos_lotes",
                "description": "Busca informações de múltiplos lotes de uma vez",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "lotes": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Lista com códigos de lotes (ex: ['CL03-001', 'CL03-002'])"
                        }
                    },
                    "required": ["lotes"]
                }
            },
            {
                "name": "gerar_doc_google",
                "description": "Cria um documento no Google Docs com os seriais do lote e retorna o link de acesso",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "lote": {"type": "string", "description": "Número do lote (ex: CL03-002)"}
                    },
                    "required": ["lote"]
                }
            },
            {
                "name": "gerar_etiqueta",
                "description": "Gera uma etiqueta de identificação em PDF para impressão",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "deposito": {"type": "string", "description": "Depósito (ex: CL03, CL04, CS06)"},
                        "material": {"type": "string", "description": "Código do material (opcional)"},
                        "descricao": {"type": "string", "description": "Descrição do material (opcional)"},
                        "qtd": {"type": "string", "description": "Quantidade (opcional)"},
                        "unidade": {"type": "string", "description": "Unidade (ex: PCS, M, CJ)"},
                        "lote": {"type": "string", "description": "Número do lote (opcional)"}
                    },
                    "required": ["deposito"]
                }
            },
            {
                "name": "consultar_material",
                "description": "Busca a descrição oficial de um material pelo código no banco de dados local",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "codigo": {"type": "string", "description": "Código numérico do material (ex: 18398041)"}
                    },
                    "required": ["codigo"]
                }
            },
            {
                "name": "consultar_multiplos_materiais",
                "description": "Busca a descrição de vários materiais de uma vez",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "codigos": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Lista de códigos de material"
                        }
                    },
                    "required": ["codigos"]
                }
            },
            {
                "name": "extrair_seriais",
                "description": "Extrai números de série de um texto ou bloco com múltiplas entradas, suportando intervalos (ex: '18398041 a 18398050')",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "texto": {"type": "string", "description": "Texto contendo seriais, linhas coladas ou intervalos"}
                    },
                    "required": ["texto"]
                }
            },
            {
                "name": "listar_lotes",
                "description": "Lista todos os lotes disponíveis na planilha de controle de estoque",
                "parameters": {
                    "type": "object",
                    "properties": {}
                }
            },
            {
                "name": "login_sap",
                "description": "Faz login no SAP usando credenciais cadastradas",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "usuario": {"type": "string", "description": "Nome do usuário cadastrado (ex: Anderson, Junior, Pedro, Adriana)"}
                    },
                    "required": ["usuario"]
                }
            },
            {
                "name": "copiar_seriais",
                "description": "Copia os seriais extraídos recentemente para a área de transferência do Windows",
                "parameters": {
                    "type": "object",
                    "properties": {}
                }
            },
            {
                "name": "info_sap",
                "description": "Retorna informações detalhadas da sessão SAP conectada (sistema, cliente, usuário, transação)",
                "parameters": {
                    "type": "object",
                    "properties": {}
                }
            }
        ]
    }
]

class ChatBot:
    def __init__(self):
        self.config = self._carregar_config()
        self.historico = []
        self.contador = 0
        self.erros = 0
        self.acao_pendente = None
        self.gerenciador_skills = obter_gerenciador_skills()

    def _carregar_config(self):
        padrao = {"api_key": "", "modelo": "local/motor-integrado"}
        caminho = obter_caminho_config("config_chat.txt")
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    salva = json.loads(f.read().strip())
                    padrao.update(salva)
            except Exception:
                pass
        return padrao

    def salvar_config(self):
        caminho = os.path.join(obter_diretorio_dados(), "config_chat.txt")
        try:
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception:
            pass

    def obter_api_key(self):
        return self.config.get("api_key", "").strip()

    def obter_modelo(self):
        return self.config.get("modelo", "local/motor-integrado")

    def _converter_tools_para_openai(self):
        tools = []
        nomes_adicionados = set()
        for bloco in TOOLS:
            for func in bloco.get("functionDeclarations", []):
                tools.append({
                    "type": "function",
                    "function": {
                        "name": func["name"],
                        "description": func["description"],
                        "parameters": func["parameters"]
                    }
                })
                nomes_adicionados.add(func["name"])

        # Incorpora dinamicamente tools das skills .md
        if hasattr(self, "gerenciador_skills") and self.gerenciador_skills:
            for tool_decl in self.gerenciador_skills.gerar_declaracoes_tools_openrouter():
                if tool_decl["name"] not in nomes_adicionados:
                    tools.append({
                        "type": "function",
                        "function": tool_decl
                    })
                    nomes_adicionados.add(tool_decl["name"])

        return tools

    def processar_mensagem(self, mensagem, funcoes_disponiveis=None):
        """
        Processa a mensagem com fallback automático para o Motor Inteligente Local:
        1. Se o modelo configurado for 'local/motor-integrado' ou não houver chave API válida, processa localmente.
        2. Se houver chave API configurada, tenta a chamada na nuvem via OpenRouter.
        3. Se a nuvem falhar (401, 429, timeout ou sem internet), processa de forma transparente no Motor Local!
        """
        api_key = self.obter_api_key()
        modelo = self.obter_modelo()

        # Modo Local Direto
        if modelo == "local/motor-integrado" or not api_key:
            return self._processar_mensagem_local(mensagem, funcoes_disponiveis)

        try:
            resposta = self._chamar_api(api_key, modelo, mensagem)
            return self._processar_resposta(resposta, mensagem, funcoes_disponiveis)
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if hasattr(e, 'response') and e.response is not None else 0
            if status == 401:
                # Chave inválida ou expirada -> Fallback transparente para o Motor Local
                resp_local = self._processar_mensagem_local(mensagem, funcoes_disponiveis)
                return resp_local + "\n\n*(⚡ Processado pelo Motor Local Integrado • Chave OpenRouter inválida no ⚙️)*"
            elif status == 429:
                resp_local = self._processar_mensagem_local(mensagem, funcoes_disponiveis)
                return resp_local + "\n\n*(⚡ Processado pelo Motor Local • Limite de taxa OpenRouter atingido)*"
            # Outros erros HTTP -> tenta local
            return self._processar_mensagem_local(mensagem, funcoes_disponiveis)
        except Exception:
            # Erro de conexão/timeout -> Fallback local
            return self._processar_mensagem_local(mensagem, funcoes_disponiveis)

    def _chamar_api(self, api_key, model_id, mensagem):
        mensagens = [{"role": "system", "content": SYSTEM_PROMPT}]
        for msg in self.historico[-10:]:
            role = "assistant" if msg.get("role") == "model" else "user"
            text = ""
            for p in msg.get("parts", []):
                if "text" in p:
                    text = p["text"]
            if text:
                mensagens.append({"role": role, "content": text})
        mensagens.append({"role": "user", "content": mensagem})

        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
                "HTTP-Referer": "https://conectahub.local",
                "X-Title": "Conecta Hub 9.0"
            },
            json={
                "model": model_id,
                "messages": mensagens,
                "tools": self._converter_tools_para_openai(),
                "tool_choice": "auto"
            },
            timeout=(8, 35),
        )
        resp.raise_for_status()
        return resp.json()

    def _processar_resposta(self, resposta, mensagem, funcoes_disponiveis):
        choices = resposta.get("choices", [])
        if not choices:
            return self._processar_mensagem_local(mensagem, funcoes_disponiveis)
        msg = choices[0].get("message", {})
        self.contador += 1

        if msg.get("tool_calls"):
            resultados_executados = []
            for tc in msg["tool_calls"]:
                func_info = tc.get("function", {})
                nome_func = func_info.get("name", "")
                try:
                    args = json.loads(func_info.get("arguments", "{}"))
                except Exception:
                    args = {}

                resultado = self._executar_funcao(nome_func, args, funcoes_disponiveis)
                resultados_executados.append({
                    "funcao": nome_func,
                    "args": args,
                    "resultado": resultado
                })
                self._adicionar_historico("user", mensagem)
                self._adicionar_historico("model", f"[Executou {nome_func}: {json.dumps(args, ensure_ascii=False)}]")

            linhas = []
            for r in resultados_executados:
                res = r["resultado"]
                if isinstance(res, dict):
                    if "erro" in res:
                        linhas.append(f"❌ **{r['funcao']}**: {res['erro']}")
                    elif "mensagem" in res:
                        linhas.append(res["mensagem"])
                    elif "dados_medidor" in res:
                        m = res["dados_medidor"]
                        linhas.append(
                            f"⚡ **DETALHES DO MEDIDOR (SAP IQ09)**:\n\n"
                            f"• **Medidor / Serial**: `{m.get('serial', '—')}`\n"
                            f"• **Status**: `{m.get('status', 'Sem informação')}`\n"
                            f"• **Denominação**: {m.get('denominacao', '—')}\n"
                            f"• **Material**: `{m.get('material', '—')}`\n"
                            f"• **Equipamento SAP**: `{m.get('equipamento', '—')}`\n"
                            f"• **Centro**: `{m.get('centro', 'Sem informação')}`\n"
                            f"• **Depósito**: `{m.get('deposito', 'Sem informação')}`\n\n"
                            f"*(⚡ O Pop-up detalhado também foi exibido na tela)*"
                        )
                    elif "status" in res:
                        detalhe = f"\n📎 **Arquivo/Link**: `{res.get('arquivo', res.get('link', ''))}`" if res.get('arquivo') or res.get('link') else ""
                        linhas.append(f"✅ {res['status']}{detalhe}")
                    elif "descricao" in res:
                        linhas.append(f"📦 **Material {res.get('codigo', '')}**:\n`{res.get('descricao', '')}`")
                    elif "seriais" in res:
                        linhas.append(f"🔢 **Extração Concluída**: {res.get('total', len(res['seriais']))} seriais identificados.")
                    elif "total_seriais" in res:
                        linhas.append(f"📦 **Lote {res.get('lote', '')}**:\n• Material: `{res.get('material', '')}` - {res.get('descricao', '')}\n• Depósito: `{res.get('deposito', '')}`\n• Quantidade: `{res.get('total_seriais', 0)}` seriais")
                    else:
                        linhas.append(f"✓ {r['funcao']}: {json.dumps(res, ensure_ascii=False)[:200]}")
                else:
                    linhas.append(f"✓ {r['funcao']}: {str(res)[:200]}")
            return "\n\n".join(linhas)

        content = msg.get("content", "")
        if content:
            self._adicionar_historico("user", mensagem)
            self._adicionar_historico("model", content)
            return content
        return self._processar_mensagem_local(mensagem, funcoes_disponiveis)

    # ================== MOTOR INTELIGENTE LOCAL INTEGRADO (OFFLINE) ==================
    def _obter_usuarios_sap(self):
        """Retorna a lista de nomes de usuários cadastrados para o SAP."""
        try:
            from modulo_Sap import carregar_logins, logins_salvos
            carregar_logins()
            if logins_salvos:
                return list(logins_salvos.keys())
        except Exception:
            pass
        try:
            caminho = obter_caminho_config("logins_sap.json")
            if os.path.exists(caminho):
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    if isinstance(dados, dict):
                        return list(dados.keys())
        except Exception:
            pass
        return ["Adriana", "Junior", "Pedro"]

    def _processar_mensagem_local(self, mensagem, funcoes_disponiveis):
        """
        Processador NLP integrado baseado em intenções e extração de entidades.
        Funciona 100% offline sem necessidade de chave API ou internet!
        Compreende português natural e variações do dia a dia.
        """
        msg_norm = mensagem.strip()
        msg_lower = msg_norm.lower()
        tokens = msg_lower.split()

        if not funcoes_disponiveis:
            return "⚠️ Funções do sistema ainda não estão inicializadas. Aguarde um instante."

        # 0. Verificação de Confirmação ou Cancelamento de Ação Pendente
        confirmacoes_exatas = {
            "confirmo", "confirmar", "confirma", "confirmado", "sim", "s", "yes",
            "pode", "pode ir", "pode executar", "pode mandar", "executa", "executar",
            "prosseguir", "vai", "manda", "manda ver", "positivo", "ok", "bora",
            "prossiga", "efetuar", "efetua", "manda bala", "com certeza", "manda brasa",
            "processar", "processa"
        }
        cancelamentos_exatos = {
            "cancelar", "cancela", "cancelado", "não", "nao", "n", "no", "abortar", "aborta", "parar", "para"
        }

        # Não trata como confirmação se a frase for um comando completo com equipe, osme, etc.
        eh_comando_completo = bool(re.search(r"\b(osme|migo|bandeir|transfer|lote|material|serial|equipe|recurso|[A-Za-z]{2,4}\d{2,4})\b", msg_lower))

        if not eh_comando_completo:
            # Se for cancelamento
            if any(msg_lower == c or msg_lower.startswith(c + " ") for c in cancelamentos_exatos):
                if self.acao_pendente:
                    desc = self.acao_pendente.get("descricao", "Operação")
                    self.acao_pendente = None
                    return f"🛑 **{desc} cancelada.** Nenhuma alteração foi realizada no SAP."

            # Se for confirmação
            eh_confirmacao = (msg_lower in confirmacoes_exatas) or (len(tokens) <= 4 and any(t in ["confirmo", "confirmar", "confirma", "confirmado", "sim", "executar", "executa", "pode"] for t in tokens))
            if eh_confirmacao:
                if self.acao_pendente:
                    pend = self.acao_pendente
                    self.acao_pendente = None
                    func_nome = pend.get("funcao_exec")
                    func = funcoes_disponiveis.get(func_nome)
                    if func:
                        args = pend.get("args", {})
                        res = func(**args) if args else func()
                        if isinstance(res, dict) and "erro" in res:
                            return f"❌ **Falha ao executar {pend.get('descricao', 'operação')} no SAP**:\n{res['erro']}"
                        msg_sucesso = res.get("status", "Operação iniciada com sucesso no SAP!") if isinstance(res, dict) else str(res)
                        return (
                            f"🚀 **Confirmado! Executando {pend.get('descricao', 'operação')} no SAP...**\n\n"
                            f"✅ {msg_sucesso}\n\n"
                            f"*(Acompanhe o processamento no SAP GUI e o log na aba)*"
                        )
                    else:
                        return f"⚠️ Função de execução '{func_nome}' não está disponível no sistema."
                else:
                    # Fallback inteligente: se o usuário digitou 'confirmo' mas não havia pendente em memória,
                    # tenta disparar a Saída OSME ou Transferência caso haja dados preparados na aba!
                    func_saida = funcoes_disponiveis.get("executar_saida_osme")
                    if func_saida:
                        res_s = func_saida()
                        if isinstance(res_s, dict):
                            if res_s.get("sucesso"):
                                return (
                                    f"🚀 **Confirmado! Executando Saída OSME no SAP...**\n\n"
                                    f"✅ {res_s.get('status', 'Operação iniciada no SAP!')}\n\n"
                                    f"*(Acompanhe o processamento no SAP GUI e o log na aba)*"
                                )
                            elif "erro" in res_s and "Nenhum material" not in res_s["erro"]:
                                return f"❌ **Saída OSME**: {res_s['erro']}"


                    return "ℹ️ Nenhuma ação pendente de confirmação no momento. Como posso te ajudar?"

        # 1. Saudação
        saudacoes = ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite", "hey", "e ai", "e aí", "fala", "opa", "salve"]
        tokens = [t.strip(",.!?") for t in msg_lower.split()]
        if any(t in saudacoes for t in tokens) and len(tokens) <= 5 and not any(k in msg_lower for k in ["lote", "sap", "material", "serial", "doc", "copiar", "login", "bandeira", "osme", "migo", "pedro", "adriana", "junior"]):
            return (
                "👋 **Olá! Sou o assistente virtual do N.E.X.U.S.**\n\n"
                "Sou seu agente inteligente para automação de tarefas e operações. Exemplos de skills:\n\n"
                "• 🚩 **Bandeirada**: `bandeira o csc300` ou `bandeirar csc300, csc301`\n"
                "• 📤 **Saída OSME**: `processa na osme do cvc395 330702 50m`\n"
                "• 📄 **OSME**: `extrair osme csc300`\n"
                "• 🔍 **MIGO (Clandestinas)**: `migo rcl002` ou `consultar migo rcl001`\n"
                "• 🔄 **Mudança Recurso**: `mudar recurso do csc395 para cvc396`\n"
                "• 📦 **Transf. OSME**: `transferir osme csc300`\n"
                "• 🔑 **SAP**: `sap adriana`, `sap pedro`, `info sap`\n"
                "• 📦 **Lotes**: `CL03-002` ou `listar lotes`\n"
                "• 🔍 **Materiais**: `material 18398041` ou apenas `18398041`\n"
                "• 🔢 **Seriais**: `18398041 a 18398050` ou cole uma lista\n"
                "• 📋 **Navegação**: `abrir etiquetas`, `abrir conferencia`, `abrir fichas`\n\n"
                "*(⚡ Modo Integrado Ativo • Totalmente funcional sem precisar de API)*"
            )

        # 2. Ajuda / Comandos
        if msg_lower in ["help", "ajuda", "comandos", "menu", "?", "o que você faz", "o que voce faz", "opções", "opcoes"]:
            return (
                "🤖 **Catálogo de Skills do Hub Assistant:**\n\n"
                "• 🚩 **Bandeirada**: `bandeira o csc300` • `bandeirar csc300, csc301`\n"
                "• 📤 **Saída OSME**: `processa na osme do cvc395 330702 50m` • `saida osme csc300`\n"
                "• 📄 **Extração OSME**: `extrair osme csc300` • `osme csc300`\n"
                "• 🔍 **MIGO Clandestinas**: `migo rcl002` • `consultar migo rcl001`\n"
                "• 🔄 **Mudança de Recurso**: `mudar recurso do csc395 para cvc396` • `saldo origem csc395`\n"
                "• 📦 **Transf. OSME**: `transferir osme csc300`\n"
                "• 📦 **Conferência**: `conferir seriais` • `abrir conferencia`\n"
                "• 📋 **Fichas WorkForce**: `ficha re 19595` • `abrir fichas`\n"
                "• 🚀 **Navegação**: `abrir etiquetas` • `ir para sap` • `abrir zm02`\n"
                "• 🔑 **Login SAP**: `sap adriana` • `sap pedro` • `info sap`\n"
                "• 📦 **Consultar Lote**: `CL03-002` • `listar lotes`\n"
                "• 🔍 **Consultar Material**: `material 18398041` • `18398041`\n"
                "• 🔢 **Extrair Seriais**: `18398041 a 18398050` • cole blocos de texto\n"
                "• 📄 **Google Docs**: `doc CL03-002`\n"
                "• 🏷️ **Gerar Etiqueta**: `etiqueta CL03 material 18398041 qtd 50`\n"
                "• 📋 **Copiar**: `copiar seriais`"
            )

        # 3. Informações / Status do SAP
        termos_info_sap = ["info sap", "status sap", "sap status", "sap info", "sessao sap", "sessão sap", "sap aberto", "sap tá aberto", "sap ta aberto", "sap conectado", "ver sap", "como ta o sap", "como está o sap"]
        if any(term in msg_lower for term in termos_info_sap):
            func = funcoes_disponiveis.get("info_sap")
            if func:
                res = func()
                return f"🔑 **Status da Conexão SAP**:\n`{res.get('info', 'Não foi possível ler o status.')}`"

        # 4. SKILL: Bandeirada (ex: "bandeira o csc300", "bandeirar csc300, csc301", "fazer bandeirada equipe csc300")
        tem_bandeirada = bool(re.search(r"\b(bandeir[a-z]*|bandeirada)\b", msg_lower))
        if tem_bandeirada:
            achados_eq = [
                e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper())
                if e not in ["SAP", "DOC", "PDF", "HTTP"]
            ]
            if not achados_eq:
                tokens_resto = re.sub(
                    r"\b(bandeir[a-z]*|bandeirada|o|a|os|as|equipe|equipes|recurso|recursos|pra|pro|para|por|favor|fazer|executar|iniciar)\b",
                    " ",
                    msg_norm,
                    flags=re.IGNORECASE
                ).strip().split()
                achados_eq = [t.upper() for t in tokens_resto if len(t) >= 3 and any(c.isdigit() for c in t)]

            if achados_eq:
                func = funcoes_disponiveis.get("executar_bandeirada")
                if func:
                    res = func(recursos=achados_eq)
                    if "erro" in res:
                        return f"❌ **Bandeirada**: {res['erro']}"
                    return f"🚩 **Bandeirada SAP**:\n✅ {res.get('status', 'Processo iniciado no SAP!')}\n\n*(A aba de Bandeirada foi carregada na tela)*"
            else:
                return (
                    "🚩 **Bandeirada SAP**: Qual equipe ou recurso você deseja bandeirar?\n\n"
                    "• Exemplo: `bandeira o CSC300` ou `bandeirar CSC300, CSC301`"
                )

        # 4.1 SKILL: Saída para OSME / Baixa de Materiais e Medidores
        tem_saida_osme = bool(re.search(
            r"\b(sa[ií]da(?:\s+para)?\s+osme|process[ae|ar](?:\s+na)?\s+osme|baix[ae|ar](?:\s+na)?\s+osme|lan[cç][ae|ar](?:\s+na)?\s+osme)\b",
            msg_lower
        )) or (
            bool(re.search(r"\b(process[ae|ar]|baix[ae|ar]|lan[cç][ae|ar]|dar\s+baixa|sa[ií]da)\b", msg_lower)) and
            bool(re.search(r"\bosme\b", msg_lower))
        )
        if tem_saida_osme:
            # 1. Identifica a equipe/recurso (ex: CVC395, CSC300)
            achados_eq = [
                e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper())
                if e not in ["OSME", "SAP", "DOC", "PDF", "HTTP", "MIGO", "IW39", "Z61", "Z62", "UND", "PCS"]
            ]
            eq_alvo = achados_eq[0] if achados_eq else None

            # 2. Identifica materiais não serializados e quantidades (ex: 330702 50m ou 510388 20 un)
            mat_ns = None
            qtd_ns = None

            m_ns1 = re.search(r"\b(\d{5,7})\b\s+(\d+(?:[.,]\d+)?\s*(?:m|un|pcs|cx|rolo|metros|unidades)?)\b", msg_norm, re.IGNORECASE)
            m_ns2 = re.search(r"\b(\d+(?:[.,]\d+)?\s*(?:m|un|pcs|cx|rolo|metros|unidades))\b\s+(?:de|do|material)?\s*(\b\d{5,7}\b)", msg_norm, re.IGNORECASE)
            if m_ns1:
                mat_ns = m_ns1.group(1)
                qtd_ns = m_ns1.group(2)
            elif m_ns2:
                qtd_ns = m_ns2.group(1)
                mat_ns = m_ns2.group(2)
            else:
                m_rot_mat = re.search(r"\b(?:material|codigo|código)\s*[:=]?\s*(\d{5,7})\b", msg_norm, re.IGNORECASE)
                m_rot_qtd = re.search(r"\b(?:qtd|quantidade)\s*[:=]?\s*(\d+(?:[.,]\d+)?\s*(?:m|un|pcs|cx)?)\b", msg_norm, re.IGNORECASE)
                if m_rot_mat:
                    mat_ns = m_rot_mat.group(1)
                if m_rot_qtd:
                    qtd_ns = m_rot_qtd.group(1)

            # 3. Identifica seriais ou intervalos de seriais (ex: 18398041 a 18398045 ou bloco de seriais)
            from modulo_extrator import extrair_lista_seriais
            texto_para_seriais = msg_norm
            if mat_ns:
                texto_para_seriais = texto_para_seriais.replace(mat_ns, " ")
            if eq_alvo:
                texto_para_seriais = texto_para_seriais.replace(eq_alvo, " ")

            seriais_extraidos = extrair_lista_seriais(texto_para_seriais)
            # Filtra seriais válidos com comprimento consistente
            seriais_extraidos = [s for s in seriais_extraidos if len(s) >= 6]

            func = funcoes_disponiveis.get("saida_osme")
            if func:
                if not eq_alvo:
                    return (
                        "📤 **Saída para OSME**: Para qual equipe você deseja preparar a saída?\n\n"
                        "• Exemplo: `Processa na osme do cvc395 330702 50m` ou `saida osme csc300 seriais 18398041 a 18398045`"
                    )

                res = func(
                    equipe=eq_alvo,
                    seriais=seriais_extraidos if seriais_extraidos else None,
                    material=mat_ns,
                    quantidade=qtd_ns
                )
                if "erro" in res:
                    return f"❌ **Saída para OSME**: {res['erro']}"
                self.acao_pendente = {
                    "tipo": "saida_osme",
                    "descricao": "Saída OSME",
                    "funcao_exec": "executar_saida_osme"
                }
                return res.get("mensagem", "📋 Dados preparados na aba Saída OSME.")

        # 5. SKILL: Extração OSME (ex: "extrair osme csc300", "osme do csc300", "puxar osme csc300")
        tem_osme = bool(re.search(r"\b(osme|extrair osme|puxar osme|extracao osme|extração osme)\b", msg_lower))
        if tem_osme and not any(k in msg_lower for k in [
            "transferir", "transferencia", "transferência", "transf", "z62", "z61",
            "processa", "processe", "processar", "saida", "saída", "baixa", "baixar", "dar baixa", "lancar", "lançar"
        ]):
            achados_eq = [
                e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper())
                if e not in ["OSME", "SAP", "DOC", "PDF"]
            ]
            if not achados_eq:
                tokens_resto = re.sub(
                    r"\b(extrair|puxar|extracao|extração|osme|d[oa]|de|equipe|equipes|recurso|recursos|pra|pro|para|por|favor|fazer|executar|iniciar)\b",
                    " ",
                    msg_norm,
                    flags=re.IGNORECASE
                ).strip().split()
                achados_eq = [t.upper() for t in tokens_resto if len(t) >= 3 and any(c.isdigit() for c in t)]

            if achados_eq:
                func = funcoes_disponiveis.get("extrair_osme")
                if func:
                    res = func(recursos=achados_eq)
                    if "erro" in res:
                        return f"❌ **Extração OSME**: {res['erro']}"
                    return f"📄 **Extração OSME SAP**:\n✅ {res.get('status', 'Extração iniciada no SAP!')}\n\n*(A aba de OSME foi carregada na tela)*"
            else:
                return (
                    "📄 **Extração OSME**: Qual equipe ou recurso você deseja extrair?\n\n"
                    "• Exemplo: `extrair osme CSC300` ou `osme CSC300, CSC301`"
                )

        # 6. SKILL: Consultar MIGO / IW39 Clandestinas (ex: "consultar migo rcl001", "migo rcl002", "ordem rcl001")
        tem_migo = bool(re.search(r"\b(migo|iw39|clandestina|clandestinas)\b", msg_lower))
        match_rcl_direto = re.search(r"\b(RCL00[1-8]|RCL\d{3})\b", msg_norm.upper())
        if tem_migo or match_rcl_direto:
            eq_migo = match_rcl_direto.group(1) if match_rcl_direto else None
            if not eq_migo:
                match_qualquer_eq = re.search(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper())
                if match_qualquer_eq and match_qualquer_eq.group(1) not in ["MIGO", "IW39", "SAP"]:
                    eq_migo = match_qualquer_eq.group(1)

            match_ordem = re.search(r"\b(1\d{7,11}|\d{8,12})\b", msg_norm)
            ordem_val = match_ordem.group(1) if match_ordem else ""

            if eq_migo:
                func = funcoes_disponiveis.get("consultar_migo")
                if func:
                    res = func(equipe=eq_migo, ordem=ordem_val)
                    if "erro" in res:
                        return f"❌ **Consultar MIGO**: {res['erro']}"
                    detalhe_ordem = f"\n• Ordem: `{ordem_val}`" if ordem_val else ""
                    return (
                        f"🔍 **Consulta MIGO (IW39 ➡️ MIGO)**:\n"
                        f"• Equipe: `{eq_migo}`{detalhe_ordem}\n"
                        f"✅ {res.get('status', 'Processo disparado no SAP!')}\n\n"
                        f"*(📦 O Pop-up com o Saldo Disponível será exibido automaticamente na tela)*"
                    )
            elif tem_migo:
                return (
                    "🔍 **Consultar MIGO (Clandestinas)**: Para qual equipe você deseja consultar?\n\n"
                    "• Equipes disponíveis: `RCL001`, `RCL002`, `RCL003`, `RCL004`, `RCL005`, `RCL006`, `RCL007`, `RCL008`\n"
                    "• Exemplo: `migo RCL002` ou `consultar migo RCL001`"
                )

        # 6.2 SKILL: Contagem de Carro (ZMCSE01 ➡️ IW39 ➡️ Z62 ➡️ Z61)
        tem_contagem_carro = bool(re.search(
            r"\b(contagem(?:\s+de)?\s+carro|contar(?:\s+o)?\s+carro|balan[cç]o(?:\s+de|\s+do)?\s+carro|auditoria(?:\s+de|\s+do)?\s+carro|contagem(?:\s+de)?\s+ve[ií]culo)\b",
            msg_lower
        ))
        if tem_contagem_carro:
            achados_eq = [
                e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper())
                if e not in ["SAP", "DOC", "PDF", "HTTP", "MIGO", "IW39", "Z61", "Z62", "CARRO"]
            ]
            eq_alvo = achados_eq[0] if achados_eq else None

            from modulo_extrator import extrair_lista_seriais
            seriais_extraidos = extrair_lista_seriais(msg_norm)
            seriais_extraidos = [s for s in seriais_extraidos if len(s) >= 6]

            func = funcoes_disponiveis.get("contagem_carro")
            if func:
                res = func(
                    equipe=eq_alvo,
                    seriais="\n".join(seriais_extraidos) if seriais_extraidos else None,
                    acao="osme" if not seriais_extraidos else "abrir"
                )
                if "erro" in res:
                    return f"❌ **Contagem de Carro**: {res['erro']}"
                eq_txt = f" para a equipe `{eq_alvo}`" if eq_alvo else ""
                return (
                    f"🚗 **Contagem de Carro (Balanço de Medidores)**:\n"
                    f"✅ {res.get('status', 'Aba carregada com sucesso.')}\n\n"
                    f"• Equipe: `{eq_alvo or 'Aguardando seleção'}`\n"
                    f"• Fluxo: ZMCSE01 (Auditoria) ➡️ IW39 (Ordem Recente) ➡️ MIGO Z62 (Estorno) ➡️ MIGO Z61 (Saída)\n\n"
                    f"👉 *Confira os medidores na aba e clique em **'🚀 Executar Ciclo Completo no SAP'** para processar.*"
                )

        # 7. SKILL: Transferência de OSME (ex: "transferir osme csc300", "transferencia osme csc300")
        tem_transf_osme = bool(re.search(r"\b(transferir|transferencia|transferência|transf|fila)\s*(?:de\s*)?osme\b", msg_lower))
        if tem_transf_osme:
            achados_eq = [e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper()) if e not in ["OSME", "SAP"]]
            if achados_eq:
                func = funcoes_disponiveis.get("transferir_osme")
                if func:
                    res = func(equipes=achados_eq)
                    if "erro" in res:
                        return f"❌ **Transferência OSME**: {res['erro']}"
                    return f"📦 **Transferência de OSME**:\n✅ {res.get('status', 'Processo disparado!')}"
            else:
                return "📦 **Transferência de OSME**: Informe as equipes para transferência (ex: `transferir osme CSC300`)."

        # 8. SKILL: Mudança de Recurso (ex: "mudar recurso do csc395 para cvc396", "saldo origem csc395")
        tem_mudanca = bool(re.search(r"\b(mudar|muda|mudanca|mudança|trocar|transferir)\s*recurso\b|\bsaldo\s*(?:de\s*|da\s*)?origem\b", msg_lower))
        match_transf_dupla = re.search(r"(?:mudar|transferir|passar)\s+(?:d[oa]\s+)?([A-Za-z0-9]{4,8})\s+(?:para|pro|pra|a)\s+([A-Za-z0-9]{4,8})", msg_norm, re.IGNORECASE)
        if tem_mudanca or match_transf_dupla:
            origem = ""
            destino = ""
            if match_transf_dupla:
                origem = match_transf_dupla.group(1).upper()
                destino = match_transf_dupla.group(2).upper()
            else:
                achados_eq = [e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper()) if e not in ["SAP", "DOC", "PDF"]]
                if len(achados_eq) >= 2:
                    origem = achados_eq[0]
                    destino = achados_eq[1]
                elif len(achados_eq) == 1:
                    origem = achados_eq[0]

            acao = "consultar" if ("saldo" in msg_lower or not destino) else "executar"
            if origem:
                func = funcoes_disponiveis.get("mudar_recurso")
                if func:
                    res = func(origem=origem, destino=destino, acao=acao)
                    if "erro" in res:
                        return f"❌ **Mudança de Recurso**: {res['erro']}"
                    return f"🔄 **Mudança de Recurso**:\n✅ {res.get('status', 'Preparado com sucesso!')}"
            else:
                return "🔄 **Mudança de Recurso**: Informe a equipe de origem e destino (ex: `mudar recurso do CSC395 para CVC396` ou `saldo origem CSC395`)."

        # 9. SKILL: Conferência (ex: "conferir seriais", "abrir conferencia", "comparar seriais")
        if any(term in msg_lower for term in ["conferir", "conferencia", "conferência", "comparar seriais", "bipagem", "bipar seriais"]):
            func = funcoes_disponiveis.get("abrir_conferencia")
            if func:
                res = func()
                return f"📦 **Módulo de Conferência**:\n✅ {res.get('status', 'Aba aberta para conferência física.')}"

        # 10. SKILL: Fichas WorkForce (ex: "abrir fichas", "ficha re 19595", "preencher ficha 19595")
        if any(term in msg_lower for term in ["ficha", "fichas", "workforce"]) and "lote" not in msg_lower:
            match_re = re.search(r"\b(?:re\s*[:=]?\s*)?(\d{5,7})\b", msg_norm)
            re_val = match_re.group(1) if match_re else ""
            match_rec = re.search(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper())
            rec_val = match_rec.group(1) if match_rec and match_rec.group(1) not in ["FICHA", "FICHAS"] else ""
            func = funcoes_disponiveis.get("abrir_fichas")
            if func:
                res = func(re=re_val, recurso=rec_val)
                return f"📋 **Módulo Fichas WorkForce**:\n✅ {res.get('status', 'Aba aberta com sucesso.')}"

        # 11. SKILL: Navegação entre abas (ex: "abrir etiquetas", "ir para bandeirada", "mostrar dashboard", "abrir zm02")
        gatilhos_nav = ["abrir aba", "ir para aba", "mostrar aba", "ver aba", "trocar para", "abrir modulo", "abrir módulo", "ir para", "abre a aba", "abre o módulo"]
        if any(msg_lower.startswith(g) for g in gatilhos_nav) or (msg_lower.startswith("abrir ") and len(tokens) <= 3 and not any(k in msg_lower for k in ["sap", "login", "doc", "sessao", "sessão"])):
            alvo = re.sub(r"^(abrir aba|ir para aba|mostrar aba|ver aba|trocar para|abrir modulo|abrir módulo|ir para|abre a aba|abre o módulo|abrir|abre)\s+", "", msg_lower).strip()
            func = funcoes_disponiveis.get("abrir_modulo")
            if func and alvo:
                res = func(modulo=alvo)
                return f"🚀 **Navegação**:\n✅ {res.get('status', f'Alternado para {alvo}')}"

        # 12. SKILL: Consultar Medidor na IQ09
        # Dicionário de Sinônimos e Mapeamento Semântico de Intenção (NLP Heurístico):
        # Suporta linguagem humana natural (ex: "olha esse medidor 123456789", "dá uma olhada no serial 18400028", "checa esse aparelho", etc.)
        # e detecção direta de medidores alfanuméricos (ex: UQM9HDN26L01533345)
        gatilhos_verbos = r"(?:olh[ae]|olhar|d[áa]\s+uma\s+olhada|vej[ao]|v[êe]|ver|chec[ae]|checar|pesquis[ae]|pesquisar|consult[ae]|consultar|busc[ae]|buscar|pux[ae]|puxar|confer[ie]|conferir|mostr[ae]|mostrar|analis[ae]|analisar|traz|traga|trazer|situa[çc][ãa]o|status)"
        gatilhos_entidades = r"(?:medidor(?:es)?|aparelho(?:s)?|rel[óo]gio(?:s)?|relogio(?:s)?|serial|seriais|equipamento(?:s)?|n[úu]mero\s+de\s+s[ée]rie|iq09)"

        padrao_alfa_serial = re.compile(r"\b(?=[A-Za-z0-9_\-]*[A-Za-z])(?=[A-Za-z0-9_\-]*\d)[A-Za-z0-9_\-]{8,22}\b")
        eh_lote_declarado = bool(re.search(r"\b[A-Z]{2,4}\d{0,3}[-_ ]\d{2,5}\b", msg_norm.upper()))
        achados_alfa_medidor = [
            t for t in padrao_alfa_serial.findall(msg_norm)
            if sum(c.isdigit() for c in t) >= 2 and t.upper() not in ["CONECTA", "WORKFORCE", "ASSISTANT", "BANDEIRADA", "TRANSFERENCIA", "CONFERENCIA"]
        ]

        tem_iq09 = False
        if not eh_lote_declarado:
            if re.search(r"\b(iq09|medidor(?:es)?|aparelho(?:s)?|rel[óo]gio(?:s)?|relogio(?:s)?)\b", msg_lower):
                tem_iq09 = True
            elif re.search(rf"\b{gatilhos_verbos}\b.*?\b{gatilhos_entidades}\b", msg_lower):
                tem_iq09 = True
            elif achados_alfa_medidor:
                tem_iq09 = True

        if tem_iq09:
            from modulo_extrator import extrair_lista_seriais
            lista_ser = extrair_lista_seriais(msg_norm)
            func = funcoes_disponiveis.get("consultar_iq09")
            if func:
                res = func(seriais=lista_ser if lista_ser else None)
                if "erro" in res:
                    return f"❌ **IQ09**: {res['erro']}"

                lista_m = res.get("lista_medidores", [])
                if len(lista_m) > 1:
                    qtd_disp = sum(1 for item in lista_m if "DISP" in str(item.get("status", "")).upper())
                    linhas = [
                        f"⚡ **CONSULTA IQ09 NO SAP ({len(lista_m)} MEDIDORES)**:\n",
                        f"📊 **Resumo**: `{qtd_disp}` Disponíveis (DISP) | `{len(lista_m) - qtd_disp}` Outros Status\n"
                    ]
                    for m in lista_m[:12]:
                        st_m = m.get("status", "Sem informação")
                        mat_m = m.get("material", "—")
                        dep_m = m.get("deposito", "—")
                        desc_m = m.get("denominacao", "")
                        linhas.append(f"• **{m.get('serial', '—')}**: `{st_m}` | Mat `{mat_m}` | Dep `{dep_m}` | {desc_m}")
                    if len(lista_m) > 12:
                        linhas.append(f"\n*(e mais {len(lista_m) - 12} medidores...)*")
                    linhas.append("\n*(⚡ A janela com a tabela de todos os medidores foi exibida na tela)*")
                    return "\n".join(linhas)

                elif "dados_medidor" in res and res["dados_medidor"]:
                    m = res["dados_medidor"]
                    serial_exib = m.get("serial") or (lista_ser[0] if lista_ser else "—")
                    return (
                        f"⚡ **DETALHES DO MEDIDOR (SAP IQ09)**:\n\n"
                        f"• **Medidor / Serial**: `{serial_exib}`\n"
                        f"• **Status**: `{m.get('status', 'Sem informação')}`\n"
                        f"• **Denominação**: {m.get('denominacao', '—')}\n"
                        f"• **Material**: `{m.get('material', '—')}`\n"
                        f"• **Equipamento SAP**: `{m.get('equipamento', '—')}`\n"
                        f"• **Centro**: `{m.get('centro', 'Sem informação')}`\n"
                        f"• **Depósito**: `{m.get('deposito', 'Sem informação')}`\n\n"
                        f"*(⚡ O Pop-up detalhado foi exibido na tela)*"
                    )
                return f"⚡ **Consulta IQ09**: {res.get('status', 'Pesquisa concluída.')}"
            else:
                return "⚡ **Consulta IQ09**: Qual medidor, serial ou intervalo você deseja pesquisar no SAP? (Ex: `iq09 18400025 a 18400028`, `UQM9HDN26L01533345` ou `pesquise o medidor 18398041`)"

        # 13. Login SAP (Compreensão natural avançada)
        usuarios_cadastrados = self._obter_usuarios_sap()
        gatilhos_login = ["sap", "login", "logar", "logue", "loga", "entra", "entrar", "conectar", "conecta", "acessar", "acesse"]
        tem_gatilho = any(re.search(rf"\b{g}\b", msg_lower) for g in gatilhos_login)

        # Se o usuário digitou APENAS o nome de um usuário cadastrado (ex: "pedro", "adriana")
        usuario_exato = next((u for u in usuarios_cadastrados if u.lower() == msg_lower.strip()), None)

        if tem_gatilho or usuario_exato:
            usuario_alvo = None
            if usuario_exato:
                usuario_alvo = usuario_exato
            else:
                # Procura por qualquer nome de usuário cadastrado dentro da frase
                for u in usuarios_cadastrados:
                    if re.search(rf"\b{re.escape(u)}\b", msg_norm, re.IGNORECASE):
                        usuario_alvo = u
                        break

            # Se ainda não encontrou pelos cadastrados, tenta extrair o nome limpando verbos e preposições
            if not usuario_alvo and ("login" in msg_lower or "logar" in msg_lower or "logue" in msg_lower or "loga" in msg_lower or "sap" in msg_lower or "entra" in msg_lower):
                texto_limpo = re.sub(
                    r"\b(login|logar|logue|loga|entrar|entra|conectar|conecta|abrir|abre|acessar|acesse|sap|no|na|nos|nas|do|da|dos|das|de|com|como|o|a|os|as|em|para|pra|pro|por|favor)\b",
                    " ",
                    msg_norm,
                    flags=re.IGNORECASE
                ).strip()
                palavras_restantes = [p for p in texto_limpo.split() if len(p) >= 2 and p.isalpha()]
                if palavras_restantes:
                    usuario_alvo = palavras_restantes[0].capitalize()

            # Se identificou um usuário, dispara o login
            if usuario_alvo:
                func = funcoes_disponiveis.get("login_sap")
                if func:
                    res = func(usuario=usuario_alvo)
                    if "erro" in res:
                        return f"❌ {res['erro']}"
                    return f"🔑 **Login SAP**: {res.get('status', f'Iniciando login para {usuario_alvo}...')}"

            # Se pediu login/sap mas não especificou o usuário, lista os disponíveis
            if "sap" in msg_lower or "login" in msg_lower or "logar" in msg_lower or "logue" in msg_lower or "loga" in msg_lower:
                lista_usuarios = "\n".join(f"• `{u}`" for u in usuarios_cadastrados) if usuarios_cadastrados else "• `Adriana`\n• `Junior`\n• `Pedro`"
                return (
                    "🔑 **Login SAP**: Para qual usuário você deseja fazer login?\n\n"
                    f"**Usuários cadastrados:**\n{lista_usuarios}\n\n"
                    "*(Dica: basta responder com `sap Pedro` ou apenas `Pedro`)*"
                )

        # 5. Copiar Seriais
        termos_copiar = ["copiar", "copiar seriais", "copiar serial", "copia", "copia os seriais", "copiar lista", "copiar tudo", "transferir seriais"]
        if any(term in msg_lower for term in termos_copiar) or (msg_lower.startswith("copia") and len(tokens) <= 3):
            func = funcoes_disponiveis.get("copiar_seriais")
            if func:
                res = func()
                return f"📋 {res.get('status', 'Seriais copiados com sucesso!')}"

        # 6. Listar Lotes
        termos_listar = ["listar lotes", "todos os lotes", "ver lotes", "quais lotes", "lotes disponiveis", "lotes disponíveis", "lotes estoque", "lista de lotes", "mostrar lotes", "ver estoque"]
        if any(term in msg_lower for term in termos_listar) or msg_lower == "lotes":
            func = funcoes_disponiveis.get("listar_lotes")
            if func:
                res = func()
                if "erro" in res:
                    return f"❌ {res['erro']}"
                lotes = res.get("lotes", [])
                total = res.get("total", len(lotes))
                lista_txt = ", ".join(lotes[:30])
                mais = f"\n*(e mais {total - 30} lotes...)*" if total > 30 else ""
                return f"📦 **Total de {total} Lotes Disponíveis no Estoque**:\n\n`{lista_txt}`{mais}"

        # 7. Gerar Google Doc a partir de Lote
        if any(k in msg_lower for k in ["doc", "documento", "drive", "google"]):
            achados_lote_doc = re.findall(r"\b([A-Z]{2,4}\d{0,3}[-_ ]?\d{2,5})\b", msg_norm.upper())
            if achados_lote_doc:
                lote_doc = re.sub(r"\s+", "-", achados_lote_doc[0]).replace("_", "-")
                func = funcoes_disponiveis.get("gerar_doc_google")
                if func:
                    res = func(lote=lote_doc)
                    if "erro" in res:
                        return f"❌ Erro ao gerar Google Doc para lote `{lote_doc}`:\n{res['erro']}"
                    return f"📄 **Google Docs Criado com Sucesso!**\n• Lote: `{lote_doc}`\n• Quantidade: `{res.get('quantidade', '')}` seriais\n• Link de Acesso: {res.get('link', '')}"

        # 8. Consulta de Lote(s)
        achados_lotes_raw = re.findall(r"\b([A-Z]{2,4}\d{0,3}[-_ ]?\d{2,5})\b", msg_norm.upper())
        # Filtra siglas comuns que não são lotes (ex: "SAP", "PCS", "DOC")
        achados_lotes = [
            re.sub(r"\s+", "-", l).replace("_", "-")
            for l in achados_lotes_raw
            if not any(l.startswith(ign) for ign in ["SAP", "DOC", "PDF", "HTTP"])
        ]
        if achados_lotes and ("extrair" not in msg_lower and "seriais" not in msg_lower):
            if len(achados_lotes) == 1:
                func = funcoes_disponiveis.get("buscar_lote")
                if func:
                    res = func(lote=achados_lotes[0])
                    if "erro" in res:
                        return f"❌ Lote `{achados_lotes[0]}`: {res['erro']}"
                    return (
                        f"📦 **Informações do Lote `{res.get('lote', achados_lotes[0])}`**:\n"
                        f"• Material: `{res.get('material', '-')}` ({res.get('descricao', 'Sem descrição')})\n"
                        f"• Depósito: `{res.get('deposito', '-')}`\n"
                        f"• Quantidade: **{res.get('total_seriais', res.get('quantidade', 0))}** seriais\n"
                        f"*(Campos preenchidos automaticamente na aba Gerador de Etiquetas!)*"
                    )
            else:
                func_mult = funcoes_disponiveis.get("buscar_multiplos_lotes")
                if func_mult:
                    res = func_mult(lotes=achados_lotes)
                    linhas = [f"📦 **Resultado da Consulta de {len(achados_lotes)} Lotes**:"]
                    for l, dados in res.get("resultados", {}).items():
                        if "erro" in dados:
                            linhas.append(f"• **{l}**: ❌ {dados['erro']}")
                        else:
                            linhas.append(f"• **{l}**: Mat `{dados.get('material', '')}` | Dep `{dados.get('deposito', '')}` | **{dados.get('total_seriais', 0)}** seriais")
                    return "\n".join(linhas)

        # 9. Extração de Seriais (Intervalo ou Lista de Seriais)
        tem_intervalo = re.search(r"\b\d{5,18}\s*(?:a|à|-|ate|até)\s*\d{5,18}\b", msg_norm, re.IGNORECASE)
        achados_digitos = re.findall(r"\b\d{6,18}\b", msg_norm)
        if tem_intervalo or len(achados_digitos) >= 2 or ("extrair" in msg_lower and achados_digitos):
            func = funcoes_disponiveis.get("extrair_seriais")
            if func:
                res = func(texto=msg_norm)
                if "erro" in res:
                    return f"❌ {res['erro']}"
                total = res.get("total_extraidos", res.get("total", 0))
                seriais = res.get("seriais", [])
                amostra = ", ".join(seriais[:8]) + ("..." if total > 8 else "")
                return (
                    f"🔢 **Extração de Seriais Concluída!**\n"
                    f"• Total extraído: **{total}** seriais\n"
                    f"• Amostra: `{amostra}`\n\n"
                    f"*(Os seriais já foram carregados na aba Extrator. Digite `copiar seriais` para transferir para o Windows)*"
                )

        # 10. Consulta de Material(is)
        match_mat = re.findall(r"\b(18\d{6}|[0-9]{7,8})\b", msg_norm)
        if match_mat and ("material" in msg_lower or "código" in msg_lower or "codigo" in msg_lower or "descrição" in msg_lower or len(match_mat) == 1):
            if len(match_mat) == 1:
                func = funcoes_disponiveis.get("consultar_material")
                if func:
                    res = func(codigo=match_mat[0])
                    if "erro" in res:
                        return f"❌ {res['erro']}"
                    return f"🔍 **Material `{res.get('codigo', match_mat[0])}`**:\n`{res.get('descricao', 'Descrição não encontrada')}`"
            else:
                func_mult = funcoes_disponiveis.get("consultar_multiplos_materiais")
                if func_mult:
                    res = func_mult(codigos=match_mat)
                    linhas = [f"🔍 **Consulta de {len(match_mat)} Materiais**:"]
                    for c, dados in res.get("materiais", {}).items():
                        linhas.append(f"• **{c}**: {dados.get('descricao', '-')}")
                    return "\n".join(linhas)

        # 18. Avaliação Dinâmica de Skills em Markdown (.md)
        if hasattr(self, "gerenciador_skills") and self.gerenciador_skills:
            skill_match = self.gerenciador_skills.buscar_por_gatilho(msg_norm)
            if skill_match:
                args = {}
                # Se a skill espera 'serial' ou 'seriais'
                if "serial" in skill_match.parameters or "seriais" in skill_match.parameters:
                    param_key = "serial" if "serial" in skill_match.parameters else "seriais"
                    match_seriais = re.findall(r"\b([A-Za-z0-9]{6,25})\b", msg_norm)
                    palavras_ignorar = {
                        "consultar", "pesquisar", "planilha", "verificar", "medidor",
                        "medidores", "serial", "seriais", "equipe", "recurso", "status",
                        "rete", "olha", "olhar", "veja", "checar", "busca", "buscar"
                    }
                    candidatos = [s for s in match_seriais if s.lower() not in palavras_ignorar]
                    if candidatos:
                        args[param_key] = candidatos[0] if param_key == "serial" else " ".join(candidatos)

                # Se a skill espera 'recursos' ou 'equipes'
                if "recursos" in skill_match.parameters or "equipes" in skill_match.parameters:
                    param_key = "recursos" if "recursos" in skill_match.parameters else "equipes"
                    match_eq = [e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper()) if e not in ["SAP", "DOC", "PDF", "HTTP", "RETE"]]
                    if match_eq:
                        args[param_key] = match_eq

                # Se a skill espera 'equipe'
                if "equipe" in skill_match.parameters:
                    match_eq = [e for e in re.findall(r"\b([A-Za-z]{2,4}\d{2,4})\b", msg_norm.upper()) if e not in ["SAP", "DOC", "PDF", "HTTP", "RETE"]]
                    if match_eq:
                        args["equipe"] = match_eq[0]

                res = self.gerenciador_skills.executar_skill(skill_match.name, args, funcoes_disponiveis)
                if isinstance(res, dict):
                    if "resposta" in res:
                        return res["resposta"]
                    if "erro" in res:
                        return f"❌ **{skill_match.name}**: {res['erro']}"
                    if "status" in res:
                        return f"✅ **{skill_match.name}**: {res['status']}"
                    return f"📄 **{skill_match.name}**:\n`{json.dumps(res, ensure_ascii=False, indent=2)}`"
                elif res is not None:
                    return str(res)

        # 19. Resposta Padrão Amigável com Reconhecimento
        return (
            f"🤖 Entendido! Você digitou: *\"{msg_norm}\"*\n\n"
            "Não identifiquei uma ação direta ou o formato não foi reconhecido. Como posso te ajudar?\n"
            "• 🚩 **Bandeirada**: `bandeira o CSC300` ou `bandeirar CSC300, CSC301`\n"
            "• 📄 **OSME**: `extrair osme CSC300`\n"
            "• 🔍 **MIGO (Clandestinas)**: `migo RCL002`\n"
            "• 🔄 **Mudança Recurso**: `mudar recurso do CSC395 para CVC396`\n"
            "• 📦 **Lotes**: `CL03-002` ou `listar lotes`\n"
            "• 🔢 **Seriais**: `18398041 a 18398050` ou cole texto\n"
            "• 🔑 **SAP**: `sap Adriana` ou `info sap`\n"
            "• Digite `ajuda` para ver todos os comandos!"
        )

    def _executar_funcao(self, nome, args, funcoes):
        if funcoes and nome in funcoes:
            if nome == "saida_osme":
                self.acao_pendente = {
                    "tipo": "saida_osme",
                    "descricao": "Saída OSME",
                    "funcao_exec": "executar_saida_osme"
                }
            elif nome == "executar_saida_osme":
                self.acao_pendente = None

            func = funcoes.get(nome)
            try:
                return func(**args)
            except TypeError as te:
                try:
                    return func()
                except Exception:
                    return {"erro": f"Parâmetros inválidos para {nome}: {str(te)}"}
            except Exception as e:
                return {"erro": str(e)}

        # Tenta executar via Gerenciador de Skills em Markdown
        if hasattr(self, "gerenciador_skills") and self.gerenciador_skills:
            return self.gerenciador_skills.executar_skill(nome, args, funcoes)

        return {"erro": f"Função '{nome}' não cadastrada"}

    def _adicionar_historico(self, role, texto):
        self.historico.append({"role": role, "parts": [{"text": texto}]})
        if len(self.historico) > 30:
            self.historico = self.historico[-30:]

    def limpar_historico(self):
        self.historico = []
        self.acao_pendente = None


# ================== JANELA DO CHATBOT CTK ==================
class JanelaChat:
    def __init__(self, root, paletas, obter_tema_atual, funcoes_disponiveis=None):
        self.root = root
        self.paletas = paletas
        self.obter_tema_atual = obter_tema_atual
        self.funcoes_disponiveis = funcoes_disponiveis or {}
        self.chatbot = ChatBot()
        self._fechando = False

        self.janela = ctk.CTkToplevel(root)
        self.janela.title("N.E.X.U.S AI Chat")
        self.janela.geometry("480x720")
        self.janela.minsize(420, 520)
        self.janela.protocol("WM_DELETE_WINDOW", self._fechar)

        self._construir_interface()
        self._enviar_boas_vindas()

    def _construir_interface(self):
        bg_color = "#090d16"
        header_color = "#111827"
        card_color = "#1a2436"
        self.janela.configure(fg_color=bg_color)

        # 1. TOPO / HEADER (Top)
        frame_topo = ctk.CTkFrame(self.janela, fg_color=header_color, corner_radius=0, height=55)
        frame_topo.pack(side="top", fill="x")
        frame_topo.pack_propagate(False)

        lbl_header = ctk.CTkLabel(frame_topo, text="⚡ N.E.X.U.S AI Assistant", font=("Segoe UI", 15, "bold"), text_color="#00d2ff")
        lbl_header.pack(side="left", padx=15, pady=10)

        btn_limpar = ctk.CTkButton(
            frame_topo, text="🗑️", width=34, height=34, font=("Segoe UI", 14),
            fg_color="transparent", hover_color="#333333", text_color="#a1a1aa",
            command=self._limpar_chat
        )
        btn_limpar.pack(side="right", padx=(0, 10), pady=10)

        btn_config = ctk.CTkButton(
            frame_topo, text="⚙️", width=34, height=34, font=("Segoe UI", 14),
            fg_color="transparent", hover_color="#333333",
            command=self._abrir_config
        )
        btn_config.pack(side="right", padx=(0, 4), pady=10)

        # 2. ÁREA DE ENTRADA (Bottom - empacotado ANTES do frame expansível do meio)
        frame_input_container = ctk.CTkFrame(self.janela, fg_color=header_color, corner_radius=0)
        frame_input_container.pack(side="bottom", fill="x")

        lbl_hint = ctk.CTkLabel(
            frame_input_container, text="Dica: Enter envia • Shift+Enter pula linha • Cole múltiplos seriais ou lotes",
            font=("Segoe UI", 9), text_color="#a1a1aa"
        )
        lbl_hint.pack(side="bottom", pady=(0, 5))

        input_box = ctk.CTkFrame(frame_input_container, fg_color=card_color, corner_radius=12, border_width=1, border_color="#3f3f46")
        input_box.pack(fill="x", expand=True, padx=12, pady=(8, 4))

        self.txt_msg = ctk.CTkTextbox(
            input_box, font=("Segoe UI", 13), height=65,
            fg_color="transparent", text_color="white", wrap="word"
        )
        self.txt_msg.pack(side="left", fill="both", expand=True, padx=(8, 4), pady=6)
        self.txt_msg.bind("<Return>", self._ao_pressionar_enter)
        self.txt_msg.bind("<Shift-Return>", self._ao_pressionar_shift_enter)
        self.txt_msg.bind("<Control-Return>", self._ao_pressionar_shift_enter)

        self.btn_enviar = ctk.CTkButton(
            input_box, text="➤", width=42, height=42, font=("Segoe UI", 18, "bold"),
            fg_color="#f57c00", hover_color="#e65100", text_color="white", corner_radius=21,
            command=self._enviar_mensagem
        )
        self.btn_enviar.pack(side="right", padx=8, pady=8, anchor="s")

        # 3. MENSAGENS / SCROLL (Middle - preenche todo o espaço restante)
        self.frame_mensagens = ctk.CTkScrollableFrame(self.janela, fg_color=bg_color, corner_radius=0)
        self.frame_mensagens.pack(side="top", fill="both", expand=True, padx=0, pady=0)

        self.janela.after(100, lambda: self.txt_msg.focus_set())

    def _ao_pressionar_enter(self, event):
        if event.state & 0x0001 or event.state & 0x0004:
            return None
        self._enviar_mensagem()
        return "break"

    def _ao_pressionar_shift_enter(self, event):
        return None

    def _limpar_chat(self):
        for w in self.frame_mensagens.winfo_children():
            w.destroy()
        self.chatbot.limpar_historico()
        self._enviar_boas_vindas()

    def _enviar_boas_vindas(self):
        self._adicionar_mensagem_bot(
            "👋 **Hub Assistant AI • N.E.X.U.S**\n\n"
            "⚡ **Agente Orquestrador Ativo**: Dispare qualquer operação do sistema diretamente pelo chat sem precisar de API!\n\n"
            "**Skills disponíveis para automação:**\n"
            "• 🚩 **Bandeirada**: `bandeira o csc300` ou `bandeirar csc300, csc301`\n"
            "• 📤 **Saída OSME**: `processa na osme do cvc395 330702 50m`\n"
            "• 📄 **OSME**: `extrair osme csc300` ou `osme csc300`\n"
            "• 🔍 **MIGO Clandestinas**: `migo rcl002` ou `consultar migo rcl001`\n"
            "• 🔄 **Mudança de Recurso**: `mudar recurso do csc395 para cvc396`\n"
            "• 📦 **Transferência OSME**: `transferir osme csc300`\n"
            "• 📋 **Conferência & Fichas**: `conferir seriais` ou `ficha re 19595`\n"
            "• 🚀 **Navegação**: `abrir etiquetas`, `abrir sap`, `abrir extrator`\n"
            "• 🔑 **SAP**: `sap Adriana` ou `info sap`\n"
            "• 📦 **Lotes & Estoque**: `CL03-002` ou `listar lotes`\n"
            "• 🔍 **Materiais & Seriais**: `material 18398041` ou `18398041 a 18398050`\n\n"
            "*(Dica: digite `ajuda` para ver o manual completo)*"
        )

    def _adicionar_mensagem_bot(self, texto):
        if self._fechando or not self.janela.winfo_exists():
            return
        frame = ctk.CTkFrame(self.frame_mensagens, fg_color="transparent")
        frame.pack(fill="x", padx=12, pady=6)

        avatar = ctk.CTkLabel(
            frame, text="✨", font=("Segoe UI", 16), width=32, height=32,
            corner_radius=16, fg_color="#0284c7", text_color="white"
        )
        avatar.pack(side="left", anchor="n", padx=(0, 8))

        bubble = ctk.CTkFrame(frame, fg_color="#27272a", corner_radius=12)
        bubble.pack(side="left", fill="x", expand=True)

        header = ctk.CTkLabel(bubble, text="Hub Assistant", font=("Segoe UI", 11, "bold"), text_color="#38bdf8")
        header.pack(anchor="w", padx=12, pady=(8, 2))

        lbl_conteudo = ctk.CTkLabel(
            bubble, text=texto, font=("Segoe UI", 13), text_color="#f4f4f5",
            wraplength=340, justify="left"
        )
        lbl_conteudo.pack(anchor="w", padx=12, pady=(0, 8))

        self.janela.after(50, lambda: self._rolar_ao_fim())

    def _adicionar_mensagem_usuario(self, texto):
        if self._fechando or not self.janela.winfo_exists():
            return
        frame = ctk.CTkFrame(self.frame_mensagens, fg_color="transparent")
        frame.pack(fill="x", padx=12, pady=6)

        bubble = ctk.CTkFrame(frame, fg_color="#f57c00", corner_radius=12)
        bubble.pack(side="right", anchor="e")

        lbl_conteudo = ctk.CTkLabel(
            bubble, text=texto, font=("Segoe UI", 13), text_color="white",
            wraplength=320, justify="left"
        )
        lbl_conteudo.pack(padx=12, pady=8)

        self.janela.after(50, lambda: self._rolar_ao_fim())

    def _rolar_ao_fim(self):
        try:
            if hasattr(self, "frame_mensagens") and self.frame_mensagens.winfo_exists():
                self.frame_mensagens._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass

    def _mostrar_pensando(self):
        if self._fechando or not self.janela.winfo_exists():
            return
        self.frame_pensando = ctk.CTkFrame(self.frame_mensagens, fg_color="transparent")
        self.frame_pensando.pack(fill="x", padx=12, pady=6)

        avatar = ctk.CTkLabel(
            self.frame_pensando, text="⏳", font=("Segoe UI", 16), width=32, height=32,
            corner_radius=16, fg_color="#3f3f46", text_color="white"
        )
        avatar.pack(side="left", anchor="n", padx=(0, 8))

        bubble = ctk.CTkFrame(self.frame_pensando, fg_color="#27272a", corner_radius=12)
        bubble.pack(side="left", fill="x", expand=True)

        self.lbl_pensando = ctk.CTkLabel(bubble, text="Processando solicitação...", font=("Segoe UI", 12, "italic"), text_color="#a1a1aa")
        self.lbl_pensando.pack(anchor="w", padx=12, pady=8)

        self._animacao_pensando = 0
        self._animar_pensando()
        self.janela.after(50, self._rolar_ao_fim)

    def _animar_pensando(self):
        if self._fechando or not hasattr(self, 'frame_pensando') or not self.frame_pensando.winfo_exists():
            return
        pontos = "." * ((self._animacao_pensando % 4) + 1)
        self.lbl_pensando.configure(text=f"Processando solicitação{pontos}")
        self._animacao_pensando += 1
        self.janela.after(350, self._animar_pensando)

    def _remover_pensando(self):
        if hasattr(self, 'frame_pensando') and self.frame_pensando.winfo_exists():
            self.frame_pensando.destroy()

    def _enviar_mensagem(self, event=None):
        texto = self.txt_msg.get("1.0", "end").strip()
        if not texto:
            return
        self.txt_msg.delete("1.0", "end")
        self._adicionar_mensagem_usuario(texto)
        self.btn_enviar.configure(state="disabled")
        self.txt_msg.configure(state="disabled")
        self._mostrar_pensando()
        threading.Thread(target=self._processar_em_thread, args=(texto,), daemon=True).start()

    def _processar_em_thread(self, texto):
        try:
            resposta = self.chatbot.processar_mensagem(texto, self.funcoes_disponiveis)
        except Exception as e:
            resposta = f"❌ Erro inesperado: {str(e)}"
        finally:
            def _atualizar_ui():
                if self._fechando or not self.janela.winfo_exists():
                    return
                self._remover_pensando()
                self._adicionar_mensagem_bot(resposta)
                self.btn_enviar.configure(state="normal")
                self.txt_msg.configure(state="normal")
                self.txt_msg.focus_set()

            try:
                self.janela.after(0, _atualizar_ui)
            except Exception:
                pass

    def _abrir_config(self):
        janela_cfg = ctk.CTkToplevel(self.janela)
        janela_cfg.title("Configurações do Chatbot")
        janela_cfg.geometry("460x420")
        janela_cfg.configure(fg_color="#18181b")
        janela_cfg.transient(self.janela)
        janela_cfg.grab_set()

        ctk.CTkLabel(janela_cfg, text="Configuração de IA", font=("Segoe UI", 16, "bold"), text_color="#38bdf8").pack(pady=(20, 4))
        ctk.CTkLabel(janela_cfg, text="Escolha entre o Motor Local Integrado (Sem API) ou Nuvem (OpenRouter).", font=("Segoe UI", 11), text_color="#a1a1aa").pack(pady=(0, 15))

        frame_campos = ctk.CTkFrame(janela_cfg, fg_color="#27272a", corner_radius=10)
        frame_campos.pack(fill="x", padx=20, pady=5)

        ctk.CTkLabel(frame_campos, text="Modelo de IA:", font=("Segoe UI", 12, "bold"), text_color="white").pack(anchor="w", padx=15, pady=(12, 4))
        modelo_var = ctk.StringVar(value=self.chatbot.obter_modelo())
        opcoes = [m["id"] for m in MODELOS_OPENROUTER]
        if modelo_var.get() not in opcoes:
            opcoes.insert(0, modelo_var.get())

        combo_modelo = ctk.CTkOptionMenu(
            frame_campos, variable=modelo_var, values=opcoes, font=("Segoe UI", 12),
            height=38, fg_color="#18181b", button_color="#0284c7", button_hover_color="#0369a1"
        )
        combo_modelo.pack(fill="x", padx=15, pady=(0, 12))

        ctk.CTkLabel(frame_campos, text="API Key OpenRouter (Opcional):", font=("Segoe UI", 12, "bold"), text_color="white").pack(anchor="w", padx=15, pady=(4, 4))
        entry_key = ctk.CTkEntry(
            frame_campos, font=("Consolas", 12), height=38,
            fg_color="#18181b", border_color="#3f3f46", text_color="white", show="*"
        )
        entry_key.pack(fill="x", padx=15, pady=(0, 6))
        entry_key.insert(0, self.chatbot.obter_api_key())

        ctk.CTkLabel(
            frame_campos,
            text="💡 Chaves OpenRouter começam com 'sk-or-v1-...' (openrouter.ai)\nDeixe em branco para usar o Motor Local Integrado sem API.",
            font=("Segoe UI", 10), text_color="#a1a1aa", justify="left"
        ).pack(anchor="w", padx=15, pady=(0, 14))

        def salvar():
            nova_chave = entry_key.get().strip()
            novo_modelo = modelo_var.get()
            self.chatbot.config["api_key"] = nova_chave
            self.chatbot.config["modelo"] = novo_modelo
            self.chatbot.salvar_config()
            self.chatbot.contador = 0
            self.chatbot.erros = 0
            janela_cfg.destroy()
            self._adicionar_mensagem_bot(f"✓ Configurações salvas! Modo ativo: `{novo_modelo}`.")

        ctk.CTkButton(
            janela_cfg, text="💾 Salvar Configurações", font=("Segoe UI", 13, "bold"),
            fg_color="#10b981", hover_color="#059669", text_color="white",
            height=40, corner_radius=10, command=salvar
        ).pack(fill="x", padx=20, pady=20)

    def _fechar(self):
        self._fechando = True
        try:
            self.janela.destroy()
        except Exception:
            pass

    def aplicar_tema(self, tema_nome):
        pass
