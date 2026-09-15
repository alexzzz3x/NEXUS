import os
import re
import json
import threading
import tkinter as tk
import customtkinter as ctk
import requests
from utils_os import obter_caminho_config, obter_diretorio_dados

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

SYSTEM_PROMPT = """Você é o assistente virtual do Conecta Hub 9.0, um sistema de gerenciamento de estoque, etiquetas e automação SAP da Conecta Empreendimentos.
Responda em português brasileiro, de forma amigável, clara e profissional.
Use ferramentas sempre que o usuário solicitar consulta de lotes, materiais, extração de seriais, login SAP ou geração de etiquetas/docs.
Quando o usuário enviar múltiplos seriais, múltiplos lotes ou múltiplos materiais, processe todos e forneça um resumo estruturado.
Confirme ações e exiba resultados com clareza."""

TOOLS = [
    {
        "functionDeclarations": [
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
    def _processar_mensagem_local(self, mensagem, funcoes_disponiveis):
        """
        Processador NLP integrado baseado em intenções e extração de entidades.
        Funciona 100% offline sem necessidade de chave API ou internet!
        """
        msg_norm = mensagem.strip()
        msg_lower = msg_norm.lower()

        if not funcoes_disponiveis:
            return "⚠️ Funções do sistema ainda não estão inicializadas. Aguarde um instante."

        # 1. Saudação / Ajuda
        if msg_lower in ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite", "help", "ajuda", "comandos", "?", "a"]:
            return (
                "🤖 **Assistente Inteligente Integrado (Conecta Hub 9.0)**\n\n"
                "Estou pronto para executar ações diretamente no sistema. Exemplos de comandos:\n\n"
                "• 📦 **Consultar Lote**: `CL03-002` ou `buscar lote CL03-002`\n"
                "• 📋 **Listar Lotes**: `listar lotes` ou `todos os lotes`\n"
                "• 🔍 **Consultar Material**: `material 18398041` ou `18398041`\n"
                "• 🔢 **Extrair Seriais**: `18398041 a 18398050` ou colar lista de seriais\n"
                "• 📄 **Google Docs**: `gerar doc CL03-002` ou `doc CL03-002`\n"
                "• 🏷️ **Gerar Etiqueta**: `etiqueta CL03 material 18398041 qtd 50`\n"
                "• 🔑 **SAP**: `login Anderson` ou `info sap`\n"
                "• 📋 **Copiar**: `copiar seriais`"
            )

        # 2. Informações do SAP
        if any(term in msg_lower for term in ["info sap", "status sap", "sap status", "sap info", "sessao sap"]):
            func = funcoes_disponiveis.get("info_sap")
            if func:
                res = func()
                return f"🔑 **Status da Conexão SAP**:\n`{res.get('info', 'Não foi possível ler o status.')}`"

        # 3. Login SAP
        match_login = re.search(r"(?:login|logar|conectar|usuario|usuário)\s*(?:no\s*sap|sap)?\s*(?:como\s*)?([a-zA-ZáéíóúÁÉÍÓÚãõÃÕçÇ]+)", msg_norm, re.IGNORECASE)
        if match_login and ("sap" in msg_lower or "login" in msg_lower or "logar" in msg_lower):
            usuario = match_login.group(1).strip()
            if usuario.lower() not in ["no", "em", "sap"]:
                func = funcoes_disponiveis.get("login_sap")
                if func:
                    res = func(usuario=usuario)
                    if "erro" in res:
                        return f"❌ {res['erro']}"
                    return f"🔑 **Login SAP**: {res.get('status', 'Iniciando login...')}"

        # 4. Copiar Seriais
        if msg_lower in ["copiar", "copiar seriais", "copiar serial", "copia"]:
            func = funcoes_disponiveis.get("copiar_seriais")
            if func:
                res = func()
                return f"📋 {res.get('status', 'Seriais copiados com sucesso!')}"

        # 5. Listar Lotes
        if any(term in msg_lower for term in ["listar lotes", "todos os lotes", "ver lotes", "quais lotes", "lotes disponiveis", "lotes estoque"]):
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

        # 6. Gerar Google Doc a partir de Lote
        match_doc = re.search(r"(?:doc|documento|drive|google)\s*(?:do\s*lote|lote)?\s*([a-zA-Z0-9]{2,5}[-_][0-9]{2,5})", msg_norm, re.IGNORECASE)
        if match_doc or ("doc" in msg_lower and "lote" in msg_lower):
            lote_doc = match_doc.group(1).strip().upper() if match_doc else ""
            if not lote_doc:
                achados = re.findall(r"\b[A-Z]{2,4}[-_]?\d{3,4}\b", msg_norm.upper())
                if achados:
                    lote_doc = achados[0]
            if lote_doc:
                func = funcoes_disponiveis.get("gerar_doc_google")
                if func:
                    res = func(lote=lote_doc)
                    if "erro" in res:
                        return f"❌ Erro ao gerar Google Doc para lote `{lote_doc}`:\n{res['erro']}"
                    return f"📄 **Google Docs Criado com Sucesso!**\n• Lote: `{lote_doc}`\n• Quantidade: `{res.get('quantidade', '')}` seriais\n• Link de Acesso: {res.get('link', '')}"

        # 7. Extração de Seriais (Intervalo ou Lista de Seriais)
        tem_intervalo = re.search(r"\d+\s*(?:a|à|-|ate|até)\s*\d+", msg_norm, re.IGNORECASE)
        achados_digitos = re.findall(r"\b\d{6,18}\b", msg_norm)
        if tem_intervalo or len(achados_digitos) >= 3 or ("extrair" in msg_lower and achados_digitos):
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

        # 8. Consulta de Lote(s)
        achados_lotes = re.findall(r"\b(CL\d{2}[-_]\d{3,4}|CS\d{2}[-_]\d{3,4}|[A-Z]{2,4}[-_]\d{3,4})\b", msg_norm.upper())
        if achados_lotes:
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

        # 9. Consulta de Material(is)
        match_mat = re.findall(r"\b(18\d{6}|[0-9]{7,8})\b", msg_norm)
        if match_mat and ("material" in msg_lower or "código" in msg_lower or "codigo" in msg_lower or len(match_mat) == 1):
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

        # 10. Resposta Padrão Amigável com Reconhecimento
        return (
            f"🤖 Entendido! Você digitou: *\"{msg_norm}\"*\n\n"
            "Não identifiquei uma ação direta ou o formato não foi reconhecido. Como posso te ajudar?\n"
            "• Digite um código de lote (ex: `CL03-002`)\n"
            "• Cole seriais para extrair (ex: `18398041 a 18398050`)\n"
            "• Consulte um material (ex: `18398041`)\n"
            "• Peça login SAP (ex: `login Anderson`)\n"
            "• Digite `ajuda` para ver todos os comandos!"
        )

    def _executar_funcao(self, nome, args, funcoes):
        if not funcoes:
            return {"erro": "Funções do sistema não disponíveis"}
        func = funcoes.get(nome)
        if not func:
            return {"erro": f"Função '{nome}' não cadastrada"}
        try:
            return func(**args)
        except TypeError as te:
            try:
                return func()
            except Exception:
                return {"erro": f"Parâmetros inválidos para {nome}: {str(te)}"}
        except Exception as e:
            return {"erro": str(e)}

    def _adicionar_historico(self, role, texto):
        self.historico.append({"role": role, "parts": [{"text": texto}]})
        if len(self.historico) > 30:
            self.historico = self.historico[-30:]

    def limpar_historico(self):
        self.historico = []


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
        self.janela.title("Conecta Hub Chat 9.0")
        self.janela.geometry("480x720")
        self.janela.minsize(420, 520)
        self.janela.protocol("WM_DELETE_WINDOW", self._fechar)

        self._construir_interface()
        self._enviar_boas_vindas()

    def _construir_interface(self):
        bg_color = "#18181b"
        header_color = "#09090b"
        card_color = "#27272a"
        self.janela.configure(fg_color=bg_color)

        # 1. TOPO / HEADER (Top)
        frame_topo = ctk.CTkFrame(self.janela, fg_color=header_color, corner_radius=0, height=55)
        frame_topo.pack(side="top", fill="x")
        frame_topo.pack_propagate(False)

        lbl_header = ctk.CTkLabel(frame_topo, text="✨ Hub Assistant AI 9.0", font=("Segoe UI", 15, "bold"), text_color="#38bdf8")
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
            "👋 **Assistente Inteligente Conecta Hub 9.0 Pronto!**\n\n"
            "⚡ **Modo Integrado Ativo**: Você não precisa de nenhuma chave de API para usar as funções do sistema!\n\n"
            "Posso te ajudar com:\n"
            "• 📦 **Lotes e Estoque**: `CL03-002` ou `listar lotes`\n"
            "• 🔢 **Seriais**: Cole qualquer bloco com seriais ou `18398041 a 18398050`\n"
            "• 🔍 **Materiais**: `material 18398041` ou apenas o código\n"
            "• 📄 **Google Docs**: `gerar doc CL03-002`\n"
            "• 🔑 **SAP**: `login Anderson` ou `info sap`\n\n"
            "*(Opcional: Se desejar respostas conversacionais com IA na nuvem, configure sua chave OpenRouter no ⚙️)*"
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
        entry_key.pack(fill="x", padx=15, pady=(0, 15))
        entry_key.insert(0, self.chatbot.obter_api_key())

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
