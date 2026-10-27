import os
import re
import json
import datetime
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk

from modulo_Sap import verificar_sap_aberto, set_sap_ocupado
from utils_os import obter_caminho_script, obter_caminho_config
from utils_metrics import incrementar_metrica

DEFAULT_PALETAS = {
    "Nexus": {"bg": "#090d16", "container": "#111827", "texto": "#f8fafc", "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"},
    "Azul Delray": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"},
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"}
}

def obter_script_bandeirada():
    return obter_caminho_script("SCRIPT DA BANDEIRADA.vbs")

SCRIPT_BANDEIRADA = obter_script_bandeirada()

def formatar_data_vigencia(data_str):
    """Converte e sanitiza datas retornadas do SAP (ex: '20.11.2026', '20261120') para 'DD/MM/YYYY'."""
    if not data_str or str(data_str).strip() in ("—", "-", "None", ""):
        return "—"
    s = str(data_str).strip()
    if any(p in s.lower() for p in ("fim", "base", "data", "ordem", "caufvd", "lbl", "ctxt")):
        return "—"
    m = re.match(r"^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})", s)
    if m:
        d, mes, a = m.groups()
        return f"{int(d):02d}/{int(mes):02d}/{a}"
    m_iso = re.match(r"^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})", s)
    if m_iso:
        a, mes, d = m_iso.groups()
        return f"{int(d):02d}/{int(mes):02d}/{a}"
    m_raw = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m_raw:
        a, mes, d = m_raw.groups()
        return f"{d}/{mes}/{a}"
    return s

class AbaBandeirada:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._renderizar_historico()
        self._log("🚩 Aba BANDEIRADA pronta.")

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        bg = c.get("bg", "#090d16")
        container = c.get("container", "#111827")
        inp = c.get("input", "#1a2436")
        texto = c.get("texto", "#f8fafc")
        accent = c.get("accent", "#00d2ff")
        hover = c.get("hover", "#0284c7")

        self.parent_frame.configure(fg_color=bg)

        for widget in [self.txt_recursos, self.txt_log]:
            widget.configure(fg_color=inp, text_color=texto, border_color=container, border_width=1)

        self.btn_iniciar.configure(fg_color=accent, hover_color=hover, text_color="#090d16" if "Light" not in tema_nome else "#ffffff")
        self.btn_limpar_log.configure(fg_color=container, hover_color=hover, text_color=texto)
        self.btn_limpar_hist.configure(fg_color=inp, hover_color=hover, text_color=texto)

        self.lbl_titulo.configure(text_color=texto)
        self.lbl_log.configure(text_color=texto)
        self.lbl_hist_tit.configure(text_color=accent)
        self.lbl_hist_hint.configure(text_color="#94a3b8")

        self.card_hist.configure(fg_color=container)

        estilo = ttk.Style()
        if "clam" in estilo.theme_names():
            estilo.theme_use("clam")
        estilo.configure(
            "HistoricoBandeirada.Treeview",
            background=container,
            foreground=texto,
            fieldbackground=inp,
            bordercolor=container,
            borderwidth=0,
            rowheight=26,
            relief="flat"
        )
        estilo.configure(
            "HistoricoBandeirada.Treeview.Heading",
            background=inp,
            foreground=accent,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            borderwidth=1
        )
        estilo.map(
            "HistoricoBandeirada.Treeview",
            background=[("selected", accent)],
            foreground=[("selected", "white")]
        )
        estilo.map(
            "HistoricoBandeirada.Treeview.Heading",
            background=[("active", hover)],
            foreground=[("active", "white")]
        )

    # ---------------------------------------------------------
    # Persistência de Histórico Recente
    # ---------------------------------------------------------
    def _caminho_historico(self):
        return obter_caminho_config("historico_bandeirada.json")

    def _carregar_historico(self):
        caminho = self._caminho_historico()
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    if isinstance(dados, list):
                        # Padroniza todos os recursos em maiúsculo e garante campo de vigência
                        for d in dados:
                            if isinstance(d, dict):
                                if "recurso" in d:
                                    d["recurso"] = str(d["recurso"]).strip().upper()
                                if "vigencia" not in d:
                                    d["vigencia"] = "—"
                        return dados
            except Exception:
                pass
        return []

    def _salvar_historico(self, itens):
        caminho = self._caminho_historico()
        try:
            os.makedirs(os.path.dirname(caminho), exist_ok=True)
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(itens[:50], f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _adicionar_historico(self, recurso, ordem, vigencia, status):
        hora = datetime.datetime.now().strftime("%H:%M")
        itens = self._carregar_historico()
        itens.insert(0, {
            "recurso": str(recurso).strip().upper(),
            "ordem": str(ordem).strip(),
            "vigencia": str(vigencia).strip() if vigencia else "—",
            "horario": hora,
            "status": status
        })
        self._salvar_historico(itens)
        self._renderizar_historico(itens)

    def _limpar_historico(self):
        self._salvar_historico([])
        self._renderizar_historico([])

    def _renderizar_historico(self, itens=None):
        if itens is None:
            itens = self._carregar_historico()
        for item in self.tree_hist.get_children():
            self.tree_hist.delete(item)
        for d in itens:
            rec = str(d.get("recurso", "-")).strip().upper()
            ordem = str(d.get("ordem", "-")).strip()
            vig = str(d.get("vigencia", "—")).strip() or "—"
            hora = str(d.get("horario", "-")).strip()
            st = str(d.get("status", "-")).strip()
            self.tree_hist.insert("", "end", values=(rec, ordem, vig, hora, st))

    def _ao_clique_duplo_historico(self, event):
        selecionado = self.tree_hist.selection()
        if selecionado:
            valores = self.tree_hist.item(selecionado[0], "values")
            if valores:
                rec = str(valores[0]).strip().upper()
                conteudo_atual = self.txt_recursos.get("1.0", tk.END).strip()
                if conteudo_atual:
                    if rec not in conteudo_atual.splitlines():
                        self.txt_recursos.insert(tk.END, f"\n{rec}")
                else:
                    self.txt_recursos.insert(tk.END, rec)

    # ---------------------------------------------------------
    # Execução da Bandeirada
    # ---------------------------------------------------------
    def iniciar_bandeirada(self):
        if self.executando:
            messagebox.showwarning("Aviso", "A bandeirada já está em execução.")
            return

        recursos = [linha.strip().upper() for linha in self.txt_recursos.get("1.0", tk.END).splitlines() if linha.strip()]
        if not recursos:
            messagebox.showwarning("Aviso", "Insira ao menos uma equipe/recurso.")
            return

        # Uniformiza texto para maiúsculo na interface
        self.txt_recursos.delete("1.0", tk.END)
        self.txt_recursos.insert("1.0", "\n".join(recursos))

        script_bandeirada = obter_script_bandeirada()
        if not os.path.exists(script_bandeirada):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_bandeirada}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para executar a Bandeirada.\nAbra o SAP e tente novamente.")
            return

        self.executando = True
        self.btn_iniciar.configure(state="disabled")
        self._status("Bandeirada em execução...")
        self._log("\n🏁 Iniciando bandeirada inteligente...")

        thread = threading.Thread(target=self._executar_recursos, args=(recursos,), daemon=True)
        thread.start()

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("🚩 Log limpo.")

    def _executar_recursos(self, recursos):
        set_sap_ocupado(True)
        try:
            for recurso in recursos:
                recurso = str(recurso).strip().upper()
                self._log_threadsafe(f"\n⏳ Pesquisando ORDEM para: {recurso}...")
                script_temp = self._criar_script_temporario(recurso)

                try:
                    self._log_threadsafe("🔍 Executando consulta no SAP para obter a ordem do recurso...")

                    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                    resultado = subprocess.run(
                        ["cscript.exe", "//nologo", script_temp],
                        capture_output=True, text=True, timeout=None, creationflags=creationflags,
                    )

                    saida = (resultado.stdout or "").strip()
                    if saida:
                        for linha in saida.splitlines():
                            self._log_threadsafe(linha.strip())

                    # Extração da Ordem e da Vigência
                    ordem_encontrada = "—"
                    vigencia_encontrada = "—"

                    # 1. Tenta pelo marcador estruturado BANDEIRADA_INFO ou BANDEIRADA_CONCLUIDA
                    match_info = re.search(r"BANDEIRADA_(?:INFO|CONCLUIDA):ORDEM=(\d+)(?:\|VENC=([^|\r\n]*))?", saida, re.IGNORECASE)
                    if match_info:
                        ordem_encontrada = match_info.group(1).strip()
                        if match_info.group(2):
                            vigencia_encontrada = formatar_data_vigencia(match_info.group(2).strip())
                    else:
                        # 2. Fallbacks abrangentes para capturar a ordem
                        padroes_ordem = [
                            r"Ordem encontrada.*?:\s*(\d+)",
                            r"Ordem mais recente encontrada:\s*(\d+)",
                            r"Bandeirada conclu[íi]da com sucesso para ordem:\s*(\d+)",
                            r"ordem:\s*(\d{8,12})",
                            r"\b(1000\d{8})\b"
                        ]
                        for padrao in padroes_ordem:
                            m = re.search(padrao, saida, re.IGNORECASE)
                            if m:
                                ordem_encontrada = m.group(1).strip()
                                break

                        # 3. Fallbacks para vigência
                        m_venc = re.search(r"Vig[êe]ncia.*?:\s*([0-9\.\/-]+)", saida, re.IGNORECASE)
                        if m_venc:
                            vigencia_encontrada = formatar_data_vigencia(m_venc.group(1).strip())

                    if resultado.returncode == 0:
                        msg_fim = f"✅ Bandeirada finalizada para: {recurso}"
                        if ordem_encontrada != "—":
                            msg_fim += f" (Ordem: {ordem_encontrada}"
                            if vigencia_encontrada != "—":
                                msg_fim += f" | Vigência: {vigencia_encontrada}"
                            msg_fim += ")"
                        self._log_threadsafe(msg_fim)
                        try:
                            incrementar_metrica("bandeiradas_executadas", 1)
                        except Exception:
                            pass
                        self.parent_frame.after(
                            0,
                            lambda r=recurso, o=ordem_encontrada, v=vigencia_encontrada: self._adicionar_historico(r, o, v, "✅ Concluído")
                        )
                    else:
                        erro = (resultado.stderr or resultado.stdout or "Erro desconhecido.").strip()
                        self._log_threadsafe(f"⚠️ Falha na bandeirada de {recurso}: {erro}")
                        self.parent_frame.after(
                            0,
                            lambda r=recurso, o=ordem_encontrada, v=vigencia_encontrada: self._adicionar_historico(r, o, v, "❌ Falha")
                        )
                finally:
                    try: os.remove(script_temp)
                    except OSError: pass

            self._status_threadsafe("Bandeirada finalizada.")
            self._log_threadsafe("\n✅ Processo de bandeirada concluído.")
        except subprocess.TimeoutExpired:
            self._log_threadsafe("⚠️ Tempo limite atingido ao executar o script VBS.")
            self._status_threadsafe("Bandeirada interrompida por tempo limite.")
        except Exception as e:
            self._log_threadsafe(f"⚠️ Erro inesperado: {e}")
            self._status_threadsafe("Erro na bandeirada.")
        finally:
            set_sap_ocupado(False)
            self.parent_frame.after(0, self._finalizar_execucao)

    def _criar_script_temporario(self, recurso):
        script_bandeirada = obter_script_bandeirada()
        with open(script_bandeirada, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        conteudo = self._inserir_reset_sap(conteudo)
        conteudo = re.sub(
            r'(recurso\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtDY_PARNR"\)\.text\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtDY_PARNR"\)\.caretPosition\s*=\s*)\d+',
            lambda match: f"{match.group(1)}{len(recurso)}",
            conteudo, count=1, flags=re.IGNORECASE,
        )

        fd, caminho_temp = tempfile.mkstemp(prefix="sap_bandeirada_", suffix=".vbs")
        with open(fd, "w", encoding="utf-8", errors="ignore") as f:
            f.write(conteudo)
        return caminho_temp

    def _inserir_reset_sap(self, conteudo):
        reset_cmd = (
            'On Error Resume Next\n'
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/n"\n'
            'session.findById("wnd[0]").sendVKey 0\n'
            'On Error GoTo 0\n'
        )
        alvo = 'If Not IsObject(connection) Then'
        if alvo in conteudo:
            partes = conteudo.split(alvo, 1)
            return partes[0] + alvo + partes[1].replace(
                'Set session = connection.Children(0)',
                'Set session = connection.Children(0)\n' + reset_cmd,
                1,
            )

        ponto_ancora = 'session.findById("wnd[0]/mbar/menu[0]/menu[7]").select'
        bloco_captura = (
            'Dim ordemBandeirada, txtOrdemDet\n'
            'ordemBandeirada = ""\n'
            'If Not IsObject(grid) Then\n'
            '    Set txtOrdemDet = Nothing\n'
            '    On Error Resume Next\n'
            '    Set txtOrdemDet = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR")\n'
            '    If txtOrdemDet Is Nothing Then\n'
            '        Set txtOrdemDet = session.findById("wnd[0]/usr/txtCAUFVD-AUFNR")\n'
            '    End If\n'
            '    If Not txtOrdemDet Is Nothing Then\n'
            '        ordemBandeirada = Trim(txtOrdemDet.text)\n'
            '    End If\n'
            '    On Error GoTo 0\n'
            'End If\n'
            'If Trim(ordemBandeirada) = "" Then\n'
            '    WScript.Echo "Nenhuma ordem encontrada para o recurso pesquisado."\n'
            '    WScript.Quit 2\n'
            'End If\n'
            'WScript.Echo "Ordem encontrada: " & ordemBandeirada & ". Bandeirando..."\n'
        )
        if ponto_ancora in conteudo:
            substituicao = bloco_captura + 'On Error Resume Next\n' + ponto_ancora
            return conteudo.replace(ponto_ancora, substituicao, 1)
        return bloco_captura + conteudo

    # ---------------------------------------------------------
    # Construção da Interface (2 Colunas: Operação & Histórico)
    # ---------------------------------------------------------
    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=16)

        self.container.grid_columnconfigure(0, weight=6)  # Coluna Esquerda: 60%
        self.container.grid_columnconfigure(1, weight=4)  # Coluna Direita: 40%
        self.container.grid_rowconfigure(0, weight=1)

        # 1. COLUNA ESQUERDA: Execução & Logs
        col_esq = ctk.CTkFrame(self.container, fg_color="transparent")
        col_esq.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.lbl_titulo = ctk.CTkLabel(col_esq, text="🚩 Recursos / Equipes (um por linha):", font=("Segoe UI", 13, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 6))

        self.txt_recursos = ctk.CTkTextbox(col_esq, height=110, font=("Consolas", 13), wrap="none", corner_radius=8)
        self.txt_recursos.pack(fill="x", pady=(0, 10))

        self.frame_botoes = ctk.CTkFrame(col_esq, fg_color="transparent")
        self.frame_botoes.pack(fill="x", pady=(0, 12))

        self.btn_iniciar = ctk.CTkButton(self.frame_botoes, text="🏁 Iniciar Bandeirada", command=self.iniciar_bandeirada, font=("Segoe UI", 13, "bold"), height=38, corner_radius=8)
        self.btn_iniciar.pack(side="left", fill="x", expand=True)

        self.btn_limpar_log = ctk.CTkButton(self.frame_botoes, text="🧹 Limpar Log", command=self.limpar_log, font=("Segoe UI", 11, "bold"), width=110, height=38, corner_radius=8)
        self.btn_limpar_log.pack(side="left", padx=(10, 0))

        self.lbl_log = ctk.CTkLabel(col_esq, text="📋 Log de Execução:", font=("Segoe UI", 12, "bold"))
        self.lbl_log.pack(anchor="w", pady=(0, 4))

        self.txt_log = ctk.CTkTextbox(col_esq, font=("Segoe UI", 11), wrap="word", state="disabled", corner_radius=8)
        self.txt_log.pack(fill="both", expand=True)

        # 2. COLUNA DIREITA: Painel de Histórico Recente
        self.card_hist = ctk.CTkFrame(self.container, corner_radius=10)
        self.card_hist.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        frame_hist_header = ctk.CTkFrame(self.card_hist, fg_color="transparent")
        frame_hist_header.pack(fill="x", padx=12, pady=(12, 8))

        self.lbl_hist_tit = ctk.CTkLabel(frame_hist_header, text="🕒 Histórico de Recursos", font=("Segoe UI", 13, "bold"))
        self.lbl_hist_tit.pack(side="left")

        self.btn_limpar_hist = ctk.CTkButton(
            frame_hist_header,
            text="🗑️ Limpar",
            width=70,
            height=26,
            font=("Segoe UI", 10, "bold"),
            command=self._limpar_historico,
            corner_radius=6
        )
        self.btn_limpar_hist.pack(side="right")

        # Tabela Treeview de Histórico
        frame_tree = ctk.CTkFrame(self.card_hist, fg_color="transparent")
        frame_tree.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        colunas = ("recurso", "ordem", "vigencia", "horario", "status")
        self.tree_hist = ttk.Treeview(
            frame_tree,
            columns=colunas,
            show="headings",
            style="HistoricoBandeirada.Treeview",
            selectmode="browse"
        )
        self.tree_hist.heading("recurso", text="Recurso")
        self.tree_hist.heading("ordem", text="Ordem")
        self.tree_hist.heading("vigencia", text="Vigência")
        self.tree_hist.heading("horario", text="Hora")
        self.tree_hist.heading("status", text="Status")

        self.tree_hist.column("recurso", width=70, anchor="center")
        self.tree_hist.column("ordem", width=105, anchor="center")
        self.tree_hist.column("vigencia", width=85, anchor="center")
        self.tree_hist.column("horario", width=50, anchor="center")
        self.tree_hist.column("status", width=90, anchor="center")

        scroll_y = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tree_hist.yview)
        self.tree_hist.configure(yscrollcommand=scroll_y.set)

        self.tree_hist.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        self.tree_hist.bind("<Double-1>", self._ao_clique_duplo_historico)

        self.lbl_hist_hint = ctk.CTkLabel(
            self.card_hist,
            text="💡 Duplo clique em um item para inseri-lo no campo.",
            font=("Segoe UI", 10, "italic")
        )
        self.lbl_hist_hint.pack(padx=12, pady=(0, 8))

    def _finalizar_execucao(self):
        self.executando = False
        self.btn_iniciar.configure(state="normal")

    def _log(self, texto):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, texto + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._log(texto))

    def _status(self, texto):
        if self.label_status_global:
            self.label_status_global.configure(text=f" {texto}")

    def _status_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._status(texto))

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, next(iter(self.paletas.values())))

def construir_aba_bandeirada(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaBandeirada(parent_frame, label_status_global, paletas, obter_tema_atual)
