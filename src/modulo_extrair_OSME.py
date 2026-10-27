import os
import re
import json
import datetime
import subprocess
import tempfile
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk

import pythoncom
import win32com.client
import win32con
import win32gui

from modulo_Sap import verificar_sap_aberto, set_sap_ocupado
from utils_os import obter_diretorio_executavel, obter_caminho_recurso, obter_caminho_script, obter_caminho_config
from utils_metrics import incrementar_metrica

DEFAULT_PALETAS = {
    "Nexus": {"bg": "#090d16", "container": "#111827", "texto": "#f8fafc", "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"},
    "Azul Delray": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"},
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"}
}

def obter_script_osme():
    return obter_caminho_script("Sciprt OSME2.vbs")

SCRIPT_OSME = obter_script_osme()

class AbaExtrairOSME:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._renderizar_historico()
        self._log("📄 Aba Extrator OSME pronta.")

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

        self.btn_extrair.configure(fg_color=accent, hover_color=hover, text_color="#090d16" if "Light" not in tema_nome else "#ffffff")
        self.btn_limpar_log.configure(fg_color=container, hover_color=hover, text_color=texto)
        self.btn_limpar_hist.configure(fg_color=inp, hover_color=hover, text_color=texto)
        self.check_imprimir_excel.configure(fg_color=accent, text_color=texto)

        self.lbl_titulo.configure(text_color=texto)
        self.lbl_log.configure(text_color=texto)
        self.lbl_hist_tit.configure(text_color=accent)
        self.lbl_hist_hint.configure(text_color="#94a3b8")

        self.card_hist.configure(fg_color=container)

        estilo = ttk.Style()
        if "clam" in estilo.theme_names():
            estilo.theme_use("clam")
        estilo.configure(
            "HistoricoOSME.Treeview",
            background=container,
            foreground=texto,
            fieldbackground=inp,
            borderwidth=0,
            rowheight=26,
            relief="flat"
        )
        estilo.configure(
            "HistoricoOSME.Treeview.Heading",
            background=inp,
            foreground=accent,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            borderwidth=1
        )
        estilo.map(
            "HistoricoOSME.Treeview",
            background=[("selected", accent)],
            foreground=[("selected", "white")]
        )
        estilo.map(
            "HistoricoOSME.Treeview.Heading",
            background=[("active", hover)],
            foreground=[("active", "white")]
        )

    # ---------------------------------------------------------
    # Persistência de Histórico Recente
    # ---------------------------------------------------------
    def _caminho_historico(self):
        return obter_caminho_config("historico_extracao_osme.json")

    def _carregar_historico(self):
        caminho = self._caminho_historico()
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    if isinstance(dados, list):
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

    def _adicionar_historico(self, recurso, status):
        hora = datetime.datetime.now().strftime("%H:%M")
        itens = self._carregar_historico()
        itens.insert(0, {
            "recurso": recurso,
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
            self.tree_hist.insert("", "end", values=(d.get("recurso", "-"), d.get("horario", "-"), d.get("status", "-")))

    def _ao_clique_duplo_historico(self, event):
        selecionado = self.tree_hist.selection()
        if selecionado:
            valores = self.tree_hist.item(selecionado[0], "values")
            if valores:
                rec = valores[0]
                conteudo_atual = self.txt_recursos.get("1.0", tk.END).strip()
                if conteudo_atual:
                    if rec not in conteudo_atual.splitlines():
                        self.txt_recursos.insert(tk.END, f"\n{rec}")
                else:
                    self.txt_recursos.insert(tk.END, rec)

    # ---------------------------------------------------------
    # Execução da Extração
    # ---------------------------------------------------------
    def iniciar_extracao(self):
        if self.executando:
            messagebox.showwarning("Aviso", "A extração OSME já está em execução.")
            return

        recursos = [linha.strip() for linha in self.txt_recursos.get("1.0", tk.END).splitlines() if linha.strip()]
        if not recursos:
            messagebox.showwarning("Aviso", "Insira ao menos uma equipe/recurso.")
            return

        script_osme = obter_script_osme()
        if not os.path.exists(script_osme):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_osme}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para realizar a extração OSME.\nAbra o SAP e tente novamente.")
            return

        self.executando = True
        self.btn_extrair.configure(state="disabled")
        self._status("Extração OSME em execução...")
        self._log("\n📄 Iniciando extração OSME...")

        thread = threading.Thread(target=self._executar_recursos, args=(recursos,), daemon=True)
        thread.start()

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")
        self._log("🧹 Log limpo.")

    def _executar_recursos(self, recursos):
        set_sap_ocupado(True)
        try:
            from modulo_Sap import is_abortar_solicitado, registrar_processo_sap, limpar_processo_sap
            from utils_som import emitir_alerta_sonoro
        except Exception:
            is_abortar_solicitado = lambda: False
            registrar_processo_sap = lambda p: None
            limpar_processo_sap = lambda: None
            emitir_alerta_sonoro = lambda s: None

        try:
            for recurso in recursos:
                if is_abortar_solicitado():
                    self._log_threadsafe("\n🛑 Processo interrompido pelo usuário.")
                    break

                self._log_threadsafe(f"\n⏳ Consultando OSME para: {recurso}...")
                script_temp = self._criar_script_temporario(recurso)

                try:
                    self._log_threadsafe("🔍 Executando extração OSME no SAP...")

                    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                    proc = subprocess.Popen(
                        ["cscript.exe", "//nologo", script_temp],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        creationflags=creationflags
                    )
                    registrar_processo_sap(proc)
                    stdout, stderr = proc.communicate()
                    limpar_processo_sap()

                    saida = (stdout or "").strip()
                    if saida:
                        for linha in saida.splitlines():
                            self._log_threadsafe(linha.strip())

                    if proc.returncode == 0:
                        if self.imprimir_excel_var.get():
                            self._log_threadsafe("🖨️ Enviando planilha do Excel para impressão...")
                            self._imprimir_excel_aberto()
                        self._log_threadsafe(f"✅ Extração OSME finalizada para: {recurso}")
                        try:
                            incrementar_metrica("osmes_extraidos", 1)
                        except Exception:
                            pass
                        self.parent_frame.after(0, lambda r=recurso: self._adicionar_historico(r, "✅ Concluído"))
                        emitir_alerta_sonoro(True)
                    else:
                        erro = (stderr or stdout or "Erro desconhecido.").strip()
                        self._log_threadsafe(f"⚠️ Falha na extração OSME de {recurso}: {erro}")
                        self.parent_frame.after(0, lambda r=recurso: self._adicionar_historico(r, "❌ Falha"))
                        emitir_alerta_sonoro(False)
                finally:
                    try: os.remove(script_temp)
                    except OSError: pass

            self._status_threadsafe("Extração OSME finalizada.")
            self._log_threadsafe("\n✅ Processo OSME concluído.")
        except subprocess.TimeoutExpired:
            self._log_threadsafe("⚠️ Tempo limite atingido ao executar o script OSME.")
            self._status_threadsafe("Extração OSME interrompida por tempo limite.")
        except Exception as e:
            self._log_threadsafe(f"⚠️ Erro inesperado: {e}")
            self._status_threadsafe("Erro na extração OSME.")
        finally:
            set_sap_ocupado(False)
            self.parent_frame.after(0, self._finalizar_execucao)

    def _criar_script_temporario(self, recurso):
        script_osme = obter_script_osme()
        with open(script_osme, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()

        conteudo = self._inserir_reset_sap(conteudo)

        # 1. Substitui campo ctxtS_EQUIPE-LOW da ZUDWM_237
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtS_EQUIPE-LOW"\)\.text\s*=\s*)".*?"',
            lambda match: f'{match.group(1)}"{recurso}"',
            conteudo, count=1, flags=re.IGNORECASE,
        )
        conteudo = re.sub(
            r'(session\.findById\("wnd\[0\]/usr/ctxtS_EQUIPE-LOW"\)\.caretPosition\s*=\s*)\d+',
            lambda match: f"{match.group(1)}{len(recurso)}",
            conteudo, count=1, flags=re.IGNORECASE,
        )

        # 2. Fallback caso o script use ctxtDY_PARNR (IW39)
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

        # 3. Fallback para tag {RECURSO}
        conteudo = re.sub(r'\{RECURSO\}', recurso, conteudo, flags=re.IGNORECASE)

        fd, caminho_temp = tempfile.mkstemp(prefix="sap_osme_", suffix=".vbs")
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
        return conteudo

    def _imprimir_excel_aberto(self):
        def _acao():
            pythoncom.CoInitialize()
            try:
                # Tenta localizar a instância do Excel com retry por até 20 segundos
                excel = None
                inicio = time.time()
                while time.time() - inicio < 20:
                    try:
                        excel = win32com.client.GetObject(Class="Excel.Application")
                        if excel and excel.Workbooks.Count > 0:
                            break
                    except Exception:
                        pass
                    time.sleep(1)

                if excel and excel.Workbooks.Count > 0:
                    try:
                        wb = excel.ActiveWorkbook
                        if wb:
                            wb.PrintOut()
                            self._log_threadsafe("🖨️ Comando de impressão enviado com sucesso ao Excel.")
                        else:
                            self._log_threadsafe("⚠️ Nenhuma pasta de trabalho ativa encontrada no Excel.")
                    except Exception as e:
                        self._log_threadsafe(f"⚠️ Erro ao disparar PrintOut no Excel: {e}")
                else:
                    self._log_threadsafe("⚠️ Excel não localizado a tempo para impressão automática.")
            except Exception as e:
                self._log_threadsafe(f"⚠️ Não foi possível imprimir automaticamente: {e}")
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=_acao, daemon=True).start()


    def _esperar_janela_excel(self, timeout=15, intervalo=0.5):
        inicio = time.time()
        while time.time() - inicio < timeout:
            janelas = []
            def enum_cb(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    cls = win32gui.GetClassName(hwnd)
                    title = win32gui.GetWindowText(hwnd)
                    if "XLMAIN" in cls or "Excel" in title or "XLS" in title or "ALV" in title:
                        janelas.append(hwnd)
            win32gui.EnumWindows(enum_cb, None)
            if janelas: return janelas[0]
            time.sleep(intervalo)
        return None

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

        self.lbl_titulo = ctk.CTkLabel(col_esq, text="📄 Recursos / Equipes para extrair OSME (um por linha):", font=("Segoe UI", 13, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 6))

        self.txt_recursos = ctk.CTkTextbox(col_esq, height=110, font=("Consolas", 13), wrap="none", corner_radius=8)
        self.txt_recursos.pack(fill="x", pady=(0, 10))

        self.frame_botoes = ctk.CTkFrame(col_esq, fg_color="transparent")
        self.frame_botoes.pack(fill="x", pady=(0, 12))

        self.btn_extrair = ctk.CTkButton(self.frame_botoes, text="📄 Iniciar extração OSME", command=self.iniciar_extracao, font=("Segoe UI", 13, "bold"), height=38, corner_radius=8)
        self.btn_extrair.pack(side="left", fill="x", expand=True)

        self.imprimir_excel_var = tk.BooleanVar(value=True)
        self.check_imprimir_excel = ctk.CTkCheckBox(self.frame_botoes, text="🖨️ Imprimir Excel", variable=self.imprimir_excel_var, font=("Segoe UI", 11, "bold"))
        self.check_imprimir_excel.pack(side="left", padx=(10, 0))

        self.btn_limpar_log = ctk.CTkButton(self.frame_botoes, text="🧹 Limpar", command=self.limpar_log, font=("Segoe UI", 11, "bold"), width=90, height=38, corner_radius=8)
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

        self.lbl_hist_tit = ctk.CTkLabel(frame_hist_header, text="🕒 Histórico de Extrações", font=("Segoe UI", 13, "bold"))
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

        colunas = ("recurso", "horario", "status")
        self.tree_hist = ttk.Treeview(
            frame_tree,
            columns=colunas,
            show="headings",
            style="HistoricoOSME.Treeview",
            selectmode="browse"
        )
        self.tree_hist.heading("recurso", text="Recurso")
        self.tree_hist.heading("horario", text="Hora")
        self.tree_hist.heading("status", text="Status")

        self.tree_hist.column("recurso", width=120, anchor="center")
        self.tree_hist.column("horario", width=70, anchor="center")
        self.tree_hist.column("status", width=120, anchor="center")

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
        self.btn_extrair.configure(state="normal")

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
        return self.paletas.get(tema, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

def construir_aba_extrair_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaExtrairOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
