# -*- coding: utf-8 -*-
"""
Módulo de Consulta MIGO / IW39 — N.E.X.U.S
Especializado para Equipes Clandestinas (RCL001 a RCL008).
- Consulta a Ordem ativa na IW39 (Parceiro Z2).
- Consulta os materiais e reservas associados na MIGO (A07 Saída de Mercadorias / R08 Ordem).
- Destaca materiais alocados no centro EMPR.
- Permite abrir a tela diretamente no SAP GUI para conferência visual.
"""

import os
import sys
import time
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk

import pythoncom
import win32com.client
import win32gui

from modulo_Sap import verificar_sap_aberto, obter_sessao_ativa_sap, set_sap_ocupado
from utils_logger import log_error, log_info
from utils_ui import show_toast, play_sound

DEFAULT_PALETAS = {
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#f57c00",
        "hover": "#e65100",
        "input": "#1f3a60",
    }
}

EQUIPES_CLANDESTINAS = [
    "RCL001", "RCL002", "RCL003", "RCL004",
    "RCL005", "RCL006", "RCL007", "RCL008"
]


class AbaConsultarMigo:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.executando = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("🔍 Módulo Consultar MIGO (Clandestinas) pronto.")

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Conecta Blue", DEFAULT_PALETAS["Conecta Blue"]))

    def _status(self, msg):
        if self.label_status_global:
            try:
                self.label_status_global.configure(text=msg)
            except Exception:
                pass

    def _status_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._status(msg))

    def _log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, msg)
        if not msg.endswith("\n"):
            self.txt_log.insert(tk.END, "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._log(msg))

    def _construir_interface(self):
        self.parent_frame.grid_columnconfigure(0, weight=1)
        self.parent_frame.grid_rowconfigure(2, weight=1)

        # 1. CABEÇALHO & SELEÇÃO DE EQUIPE
        self.frame_topo = ctk.CTkFrame(self.parent_frame, corner_radius=8)
        self.frame_topo.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        lbl_header = ctk.CTkLabel(
            self.frame_topo,
            text="🔍 CONSULTAR MIGO — OPERAÇÃO CLANDESTINA (RCL001 a RCL008)",
            font=("Segoe UI", 14, "bold")
        )
        lbl_header.pack(anchor="w", padx=15, pady=(8, 2))

        lbl_sub = ctk.CTkLabel(
            self.frame_topo,
            text="Localiza a Ordem Ativa na IW39 para o recurso e consulta os materiais no centro EMPR na MIGO.",
            font=("Segoe UI", 11),
            text_color="#94a3b8"
        )
        lbl_sub.pack(anchor="w", padx=15, pady=(0, 6))

        # Atalhos rápidos de botões RCL001..RCL008
        self.frame_botoes_equipes = ctk.CTkFrame(self.frame_topo, fg_color="transparent")
        self.frame_botoes_equipes.pack(fill="x", padx=15, pady=(2, 8))

        lbl_atalho = ctk.CTkLabel(self.frame_botoes_equipes, text="Equipe:", font=("Segoe UI", 11, "bold"))
        lbl_atalho.pack(side="left", padx=(0, 8))

        self.btn_equipes = {}
        for eq in EQUIPES_CLANDESTINAS:
            btn = ctk.CTkButton(
                self.frame_botoes_equipes,
                text=eq,
                width=65,
                height=26,
                font=("Segoe UI", 11, "bold"),
                command=lambda e=eq: self._selecionar_equipe(e)
            )
            btn.pack(side="left", padx=3)
            self.btn_equipes[eq] = btn

        # 2. CONTROLES E INPUTS
        self.frame_inputs = ctk.CTkFrame(self.parent_frame, corner_radius=8)
        self.frame_inputs.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        # Linha 1: Consulta Individual (Foco Principal)
        f_in1 = ctk.CTkFrame(self.frame_inputs, fg_color="transparent")
        f_in1.pack(fill="x", padx=15, pady=(8, 4))

        lbl_indiv = ctk.CTkLabel(f_in1, text="Individual:", font=("Segoe UI", 11, "bold"))
        lbl_indiv.pack(side="left", padx=(0, 6))

        self.combo_equipe = ctk.CTkComboBox(
            f_in1,
            values=EQUIPES_CLANDESTINAS,
            width=105,
            height=28,
            font=("Segoe UI", 11, "bold"),
            command=self._ao_mudar_combo
        )
        self.combo_equipe.set("RCL001")
        self.combo_equipe.pack(side="left", padx=(0, 10))

        lbl_ord = ctk.CTkLabel(f_in1, text="Ordem IW39:", font=("Segoe UI", 11, "bold"))
        lbl_ord.pack(side="left", padx=(0, 5))

        self.entry_ordem = ctk.CTkEntry(
            f_in1,
            placeholder_text="Auto ou Digite...",
            width=135,
            height=28,
            font=("Segoe UI", 11)
        )
        self.entry_ordem.pack(side="left", padx=(0, 15))

        self.btn_completo = ctk.CTkButton(
            f_in1,
            text="⚡ Consultar (IW39 ➡️ MIGO)",
            width=180,
            height=28,
            font=("Segoe UI", 11, "bold"),
            fg_color="#059669",
            hover_color="#047857",
            command=self.iniciar_fluxo_completo
        )
        self.btn_completo.pack(side="left", padx=3)

        self.btn_buscar_ordem = ctk.CTkButton(
            f_in1,
            text="🔎 1. Buscar Ordem",
            width=135,
            height=28,
            font=("Segoe UI", 11),
            command=self.iniciar_busca_ordem_iw39
        )
        self.btn_buscar_ordem.pack(side="left", padx=3)

        self.btn_consultar_migo = ctk.CTkButton(
            f_in1,
            text="📦 2. Consultar MIGO",
            width=140,
            height=28,
            font=("Segoe UI", 11),
            command=self.iniciar_consulta_migo
        )
        self.btn_consultar_migo.pack(side="left", padx=3)

        self.btn_abrir_sap = ctk.CTkButton(
            f_in1,
            text="🖥️ Abrir no SAP",
            width=115,
            height=28,
            font=("Segoe UI", 11),
            command=self.abrir_migo_sap_gui
        )
        self.btn_abrir_sap.pack(side="left", padx=3)

        # Linha 2: Consulta em Faixa (Opção de X a X)
        f_in2 = ctk.CTkFrame(self.frame_inputs, fg_color="transparent")
        f_in2.pack(fill="x", padx=15, pady=(2, 8))

        lbl_faixa = ctk.CTkLabel(f_in2, text="Opção em Faixa (X a X):", font=("Segoe UI", 11, "bold"))
        lbl_faixa.pack(side="left", padx=(0, 8))

        lbl_de = ctk.CTkLabel(f_in2, text="De:", font=("Segoe UI", 11))
        lbl_de.pack(side="left", padx=(0, 4))

        self.combo_de = ctk.CTkComboBox(
            f_in2,
            values=EQUIPES_CLANDESTINAS,
            width=95,
            height=26,
            font=("Segoe UI", 11)
        )
        self.combo_de.set("RCL001")
        self.combo_de.pack(side="left", padx=(0, 10))

        lbl_ate = ctk.CTkLabel(f_in2, text="Até:", font=("Segoe UI", 11))
        lbl_ate.pack(side="left", padx=(0, 4))

        self.combo_ate = ctk.CTkComboBox(
            f_in2,
            values=EQUIPES_CLANDESTINAS,
            width=95,
            height=26,
            font=("Segoe UI", 11)
        )
        self.combo_ate.set("RCL008")
        self.combo_ate.pack(side="left", padx=(0, 15))

        self.btn_varrer_faixa = ctk.CTkButton(
            f_in2,
            text="🌐 Consultar Faixa (X a X)",
            width=175,
            height=26,
            font=("Segoe UI", 11, "bold"),
            command=self.iniciar_varredura_faixa
        )
        self.btn_varrer_faixa.pack(side="left", padx=4)

        self.btn_atalho_todas = ctk.CTkButton(
            f_in2,
            text="Todas (01 a 08)",
            width=115,
            height=26,
            font=("Segoe UI", 10),
            command=self._selecionar_todas_faixa
        )
        self.btn_atalho_todas.pack(side="left", padx=4)

        # 3. CORPO CENTRAL: TABELA E TERMINAL DE LOG
        self.frame_conteudo = ctk.CTkFrame(self.parent_frame, corner_radius=8)
        self.frame_conteudo.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
        self.frame_conteudo.grid_columnconfigure(0, weight=6)
        self.frame_conteudo.grid_columnconfigure(1, weight=4)
        self.frame_conteudo.grid_rowconfigure(0, weight=1)

        # 3.1 TABELA DE RESULTADOS (ESQUERDA)
        self.frame_tabela = ctk.CTkFrame(self.frame_conteudo, corner_radius=6)
        self.frame_tabela.grid(row=0, column=0, padx=(8, 4), pady=8, sticky="nsew")
        self.frame_tabela.grid_columnconfigure(0, weight=1)
        self.frame_tabela.grid_rowconfigure(1, weight=1)

        f_topo_tab = ctk.CTkFrame(self.frame_tabela, fg_color="transparent")
        f_topo_tab.grid(row=0, column=0, padx=8, pady=(6, 4), sticky="ew")

        self.lbl_qtd_itens = ctk.CTkLabel(
            f_topo_tab,
            text="Itens Encontrados na MIGO: 0",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_qtd_itens.pack(side="left")

        self.btn_copiar_tab = ctk.CTkButton(
            f_topo_tab,
            text="📋 Copiar Dados",
            width=110,
            height=24,
            font=("Segoe UI", 10, "bold"),
            command=self.copiar_dados_tabela
        )
        self.btn_copiar_tab.pack(side="right", padx=2)

        self.btn_limpar_tab = ctk.CTkButton(
            f_topo_tab,
            text="🗑️ Limpar",
            width=80,
            height=24,
            font=("Segoe UI", 10),
            command=self.limpar_tabela
        )
        self.btn_limpar_tab.pack(side="right", padx=2)

        # Estilo Treeview
        colunas = ("equipe", "ordem", "item", "material", "descricao", "qtd", "un", "centro", "deposito")
        self.tree = ttk.Treeview(self.frame_tabela, columns=colunas, show="headings", selectmode="browse")
        
        self.tree.heading("equipe", text="Equipe")
        self.tree.heading("ordem", text="Ordem (IW39)")
        self.tree.heading("item", text="Item")
        self.tree.heading("material", text="Material")
        self.tree.heading("descricao", text="Descrição do Material")
        self.tree.heading("qtd", text="Qtd")
        self.tree.heading("un", text="UN")
        self.tree.heading("centro", text="Centro (EMPR)")
        self.tree.heading("deposito", text="Depósito")

        self.tree.column("equipe", width=70, anchor="center")
        self.tree.column("ordem", width=95, anchor="center")
        self.tree.column("item", width=45, anchor="center")
        self.tree.column("material", width=95, anchor="center")
        self.tree.column("descricao", width=220, anchor="w")
        self.tree.column("qtd", width=55, anchor="center")
        self.tree.column("un", width=40, anchor="center")
        self.tree.column("centro", width=90, anchor="center")
        self.tree.column("deposito", width=75, anchor="center")

        scroll_y = ttk.Scrollbar(self.frame_tabela, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(self.frame_tabela, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=1, column=0, sticky="nsew", padx=(6, 0), pady=(0, 6))
        scroll_y.grid(row=1, column=1, sticky="ns", padx=(0, 6), pady=(0, 6))
        scroll_x.grid(row=2, column=0, sticky="ew", padx=(6, 0), pady=(0, 6))

        # 3.2 TERMINAL DE LOG (DIREITA)
        self.frame_log = ctk.CTkFrame(self.frame_conteudo, corner_radius=6)
        self.frame_log.grid(row=0, column=1, padx=(4, 8), pady=8, sticky="nsew")
        self.frame_log.grid_columnconfigure(0, weight=1)
        self.frame_log.grid_rowconfigure(1, weight=1)

        f_topo_log = ctk.CTkFrame(self.frame_log, fg_color="transparent")
        f_topo_log.grid(row=0, column=0, padx=8, pady=(6, 4), sticky="ew")

        lbl_log_title = ctk.CTkLabel(f_topo_log, text="Terminal de Execução SAP", font=("Segoe UI", 12, "bold"))
        lbl_log_title.pack(side="left")

        btn_limpar_log = ctk.CTkButton(
            f_topo_log,
            text="Limpar Log",
            width=90,
            height=24,
            font=("Segoe UI", 10),
            command=self.limpar_log
        )
        btn_limpar_log.pack(side="right")

        self.txt_log = ctk.CTkTextbox(
            self.frame_log,
            font=("Consolas", 11),
            wrap="word",
            state="disabled"
        )
        self.txt_log.grid(row=1, column=0, padx=6, pady=(0, 6), sticky="nsew")

    def _selecionar_equipe(self, equipe):
        self.combo_equipe.set(equipe)
        self._ao_mudar_combo(equipe)

    def _ao_mudar_combo(self, escolha):
        c = self._cores(self.obter_tema_atual())
        for eq, btn in self.btn_equipes.items():
            if eq == escolha:
                btn.configure(fg_color=c["accent"], text_color="white")
            else:
                btn.configure(fg_color=c["input"], text_color=c["texto"])

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for frame in [self.frame_topo, self.frame_inputs, self.frame_conteudo, self.frame_tabela, self.frame_log]:
            frame.configure(fg_color=c.get("container", "#172a45"))

        self.txt_log.configure(fg_color=c["input"], text_color=c["texto"])
        self.entry_ordem.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["accent"])
        self.combo_equipe.configure(fg_color=c["input"], text_color=c["texto"])

        self.btn_buscar_ordem.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_consultar_migo.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_varrer_faixa.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_atalho_todas.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_abrir_sap.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_copiar_tab.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_limpar_tab.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.combo_de.configure(fg_color=c["input"], text_color=c["texto"])
        self.combo_ate.configure(fg_color=c["input"], text_color=c["texto"])

        self._ao_mudar_combo(self.combo_equipe.get())

        # Configuração das cores da Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background=c["input"],
            foreground=c["texto"],
            fieldbackground=c["input"],
            rowheight=22,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading",
            background=c["container"],
            foreground=c["texto"],
            relief="flat",
            font=("Segoe UI", 10, "bold")
        )
        style.map("Treeview", background=[("selected", c["accent"])], foreground=[("selected", "white")])

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

    def limpar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.lbl_qtd_itens.configure(text="Itens Encontrados na MIGO: 0")

    def copiar_dados_tabela(self):
        linhas = []
        linhas.append("Equipe\tOrdem\tItem\tMaterial\tDescrição\tQtd\tUN\tCentro\tDepósito")
        for item in self.tree.get_children():
            valores = self.tree.item(item, "values")
            linhas.append("\t".join(str(v) for v in valores))
        
        if len(linhas) <= 1:
            messagebox.showinfo("Aviso", "A tabela está vazia.")
            return

        texto = "\n".join(linhas)
        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(texto)
        show_toast(self.parent_frame, "Dados copiados para a área de transferência!", duration=2000, type="success")
        play_sound("done")

    # =========================================================================
    # ROTINAS DE EXECUÇÃO SAP (ASSÍNCRONAS EM THREAD)
    # =========================================================================

    def _travar_botoes(self, travado=True):
        state = "disabled" if travado else "normal"
        self.btn_buscar_ordem.configure(state=state)
        self.btn_consultar_migo.configure(state=state)
        self.btn_completo.configure(state=state)
        self.btn_varrer_faixa.configure(state=state)
        self.btn_atalho_todas.configure(state=state)
        self.btn_abrir_sap.configure(state=state)

    def iniciar_busca_ordem_iw39(self):
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe uma consulta SAP em execução.")
            return

        equipe = self.combo_equipe.get().strip().upper()
        if not equipe:
            messagebox.showwarning("Aviso", "Selecione uma equipe clandestina.")
            return

        if not self._verificar_sap_conectado():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto com uma sessão ativa.")
            return

        self.executando = True
        self._travar_botoes(True)
        self._status("Consultando Ordem na IW39...")
        self._log(f"\n🔎 Iniciando busca de Ordem na IW39 para: {equipe}...")

        threading.Thread(target=self._worker_busca_ordem, args=(equipe,), daemon=True).start()

    def _worker_busca_ordem(self, equipe):
        set_sap_ocupado(True)
        try:
            ordem = self._sap_buscar_ordem_iw39(equipe)
            if ordem:
                self.parent_frame.after(0, lambda: self.entry_ordem.delete(0, tk.END))
                self.parent_frame.after(0, lambda: self.entry_ordem.insert(0, ordem))
                self._log_threadsafe(f"🎯 Ordem ativa encontrada para {equipe}: {ordem}")
                self._status_threadsafe(f"Ordem encontrada: {ordem}")
                play_sound("done")
            else:
                self._log_threadsafe(f"⚠️ Nenhuma ordem ativa encontrada para {equipe}.")
                self._status_threadsafe("Nenhuma ordem encontrada.")
        except Exception as e:
            self._log_threadsafe(f"❌ Erro na consulta IW39: {e}")
            log_error(f"Erro IW39: {e}")
        finally:
            set_sap_ocupado(False)
            self.executando = False
            self.parent_frame.after(0, lambda: self._travar_botoes(False))

    def iniciar_consulta_migo(self):
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe uma consulta SAP em execução.")
            return

        ordem = self.entry_ordem.get().strip()
        equipe = self.combo_equipe.get().strip().upper()
        if not ordem:
            messagebox.showwarning("Aviso", "Informe o número da Ordem (ou busque na IW39 primeiro).")
            return

        if not self._verificar_sap_conectado():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto com uma sessão ativa.")
            return

        self.executando = True
        self._travar_botoes(True)
        self._status("Consultando materiais na MIGO...")
        self._log(f"\n📦 Consultando materiais na MIGO para Ordem {ordem} (Equipe: {equipe})...")

        threading.Thread(target=self._worker_consulta_migo, args=(equipe, ordem), daemon=True).start()

    def _worker_consulta_migo(self, equipe, ordem):
        set_sap_ocupado(True)
        try:
            itens = self._sap_consultar_itens_migo(ordem)
            self.parent_frame.after(0, lambda: self._inserir_itens_tabela(equipe, ordem, itens))
            self._log_threadsafe(f"✅ Consulta concluída: {len(itens)} itens encontrados para a ordem {ordem}.")
            self._status_threadsafe(f"MIGO concluída ({len(itens)} itens)")
            play_sound("done")
        except Exception as e:
            self._log_threadsafe(f"❌ Erro na consulta MIGO: {e}")
            log_error(f"Erro MIGO: {e}")
        finally:
            set_sap_ocupado(False)
            self.executando = False
            self.parent_frame.after(0, lambda: self._travar_botoes(False))

    def iniciar_fluxo_completo(self):
        """Busca a ordem na IW39 e em seguida já consulta na MIGO automaticamente."""
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe uma consulta SAP em execução.")
            return

        equipe = self.combo_equipe.get().strip().upper()
        if not equipe:
            messagebox.showwarning("Aviso", "Selecione uma equipe clandestina.")
            return

        if not self._verificar_sap_conectado():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto com uma sessão ativa.")
            return

        self.executando = True
        self._travar_botoes(True)
        self._status(f"Executando IW39 ➡️ MIGO para {equipe}...")
        self._log(f"\n⚡ Iniciando Fluxo Completo para: {equipe}...")

        def _worker():
            set_sap_ocupado(True)
            try:
                self._log_threadsafe(f"1️⃣ Buscando ordem ativa na IW39 para {equipe}...")
                ordem = self._sap_buscar_ordem_iw39(equipe)
                if not ordem:
                    self._log_threadsafe(f"⚠️ Nenhuma ordem encontrada para {equipe}. Interrompendo.")
                    return

                self.parent_frame.after(0, lambda: self.entry_ordem.delete(0, tk.END))
                self.parent_frame.after(0, lambda: self.entry_ordem.insert(0, ordem))
                self._log_threadsafe(f"🎯 Ordem localizada: {ordem}. Indo para a MIGO...")

                time.sleep(0.5)
                self._log_threadsafe(f"2️⃣ Consultando itens na MIGO para Ordem {ordem}...")
                itens = self._sap_consultar_itens_migo(ordem)
                self.parent_frame.after(0, lambda: self._inserir_itens_tabela(equipe, ordem, itens))
                self._log_threadsafe(f"✅ Sucesso: {len(itens)} itens encontrados na MIGO para {equipe} (Ordem {ordem}).")
                play_sound("done")
            except Exception as e:
                self._log_threadsafe(f"❌ Erro no fluxo completo: {e}")
                log_error(f"Erro Fluxo Completo: {e}")
            finally:
                set_sap_ocupado(False)
                self.executando = False
                self.parent_frame.after(0, lambda: self._travar_botoes(False))

        threading.Thread(target=_worker, daemon=True).start()

    def _selecionar_todas_faixa(self):
        self.combo_de.set("RCL001")
        self.combo_ate.set("RCL008")
        self.iniciar_varredura_faixa()

    def obter_faixa_equipes(self):
        de = self.combo_de.get().strip().upper()
        ate = self.combo_ate.get().strip().upper()
        try:
            i1 = EQUIPES_CLANDESTINAS.index(de)
            i2 = EQUIPES_CLANDESTINAS.index(ate)
            if i1 <= i2:
                return EQUIPES_CLANDESTINAS[i1 : i2 + 1]
            else:
                return EQUIPES_CLANDESTINAS[i2 : i1 + 1]
        except ValueError:
            return [de]

    def iniciar_varredura_faixa(self):
        """Varre a faixa selecionada de equipes (De X até X) consultando IW39 e MIGO."""
        if self.executando:
            messagebox.showwarning("Aviso", "Já existe uma consulta SAP em execução.")
            return

        if not self._verificar_sap_conectado():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto com uma sessão ativa.")
            return

        equipes_alvo = self.obter_faixa_equipes()
        texto_faixa = f"{equipes_alvo[0]} a {equipes_alvo[-1]}" if len(equipes_alvo) > 1 else equipes_alvo[0]

        if not messagebox.askyesno("Consulta em Faixa", f"Deseja consultar a faixa ({texto_faixa}) com {len(equipes_alvo)} equipe(s) na IW39 e MIGO?"):
            return

        self.executando = True
        self._travar_botoes(True)
        self.limpar_tabela()
        self._log("\n🌐 ========================================================")
        self._log(f"🌐 INICIANDO CONSULTA EM FAIXA: {texto_faixa} ({len(equipes_alvo)} equipes)")
        self._log("🌐 ========================================================")

        def _worker():
            set_sap_ocupado(True)
            try:
                for equipe in equipes_alvo:
                    self._log_threadsafe(f"\n🔎 [{equipe}] Consultando IW39...")
                    self._status_threadsafe(f"Faixa: {equipe}...")
                    try:
                        ordem = self._sap_buscar_ordem_iw39(equipe)
                        if not ordem:
                            self._log_threadsafe(f"  ↳ [{equipe}] Nenhuma ordem ativa.")
                            continue

                        self._log_threadsafe(f"  ↳ [{equipe}] Ordem: {ordem}. Consultando MIGO...")
                        itens = self._sap_consultar_itens_migo(ordem)
                        self.parent_frame.after(0, lambda e=equipe, o=ordem, it=itens: self._inserir_itens_tabela(e, o, it))
                        self._log_threadsafe(f"  ↳ [{equipe}] {len(itens)} itens encontrados na MIGO.")
                    except Exception as err:
                        self._log_threadsafe(f"  ↳ [{equipe}] Falha: {err}")
                    time.sleep(0.5)

                self._log_threadsafe(f"\n🏁 Consulta da faixa ({texto_faixa}) concluída!")
                self._status_threadsafe("Consulta concluída!")
                play_sound("done")
            finally:
                set_sap_ocupado(False)
                self.executando = False
                self.parent_frame.after(0, lambda: self._travar_botoes(False))

        threading.Thread(target=_worker, daemon=True).start()

    def abrir_migo_sap_gui(self):
        """Navega para a MIGO no SAP GUI, preenche a ordem e coloca foco na janela."""
        ordem = self.entry_ordem.get().strip()
        if not ordem:
            messagebox.showwarning("Aviso", "Informe o número da Ordem para abrir no SAP.")
            return

        if not self._verificar_sap_conectado():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto com uma sessão ativa.")
            return

        self._log(f"\n🖥️ Abrindo Ordem {ordem} na MIGO no SAP GUI...")
        vbs = f'''
On Error Resume Next
Dim SapGuiAuto, application, connection, session
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
    WScript.Echo "ERRO:SAP_DESCONECTADO"
    WScript.Quit 1
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or application Is Nothing Then
    WScript.Echo "ERRO:SCRIPTING_DESABILITADO"
    WScript.Quit 1
End If

If application.Children.Count = 0 Then
    WScript.Echo "ERRO:SEM_CONEXAO"
    WScript.Quit 1
End If

Dim c, s, connCand, sessCand
Set connection = Nothing
Set session = Nothing
For c = 0 To application.Children.Count - 1
    Set connCand = application.Children(c)
    If Not connCand Is Nothing Then
        If connCand.Children.Count > 0 Then
            For s = 0 To connCand.Children.Count - 1
                Set sessCand = connCand.Children(s)
                If Not sessCand Is Nothing Then
                    Set connection = connCand
                    Set session = sessCand
                    Exit For
                End If
            Next
        End If
    End If
    If Not session Is Nothing Then Exit For
Next

If session Is Nothing Then
    WScript.Echo "ERRO:SEM_SESSAO"
    WScript.Quit 1
End If

' Fechar eventuais popups residuais
Dim p
For p = 1 To 3
    If session.Children.Count > 1 Then
        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        If Err.Number <> 0 Then
            Err.Clear
            session.findById("wnd[1]").sendVKey 0
        End If
        WScript.Sleep 150
    End If
Next
On Error GoTo 0

Dim ordem
ordem = "{ordem}"

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 400

Dim subs, sub_idx
subs = Array("0003", "0002")
For Each sub_idx In subs
    On Error Resume Next
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_idx & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A07"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_idx & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R08"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_idx & "/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = ordem
    If Err.Number = 0 Then
        session.findById("wnd[0]").sendVKey 0
        Exit For
    End If
Next

session.findById("wnd[0]").maximize
WScript.Echo "OK"
'''
        rc, out, err = self._executar_script_vbs(vbs, timeout=30)
        if rc == 0 and "OK" in out:
            self._log(f"🖥️ Tela da MIGO aberta com a ordem {ordem} no SAP.")
            show_toast(self.parent_frame, f"Ordem {ordem} carregada na MIGO no SAP!", duration=2000, type="success")
        else:
            msg_erro = err or out
            if "ERRO:SAP_DESCONECTADO" in msg_erro:
                msg_erro = "O SAP GUI está aberto, mas não possui conexão ativa com o sistema. Conecte no SAP Logon primeiro."
            elif "ERRO:SEM_SESSAO" in msg_erro:
                msg_erro = "Nenhuma sessão ativa aberta no SAP. Conecte ao ambiente no SAP Logon."
            elif "ERRO:SCRIPTING_DESABILITADO" in msg_erro:
                msg_erro = "Scripting do SAP GUI desabilitado nas Opções do SAP GUI."
            self._log(f"❌ Erro ao abrir MIGO no SAP: {msg_erro}")
            messagebox.showerror("Aviso SAP", f"Não foi possível abrir a MIGO no SAP GUI:\n{msg_erro}")

    def _inserir_itens_tabela(self, equipe, ordem, itens):
        for item in itens:
            self.tree.insert(
                "",
                "end",
                values=(
                    equipe,
                    ordem,
                    item.get("item", ""),
                    item.get("material", ""),
                    item.get("descricao", ""),
                    item.get("qtd", ""),
                    item.get("unidade", ""),
                    item.get("centro", ""),
                    item.get("deposito", "")
                )
            )
        total = len(self.tree.get_children())
        self.lbl_qtd_itens.configure(text=f"Itens Encontrados na MIGO: {total}")

    # =========================================================================
    # FUNÇÕES DE EXECUÇÃO VIA CSCRIPT / VBSCRIPT NATIVO (100% COMPATÍVEL COM SAP)
    # =========================================================================

    def _executar_script_vbs(self, conteudo, timeout=120):
        import subprocess, tempfile
        fd, caminho = tempfile.mkstemp(prefix="nexus_migo_", suffix=".vbs", text=True)
        try:
            with os.fdopen(fd, "w", encoding="cp1252", errors="ignore") as f:
                f.write(conteudo)
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            
            cscripts = []
            cscript_32 = r"C:\Windows\SysWOW64\cscript.exe"
            if os.path.exists(cscript_32):
                cscripts.append(cscript_32)
            cscripts.append("cscript.exe")

            last_rc, last_out, last_err = -1, "", ""
            for cs in cscripts:
                try:
                    res = subprocess.run(
                        [cs, "//nologo", caminho],
                        capture_output=True, text=True, timeout=timeout, creationflags=creationflags
                    )
                    last_rc, last_out, last_err = res.returncode, res.stdout or "", res.stderr or ""
                    # Se não deu erro de Syntax/GetObject, aceita a execução
                    if "800401E4" not in last_out and "Sintaxe inválida" not in last_out:
                        return last_rc, last_out, last_err
                except Exception as e:
                    last_err = str(e)
                    continue

            return last_rc, last_out, last_err
        except Exception as e:
            return -1, "", str(e)
        finally:
            try:
                if os.path.exists(caminho):
                    os.remove(caminho)
            except OSError:
                pass

    def _verificar_sap_conectado(self):
        """Verifica se o SAP GUI está aberto e acessível."""
        return verificar_sap_aberto()

    def _sap_buscar_ordem_iw39(self, equipe):
        vbs = f'''
On Error Resume Next
Dim SapGuiAuto, application, connection, session
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
    WScript.Echo "ERRO:SAP_DESCONECTADO"
    WScript.Quit 1
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or application Is Nothing Then
    WScript.Echo "ERRO:SCRIPTING_DESABILITADO"
    WScript.Quit 1
End If

If application.Children.Count = 0 Then
    WScript.Echo "ERRO:SEM_CONEXAO"
    WScript.Quit 1
End If

Dim c, s, connCand, sessCand
Set connection = Nothing
Set session = Nothing
For c = 0 To application.Children.Count - 1
    Set connCand = application.Children(c)
    If Not connCand Is Nothing Then
        If connCand.Children.Count > 0 Then
            For s = 0 To connCand.Children.Count - 1
                Set sessCand = connCand.Children(s)
                If Not sessCand Is Nothing Then
                    Set connection = connCand
                    Set session = sessCand
                    Exit For
                End If
            Next
        End If
    End If
    If Not session Is Nothing Then Exit For
Next

If session Is Nothing Then
    WScript.Echo "ERRO:SEM_SESSAO"
    WScript.Quit 1
End If

' Fechar eventuais popups residuais
Dim p
For p = 1 To 3
    If session.Children.Count > 1 Then
        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        If Err.Number <> 0 Then
            Err.Clear
            session.findById("wnd[1]").sendVKey 0
        End If
        WScript.Sleep 150
    End If
Next
On Error GoTo 0

Dim recurso
recurso = "{equipe}"

' 1. Navega para IW39
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 400

' 2. Preenche Parceiro Z2 e codigo da equipe
session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = recurso
session.findById("wnd[0]/usr/ctxtDY_PARNR").setFocus
session.findById("wnd[0]/usr/ctxtDY_PARNR").caretPosition = Len(recurso)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 200

' 3. Filtro de Status: Apenas Ativas (Pendente e Em Processamento)
session.findById("wnd[0]/usr/chkDY_OFN").selected = True
session.findById("wnd[0]/usr/chkDY_IAR").selected = True
session.findById("wnd[0]/usr/chkDY_MAB").selected = False
session.findById("wnd[0]/usr/chkDY_ABG").selected = False
session.findById("wnd[0]/usr/chkDY_HIS").selected = False

' 4. Executa F8
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 700

' Fecha popup informativo se surgir
If session.Children.Count > 1 Then
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 200
End If

Dim ordemEncontrada
ordemEncontrada = ""

' Caso A: Grade ALV
Dim grid
Set grid = Nothing
On Error Resume Next
Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
On Error GoTo 0

If Not grid Is Nothing Then
    On Error Resume Next
    grid.setCurrentCell -1, "GSTRP"
    grid.selectColumn "GSTRP"
    If Err.Number <> 0 Then
        Err.Clear
        grid.setCurrentCell -1, "GLTRP"
        grid.selectColumn "GLTRP"
    End If
    session.findById("wnd[0]/tbar[1]/btn[40]").press
    WScript.Sleep 300
    Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    ordemEncontrada = Trim(grid.GetCellValue(0, "AUFNR"))
    If ordemEncontrada = "" Then
        ordemEncontrada = Trim(grid.GetCellValue(0, "ORDERID"))
    End If
    On Error GoTo 0
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    WScript.Sleep 200
    session.findById("wnd[0]/tbar[0]/btn[15]").press
End If

' Caso B: Detalhe individual
If ordemEncontrada = "" Then
    On Error Resume Next
    Dim txtOrd
    Set txtOrd = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR")
    If txtOrd Is Nothing Then
        Set txtOrd = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1101/txtCAUFVD-AUFNR")
    End If
    If txtOrd Is Nothing Then
        Set txtOrd = session.findById("wnd[0]/usr/txtCAUFVD-AUFNR")
    End If
    If Not txtOrd Is Nothing Then
        ordemEncontrada = Trim(txtOrd.text)
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        WScript.Sleep 200
        session.findById("wnd[0]/tbar[0]/btn[15]").press
    End If
    On Error GoTo 0
End If

If ordemEncontrada <> "" Then
    WScript.Echo "ORDEM_ENCONTRADA:" & ordemEncontrada
Else
    WScript.Echo "NENHUMA_ORDEM"
End If
'''
        rc, out, err = self._executar_script_vbs(vbs, timeout=60)
        for linha in out.splitlines():
            linha = linha.strip()
            if linha.startswith("ORDEM_ENCONTRADA:"):
                return linha.split(":", 1)[1].strip()
            elif "ERRO:SAP_DESCONECTADO" in linha or "ERRO:SAP_FECHADO" in linha:
                raise RuntimeError("O SAP GUI está aberto, mas não possui conexão ativa com o sistema. Conecte ao ambiente (ex: Produção) antes de consultar.")
            elif "ERRO:SEM_SESSAO" in linha or "ERRO:SEM_CONEXAO" in linha:
                raise RuntimeError("Nenhuma sessão ativa aberta no SAP. Conecte ao ambiente no SAP Logon.")
            elif "ERRO:SCRIPTING_DESABILITADO" in linha or "ERRO:SCRIPTING_ENGINE" in linha:
                raise RuntimeError("Scripting do SAP GUI está desabilitado. Habilite nas Opções do SAP GUI (Acessibilidade e Scripting).")
            elif "ERRO:" in linha:
                raise RuntimeError(linha)
        return ""

    def _sap_consultar_itens_migo(self, ordem):
        vbs = f'''
On Error Resume Next
Dim SapGuiAuto, application, connection, session
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
    WScript.Echo "ERRO:SAP_DESCONECTADO"
    WScript.Quit 1
End If

Set application = SapGuiAuto.GetScriptingEngine
If Err.Number <> 0 Or application Is Nothing Then
    WScript.Echo "ERRO:SCRIPTING_DESABILITADO"
    WScript.Quit 1
End If

If application.Children.Count = 0 Then
    WScript.Echo "ERRO:SEM_CONEXAO"
    WScript.Quit 1
End If

Dim c, s, connCand, sessCand
Set connection = Nothing
Set session = Nothing
For c = 0 To application.Children.Count - 1
    Set connCand = application.Children(c)
    If Not connCand Is Nothing Then
        If connCand.Children.Count > 0 Then
            For s = 0 To connCand.Children.Count - 1
                Set sessCand = connCand.Children(s)
                If Not sessCand Is Nothing Then
                    Set connection = connCand
                    Set session = sessCand
                    Exit For
                End If
            Next
        End If
    End If
    If Not session Is Nothing Then Exit For
Next

If session Is Nothing Then
    WScript.Echo "ERRO:SEM_SESSAO"
    WScript.Quit 1
End If

' Fechar eventuais popups residuais
Dim p
For p = 1 To 3
    If session.Children.Count > 1 Then
        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        If Err.Number <> 0 Then
            Err.Clear
            session.findById("wnd[1]").sendVKey 0
        End If
        WScript.Sleep 150
    End If
Next
On Error GoTo 0

Dim ordem
ordem = "{ordem}"

' Navega para MIGO
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

' Preenche A07 (Saida) e R08 (Ordem)
Dim subs, sub_item
subs = Array("0003", "0002")
For Each sub_item In subs
    Err.Clear
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A07"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R08"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = ordem
    If Err.Number = 0 Then
        session.findById("wnd[0]").sendVKey 0
        Exit For
    End If
Next

WScript.Sleep 800

' Fecha eventuais popups de advertencia
For p = 1 To 3
    If session.Children.Count > 1 Then
        session.findById("wnd[1]").sendVKey 0
        WScript.Sleep 200
    End If
Next

' Localiza tabela de itens
Dim tbl
Set tbl = Nothing
For Each sub_item In subs
    Err.Clear
    Set tbl = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM")
    If Err.Number = 0 And Not tbl Is Nothing Then Exit For
Next

If Not tbl Is Nothing Then
    Dim totalLinhas, r, maxRows
    totalLinhas = tbl.RowCount
    If totalLinhas > 50 Then maxRows = 50 Else maxRows = totalLinhas
    If maxRows = 0 Then maxRows = tbl.VisibleRowCount

    For r = 0 To maxRows - 1
        On Error Resume Next
        If r >= tbl.VisibleRowCount Then
            tbl.VerticalScrollbar.Position = r
        End If

        Dim mat, desc, qtd, un, centro, dep
        mat = ""
        desc = ""
        qtd = ""
        un = ""
        centro = ""
        dep = ""

        mat = Trim(tbl.GetCellValue(r, "MATNR"))
        If mat = "" Then
            mat = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MATNR[1," & r & "]").text)
        End If
        If mat = "" Then
            mat = Trim(session.findById(tbl.Id & "/txtGOITEM-MATNR[1," & r & "]").text)
        End If

        desc = Trim(tbl.GetCellValue(r, "MAKTX"))
        qtd = Trim(tbl.GetCellValue(r, "ERFMG"))
        un = Trim(tbl.GetCellValue(r, "ERFME"))
        centro = Trim(tbl.GetCellValue(r, "NAME1"))
        If centro = "" Then centro = Trim(tbl.GetCellValue(r, "WERKS"))
        If centro = "" Then centro = "EMPR"

        dep = Trim(tbl.GetCellValue(r, "LGORT"))
        If dep = "" Then dep = Trim(tbl.GetCellValue(r, "LGOBE"))

        If mat <> "" Or desc <> "" Or qtd <> "" Then
            WScript.Echo "ITEM:" & (r + 1) & "|" & mat & "|" & desc & "|" & qtd & "|" & un & "|" & centro & "|" & dep
        End If
    Next
Else
    ' Caso de item unico detalhado
    On Error Resume Next
    Dim txtMat
    Set txtMat = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX")
    If Not txtMat Is Nothing Then
        WScript.Echo "ITEM:1|Item Ordem|" & Trim(txtMat.text) & "|1|UN|EMPR|CL03"
    End If
End If

WScript.Echo "FIM_MIGO"
'''
        rc, out, err = self._executar_script_vbs(vbs, timeout=60)
        itens = []
        for linha in out.splitlines():
            linha = linha.strip()
            if linha.startswith("ITEM:"):
                partes = linha.split(":", 1)[1].split("|")
                if len(partes) >= 7:
                    itens.append({
                        "item": partes[0],
                        "material": partes[1],
                        "descricao": partes[2],
                        "qtd": partes[3],
                        "unidade": partes[4],
                        "centro": partes[5],
                        "deposito": partes[6]
                    })
            elif "ERRO:SAP_DESCONECTADO" in linha or "ERRO:SAP_FECHADO" in linha:
                raise RuntimeError("O SAP GUI está aberto, mas não possui conexão ativa com o sistema. Conecte ao ambiente (ex: Produção) antes de consultar.")
            elif "ERRO:SEM_SESSAO" in linha or "ERRO:SEM_CONEXAO" in linha:
                raise RuntimeError("Nenhuma sessão ativa aberta no SAP. Conecte ao ambiente no SAP Logon.")
            elif "ERRO:SCRIPTING_DESABILITADO" in linha or "ERRO:SCRIPTING_ENGINE" in linha:
                raise RuntimeError("Scripting do SAP GUI está desabilitado. Habilite nas Opções do SAP GUI (Acessibilidade e Scripting).")
            elif "ERRO:" in linha:
                raise RuntimeError(linha)
        return itens


def construir_aba_consultar_migo(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaConsultarMigo(parent_frame, label_status_global, paletas, obter_tema_atual)
