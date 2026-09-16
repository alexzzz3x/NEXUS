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

        if not verificar_sap_aberto():
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

        if not verificar_sap_aberto():
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

        if not verificar_sap_aberto():
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

        if not verificar_sap_aberto():
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

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto com uma sessão ativa.")
            return

        try:
            sessao = obter_sessao_ativa_sap()
            if not sessao:
                raise RuntimeError("Não foi possível obter a sessão ativa do SAP.")

            self._log(f"\n🖥️ Abrindo Ordem {ordem} na MIGO no SAP GUI...")
            sessao.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIGO"
            sessao.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)

            # Define A07 (Saída) e R08 (Ordem)
            try:
                sessao.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
                sessao.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R08"
                sessao.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = ordem
                sessao.findById("wnd[0]").sendVKey(0)
            except Exception:
                sessao.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
                sessao.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R08"
                sessao.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = ordem
                sessao.findById("wnd[0]").sendVKey(0)

            # Traz a janela do SAP para frente
            try:
                def enum_handler(hwnd, extra):
                    if "SAP" in win32gui.GetWindowText(hwnd):
                        win32gui.SetForegroundWindow(hwnd)
                win32gui.EnumWindows(enum_handler, None)
            except Exception:
                pass

            self._log(f"🖥️ Tela da MIGO aberta com a ordem {ordem} no SAP.")
            show_toast(self.parent_frame, f"Ordem {ordem} carregada na MIGO no SAP!", duration=2000, type="success")
        except Exception as e:
            self._log(f"❌ Erro ao abrir MIGO no SAP: {e}")
            messagebox.showerror("Erro SAP", str(e))

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
    # FUNÇÕES DE AUTOMAÇÃO SAP COM
    # =========================================================================

    def _sap_buscar_ordem_iw39(self, equipe):
        pythoncom.CoInitialize()
        sessao = obter_sessao_ativa_sap()
        if not sessao:
            raise RuntimeError("Não foi possível conectar ao SAP GUI.")

        # 1. Navega para IW39
        sessao.findById("wnd[0]/tbar[0]/btn[12]").press()
        sessao.findById("wnd[0]/tbar[0]/okcd").Text = "/nIW39"
        sessao.findById("wnd[0]").sendVKey(0)
        time.sleep(0.3)

        # 2. Configura Parceiro Z2 e código da equipe
        try:
            sessao.findById("wnd[0]/usr/cmbDY_PARVW").Key = "Z2"
            campo_eq = sessao.findById("wnd[0]/usr/ctxtDY_PARNR")
            campo_eq.Text = equipe
            campo_eq.CaretPosition = len(equipe)
            sessao.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            raise RuntimeError(f"Falha ao preencher parâmetros do parceiro na IW39: {e}")

        # 3. Filtro de Status: Apenas Ativas (Pendente e Em Processamento)
        status_chk = [
            ("wnd[0]/usr/chkDY_OFN", True),   # Pendente
            ("wnd[0]/usr/chkDY_IAR", True),   # Em processamento
            ("wnd[0]/usr/chkDY_MAB", False),  # Encerrado
            ("wnd[0]/usr/chkDY_ABG", False),  # Encerrado alternativo
            ("wnd[0]/usr/chkDY_HIS", False),  # Histórico
        ]
        for chk_id, valor in status_chk:
            try:
                ctrl = sessao.findById(chk_id, False)
                if ctrl:
                    ctrl.Selected = valor
            except Exception:
                pass

        # 4. Executa F8
        sessao.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(0.8)

        # Fecha popup informativo se surgir
        try:
            if sessao.findById("wnd[1]", False):
                sessao.findById("wnd[1]").sendVKey(0)
                time.sleep(0.3)
        except Exception:
            pass

        ordem_encontrada = ""

        # Caso A: Grade ALV (lista de ordens)
        try:
            grid = sessao.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
            if grid:
                try:
                    # Tenta ordenar decrescente pela data mais recente
                    grid.setCurrentCell(-1, "GSTRP")
                    grid.selectColumn("GSTRP")
                    sessao.findById("wnd[0]/tbar[1]/btn[40]").press()
                    time.sleep(0.3)
                    grid = sessao.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                except Exception:
                    pass

                total_linhas = int(grid.RowCount)
                for r in range(total_linhas):
                    val_ord = str(grid.getCellValue(r, "AUFNR")).strip()
                    val_rec = str(grid.getCellValue(r, "PARNR")).strip().upper()
                    if val_ord:
                        if not val_rec or val_rec == equipe:
                            ordem_encontrada = val_ord
                            break

                # Retorna da grade
                sessao.findById("wnd[0]/tbar[0]/btn[15]").press()
                sessao.findById("wnd[0]/tbar[0]/btn[15]").press()
        except Exception:
            pass

        # Caso B: Abre diretamente no detalhe de uma única ordem
        if not ordem_encontrada:
            try:
                txt_ord = sessao.findById(
                    "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR",
                    False
                )
                if txt_ord:
                    ordem_encontrada = str(txt_ord.Text).strip()
                    sessao.findById("wnd[0]/tbar[0]/btn[15]").press()
                    sessao.findById("wnd[0]/tbar[0]/btn[15]").press()
            except Exception:
                pass

        return ordem_encontrada

    def _sap_consultar_itens_migo(self, ordem):
        """Acessa a MIGO com a Ordem (A07/R08) e varre os itens e materiais."""
        pythoncom.CoInitialize()
        sessao = obter_sessao_ativa_sap()
        if not sessao:
            raise RuntimeError("Não foi possível conectar ao SAP GUI.")

        sessao.findById("wnd[0]/tbar[0]/btn[12]").press()
        sessao.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIGO"
        sessao.findById("wnd[0]").sendVKey(0)
        time.sleep(0.4)

        # Configura A07 (Saída) e R08 (Ordem)
        for sub in ("0003", "0002"):
            try:
                sessao.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
                sessao.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R08"
                sessao.findById(f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = str(ordem)
                sessao.findById("wnd[0]").sendVKey(0)
                break
            except Exception:
                continue

        time.sleep(0.8)

        # Trata popups de advertência da MIGO
        try:
            for _ in range(3):
                if sessao.findById("wnd[1]", False):
                    sessao.findById("wnd[1]").sendVKey(0)
                    time.sleep(0.3)
        except Exception:
            pass

        itens = []

        # Localiza a tabela de itens na MIGO
        tbl = None
        for sub in ("0003", "0002"):
            tbl_cand = sessao.findById(
                f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{sub}/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
                False
            )
            if tbl_cand:
                tbl = tbl_cand
                break

        if tbl:
            total_linhas = int(getattr(tbl, "RowCount", 0))
            vis_linhas = int(getattr(tbl, "VisibleRowCount", 10))
            max_iter = min(total_linhas, 50) if total_linhas > 0 else vis_linhas

            for r in range(max_iter):
                try:
                    # Garante que a linha esteja visível
                    if r >= vis_linhas:
                        try:
                            tbl.VerticalScrollbar.Position = r
                        except Exception:
                            pass

                    mat = ""
                    desc = ""
                    qtd = ""
                    un = ""
                    centro = ""
                    dep = ""

                    for nome_mat in ("txtGOITEM-MATNR", "ctxtGOITEM-MATNR"):
                        try:
                            mat_ctrl = sessao.findById(f"{tbl.Id}/{nome_mat}[1,{r}]", False)
                            if mat_ctrl and str(mat_ctrl.Text).strip():
                                mat = str(mat_ctrl.Text).strip()
                                break
                        except Exception:
                            pass

                    # Se não obteve na coluna 1, tenta varrer colunas comuns
                    if not mat:
                        try:
                            mat = str(tbl.GetCellValue(r, "MATNR")).strip()
                        except Exception:
                            pass

                    try:
                        desc = str(tbl.GetCellValue(r, "MAKTX")).strip()
                    except Exception:
                        pass

                    try:
                        qtd = str(tbl.GetCellValue(r, "ERFMG")).strip()
                    except Exception:
                        pass

                    try:
                        un = str(tbl.GetCellValue(r, "ERFME")).strip()
                    except Exception:
                        pass

                    try:
                        centro = str(tbl.GetCellValue(r, "NAME1")).strip()
                        if not centro:
                            centro = str(tbl.GetCellValue(r, "WERKS")).strip()
                    except Exception:
                        centro = "EMPR"

                    try:
                        dep = str(tbl.GetCellValue(r, "LGORT")).strip()
                        if not dep:
                            dep = str(tbl.GetCellValue(r, "LGOBE")).strip()
                    except Exception:
                        pass

                    if mat or desc or qtd:
                        itens.append({
                            "item": str(r + 1),
                            "material": mat,
                            "descricao": desc,
                            "qtd": qtd,
                            "unidade": un,
                            "centro": centro or "EMPR",
                            "deposito": dep
                        })
                except Exception as e:
                    continue
        else:
            # Caso não tenha tabela em lista (ex: item único detalhado)
            try:
                txt_mat = sessao.findById(
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX",
                    False
                )
                if txt_mat:
                    itens.append({
                        "item": "1",
                        "material": "Item Ordem",
                        "descricao": str(txt_mat.Text).strip(),
                        "qtd": "1",
                        "unidade": "UN",
                        "centro": "EMPR",
                        "deposito": "CL03"
                    })
            except Exception:
                pass

        return itens


def construir_aba_consultar_migo(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaConsultarMigo(parent_frame, label_status_global, paletas, obter_tema_atual)
