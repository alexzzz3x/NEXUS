# -*- coding: utf-8 -*-
"""
Módulo de Atualização de OSME e Vencimento (SAP IW39 -> Planilha de Controle).

Funcionalidades:
1. Lê a lista de equipes cadastradas na aba CADASTRO da planilha CONTROLE_SERIAIS.xlsx.
2. Consulta no SAP IW39 recurso por recurso (DY_PARVW="Z2", DY_PARNR=equipe).
3. Trata dinamicamente os dois cenários do SAP GUI:
   - Cenário 1 (1 ordem): Cabeçalho único (txtCAUFVD-AUFNR e ctxtCAUFVD-GLTRP).
   - Cenário 2 (2+ ordens): Grade ALV (cntlGRID1), ordenando decrescente por GLTRP (btn[40]) e lendo a linha 0.
   - Cenário 3 (0 ordens): Nenhuma ordem ativa no período.
4. Exibe comparativo em tempo real com status e cores dinâmicas.
5. Grava as alterações diretamente nas colunas H (Ordem) e I (Vencimento) da aba CADASTRO com backup automático.
"""

import os
import io
import re
import json
import time
import shutil
import tempfile
import threading
import subprocess
from datetime import datetime, timedelta
from collections import OrderedDict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from openpyxl import load_workbook

from utils_os import obter_caminho_config, obter_diretorio_dados
from modulo_Sap import verificar_sap_aberto
from utils_pendencias_planilha import (
    is_file_writable,
    adicionar_pendencia,
    obter_quantidade_pendencias,
    processar_todas_pendencias,
    ModalVisualizarPendencias
)

CAMINHO_CONFIG = obter_caminho_config("config_saida_osme_caminho.json")

DEFAULT_PALETAS = {
    "Nexus": {
        "bg": "#090d16", "container": "#111827", "texto": "#f8fafc",
        "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436",
        "sucesso": "#10b981", "alerta": "#ef4444"
    },
    "Azul Delray": {
        "bg": "#0a192f", "container": "#172a45", "texto": "#ffffff",
        "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60",
        "sucesso": "#10b981", "alerta": "#ef4444"
    },
    "Conecta Blue": {
        "bg": "#0a192f", "container": "#172a45", "texto": "#ffffff",
        "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60",
        "sucesso": "#10b981", "alerta": "#ef4444"
    }
}


def normalizar_texto(valor):
    if valor is None:
        return ""
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor))
    texto_sem_acento = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", texto_sem_acento.strip().upper())



def is_data_valida(data_str):
    """Valida se uma string é uma data válida no formato DD/MM/YYYY."""
    if not data_str or str(data_str).strip() in ("—", "-", "None", ""):
        return False
    s = str(data_str).strip()
    if any(p in s.lower() for p in ("fim", "base", "data", "ordem", "caufvd", "lbl", "ctxt")):
        return False
    return bool(re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", s))


def formatar_data_para_exibicao(data_str):
    """Converte datas do SAP (ex: '20.11.2026', '20261120') ou Excel para 'DD/MM/YYYY'.
    Retorna '—' se não for uma data válida (rejeita explicitamente rótulos como 'Fim-base').
    """
    if not data_str or str(data_str).strip() in ("—", "-", "None", ""):
        return "—"
    if isinstance(data_str, datetime):
        return data_str.strftime("%d/%m/%Y")
    try:
        from datetime import date
        if isinstance(data_str, date):
            return data_str.strftime("%d/%m/%Y")
    except Exception:
        pass

    data_str = str(data_str).strip()

    # Rejeita explicitamente se contiver palavras de rótulo (ex: 'fim', 'base', 'data', 'ordem')
    if any(p in data_str.lower() for p in ("fim", "base", "data", "ordem", "caufvd", "lbl", "ctxt")):
        return "—"

    # Formato DD.MM.YYYY ou DD/MM/YYYY ou DD-MM-YYYY (com hora opcional no final)
    m = re.match(r"^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})", data_str)
    if m:
        d, m_mes, a = m.groups()
        return f"{int(d):02d}/{int(m_mes):02d}/{a}"

    # Formato YYYY-MM-DD ou YYYY.MM.DD ou YYYY/MM/DD
    m_iso = re.match(r"^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})", data_str)
    if m_iso:
        a, m_mes, d = m_iso.groups()
        return f"{int(d):02d}/{int(m_mes):02d}/{a}"

    # Formato YYYYMMDD
    m2 = re.match(r"^(\d{4})(\d{2})(\d{2})$", data_str)
    if m2:
        a, m_mes, d = m2.groups()
        return f"{d}/{m_mes}/{a}"

    try:
        num = int(float(data_str))
        if 40000 < num < 60000:
            return (datetime(1899, 12, 30) + timedelta(days=num)).strftime("%d/%m/%Y")
    except Exception:
        pass

    return "—"


def converter_para_datetime(data_val):
    """Converte qualquer formato de data válido (string DD/MM/YYYY, DD.MM.YYYY, YYYY-MM-DD, datetime)
    em um objeto datetime(ano, mes, dia) com garantia de dia e mês preservados.
    Retorna None se não for conversível.
    """
    if not data_val:
        return None
    if isinstance(data_val, datetime):
        return datetime(data_val.year, data_val.month, data_val.day)
    try:
        from datetime import date
        if isinstance(data_val, date):
            return datetime(data_val.year, data_val.month, data_val.day)
    except Exception:
        pass

    s = str(data_val).strip()
    if any(p in s.lower() for p in ("fim", "base", "data", "ordem", "caufvd", "lbl", "ctxt", "—", "-")):
        return None

    # Formato DD.MM.YYYY ou DD/MM/YYYY ou DD-MM-YYYY
    m = re.match(r"^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})", s)
    if m:
        d, m_mes, a = m.groups()
        try:
            return datetime(int(a), int(m_mes), int(d))
        except ValueError:
            return None

    # Formato YYYY-MM-DD ou YYYY.MM.DD ou YYYY/MM/DD
    m_iso = re.match(r"^(\d{4})[\.\/-](\d{1,2})[\.\/-](\d{1,2})", s)
    if m_iso:
        a, m_mes, d = m_iso.groups()
        try:
            return datetime(int(a), int(m_mes), int(d))
        except ValueError:
            return None

    # Formato YYYYMMDD
    m2 = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m2:
        a, m_mes, d = m2.groups()
        try:
            return datetime(int(a), int(m_mes), int(d))
        except ValueError:
            return None

    try:
        num = int(float(s))
        if 40000 < num < 60000:
            dt = datetime(1899, 12, 30) + timedelta(days=num)
            return datetime(dt.year, dt.month, dt.day)
    except Exception:
        pass

    return None


def limpar_processos_excel_fantasmas():
    """Encerra processos EXCEL.EXE sem janela visível (fantasmas/órfãos) que bloqueiam planilhas."""
    try:
        subprocess.run(
            [
                "powershell", "-NoProfile", "-Command",
                "Get-Process -Name excel -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -eq '' } | Stop-Process -Force -ErrorAction SilentlyContinue"
            ],
            capture_output=True, timeout=5
        )
    except Exception:
        pass


class AbaAtualizarOSME:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.lbl_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual

        self.caminho_planilha = self._carregar_caminho_salvo()
        self.dados_equipes = []  # Lista de dicts lidos da aba CADASTRO
        self.executando = False
        self.cancelar_execucao = False
        self.carregando_planilha = False
        self.col_equipe_idx = 4   # Coluna E (0-indexed: 4)
        self.col_dep_idx = 5      # Coluna F (0-indexed: 5)
        self.col_seg_idx = 6      # Coluna G (0-indexed: 6)
        self.col_ordem_idx = 7    # Coluna H (0-indexed: 7)
        self.col_venc_idx = 8     # Coluna I (0-indexed: 8)

        self._construir_ui()
        self._iniciar_timer_pendencias()

        self.precisa_carregar_planilha = bool(self.caminho_planilha and os.path.exists(self.caminho_planilha))

        def _ao_exibir_aba(event=None):
            if self.precisa_carregar_planilha:
                self.precisa_carregar_planilha = False
                self._iniciar_leitura_planilha_thread()

        self.parent_frame.bind("<Visibility>", _ao_exibir_aba, add="+")

    def _carregar_caminho_salvo(self):
        if os.path.exists(CAMINHO_CONFIG):
            try:
                with open(CAMINHO_CONFIG, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                    return dados.get("caminho_planilha", "")
            except Exception:
                pass
        return ""

    def _salvar_caminho(self, caminho):
        self.caminho_planilha = caminho
        try:
            with open(CAMINHO_CONFIG, "w", encoding="utf-8") as f:
                json.dump({"caminho_planilha": caminho}, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def _status(self, texto):
        if self.lbl_status_global and hasattr(self.lbl_status_global, "configure"):
            self.lbl_status_global.configure(text=f" {texto}")

    def _log(self, texto):
        agora = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, f"[{agora}] {texto}\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def aplicar_tema(self, tema_nome):
        paleta = self.paletas.get(tema_nome, self.paletas.get("Nexus", DEFAULT_PALETAS.get("Nexus")))
        bg = paleta.get("bg", "#0a192f")
        container = paleta.get("container", "#172a45")
        texto = paleta.get("texto", "#ffffff")
        input_bg = paleta.get("input", "#1f3a60")

        try:
            self.parent_frame.configure(fg_color=bg)
            if hasattr(self, "container"):
                self.container.configure(fg_color=container)
            if hasattr(self, "f_barra_fixa"):
                self.f_barra_fixa.configure(fg_color=container)
            for card in [getattr(self, "card_plan", None), getattr(self, "card_acoes", None), getattr(self, "card_tabela", None), getattr(self, "card_log", None)]:
                if card:
                    try:
                        card.configure(fg_color=container)
                    except Exception:
                        pass
            if hasattr(self, "entry_caminho"):
                self.entry_caminho.configure(fg_color=input_bg, text_color=texto)
            if hasattr(self, "txt_log"):
                self.txt_log.configure(fg_color=input_bg, text_color=texto)

            estilo = ttk.Style()
            if "clam" in estilo.theme_names():
                estilo.theme_use("clam")

            estilo.configure(
                "AtualizarOSME.Treeview",
                background=input_bg,
                foreground=texto,
                fieldbackground=input_bg,
                borderwidth=0,
                relief="flat",
                font=("Segoe UI", 9),
                rowheight=26
            )
            estilo.configure(
                "AtualizarOSME.Treeview.Heading",
                background=container,
                foreground="#38bdf8",
                font=("Segoe UI", 9, "bold"),
                borderwidth=1,
                relief="flat"
            )
            estilo.map(
                "AtualizarOSME.Treeview",
                background=[("selected", paleta.get("accent", "#00d2ff"))],
                foreground=[("selected", "#ffffff")]
            )
            estilo.map(
                "AtualizarOSME.Treeview.Heading",
                background=[("active", paleta.get("hover", "#0284c7"))],
                foreground=[("active", "#ffffff")]
            )
        except Exception:
            pass

    def _construir_ui(self):
        # 1. Barra de Ações Fixa no Rodapé (SEMPRE VISÍVEL)
        self.f_barra_fixa = ctk.CTkFrame(self.parent_frame, corner_radius=10, fg_color="#172a45")
        self.f_barra_fixa.pack(side="bottom", fill="x", padx=10, pady=(4, 10))

        f_fixa_in = ctk.CTkFrame(self.f_barra_fixa, fg_color="transparent")
        f_fixa_in.pack(fill="x", padx=14, pady=10)

        self.lbl_metricas_rodape = ctk.CTkLabel(
            f_fixa_in,
            text="Carregue a planilha para iniciar a varredura.",
            font=("Segoe UI", 11, "bold"),
            text_color="#94a3b8"
        )
        self.lbl_metricas_rodape.pack(side="left")

        self.btn_gravar_planilha = ctk.CTkButton(
            f_fixa_in,
            text="💾 Gravar Alterações na Planilha (Aba CADASTRO)",
            font=("Segoe UI", 12, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=38,
            command=self.gravar_alteracoes_planilha_acao
        )
        self.btn_gravar_planilha.pack(side="right")

        # 2. Frame com Scroll para o Conteúdo Principal
        self.frame_conteudo = ctk.CTkScrollableFrame(self.parent_frame, fg_color="transparent")
        self.frame_conteudo.pack(side="top", fill="both", expand=True, padx=10, pady=(10, 0))

        # Card 1: Configuração da Planilha
        self.card_plan = card_plan = ctk.CTkFrame(self.frame_conteudo, corner_radius=10, fg_color="#172a45")
        card_plan.pack(fill="x", pady=(0, 10))

        f_topo_plan = ctk.CTkFrame(card_plan, fg_color="transparent")
        f_topo_plan.pack(fill="x", padx=14, pady=(10, 6))

        ctk.CTkLabel(
            f_topo_plan,
            text="📂 PLANILHA DE CONTROLE (ABA CADASTRO — COLUNAS H=ORDEM / I=VENCIMENTO):",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        ).pack(side="left")

        f_inp_plan = ctk.CTkFrame(card_plan, fg_color="transparent")
        f_inp_plan.pack(fill="x", padx=14, pady=(0, 10))

        self.entry_caminho = ctk.CTkEntry(
            f_inp_plan,
            placeholder_text="Selecione a planilha CONTROLE_SERIAIS.xlsx...",
            font=("Segoe UI", 11),
            height=34
        )
        self.entry_caminho.pack(side="left", fill="x", expand=True, padx=(0, 8))
        if self.caminho_planilha:
            self.entry_caminho.insert(0, self.caminho_planilha)

        btn_sel = ctk.CTkButton(
            f_inp_plan,
            text="📁 Selecionar",
            font=("Segoe UI", 11, "bold"),
            width=100,
            height=34,
            command=self._selecionar_planilha
        )
        btn_sel.pack(side="left", padx=(0, 6))

        self.btn_recarregar = ctk.CTkButton(
            f_inp_plan,
            text="🔄 Recarregar Equipes",
            font=("Segoe UI", 11, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            width=145,
            height=34,
            command=self._iniciar_leitura_planilha_thread
        )
        self.btn_recarregar.pack(side="left")

        # Badges informativos
        f_badges = ctk.CTkFrame(card_plan, fg_color="transparent")
        f_badges.pack(fill="x", padx=14, pady=(0, 10))

        self.badge_equipes = ctk.CTkLabel(f_badges, text="📋 Equipes: 0", font=("Segoe UI", 11, "bold"), text_color="#38bdf8")
        self.badge_equipes.pack(side="left", padx=(0, 15))

        self.badge_atualizadas = ctk.CTkLabel(f_badges, text="🔄 Novas OSMEs: 0", font=("Segoe UI", 11, "bold"), text_color="#10b981")
        self.badge_atualizadas.pack(side="left", padx=(0, 15))

        self.badge_sem_ordem = ctk.CTkLabel(f_badges, text="⚠️ Sem Ordem SAP: 0", font=("Segoe UI", 11), text_color="#f59e0b")
        self.badge_sem_ordem.pack(side="left")

        # Banner de Pendências de Sincronização (Auto-Retry quando a planilha estiver em uso)
        self.f_banner_pendencias = ctk.CTkFrame(card_plan, fg_color="#78350f", corner_radius=6)

        self.lbl_banner_pendencias = ctk.CTkLabel(
            self.f_banner_pendencias,
            text="",
            font=("Segoe UI", 11, "bold"),
            text_color="#fef3c7"
        )
        self.lbl_banner_pendencias.pack(side="left", padx=12, pady=6)

        self.btn_sinc_banner = ctk.CTkButton(
            self.f_banner_pendencias,
            text="🔄 Sincronizar Agora",
            font=("Segoe UI", 10, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=28,
            width=135,
            command=self._sincronizar_pendencias_manual
        )
        self.btn_sinc_banner.pack(side="right", padx=(6, 10), pady=4)

        self.btn_ver_banner = ctk.CTkButton(
            self.f_banner_pendencias,
            text="📋 Ver Detalhes",
            font=("Segoe UI", 10),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=28,
            width=95,
            command=self._abrir_modal_pendencias
        )
        self.btn_ver_banner.pack(side="right", padx=(0, 4), pady=4)

        # Card 2: Barra de Ações e Filtros de Execução
        self.card_acoes = card_acoes = ctk.CTkFrame(self.frame_conteudo, corner_radius=10, fg_color="#172a45")
        card_acoes.pack(fill="x", pady=(0, 10))

        f_controles = ctk.CTkFrame(card_acoes, fg_color="transparent")
        f_controles.pack(fill="x", padx=14, pady=10)

        self.btn_iniciar = ctk.CTkButton(
            f_controles,
            text="▶ Iniciar Varredura SAP (IW39)",
            font=("Segoe UI", 12, "bold"),
            fg_color="#f57c00",
            hover_color="#e65100",
            height=36,
            command=self.iniciar_varredura_acao
        )
        self.btn_iniciar.pack(side="left", padx=(0, 8))

        self.btn_parar = ctk.CTkButton(
            f_controles,
            text="⏹ Parar",
            font=("Segoe UI", 11, "bold"),
            fg_color="#dc2626",
            hover_color="#b91c1c",
            width=80,
            height=36,
            state="disabled",
            command=self.parar_varredura_acao
        )
        self.btn_parar.pack(side="left", padx=(0, 15))

        # Checkboxes de Opção
        self.chk_apenas_vencidas = ctk.CTkCheckBox(
            f_controles,
            text="⚡ Varrer apenas OSMEs vencidas (Vencimento <= Hoje)",
            font=("Segoe UI", 10, "bold"),
            checkbox_width=18,
            checkbox_height=18
        )
        self.chk_apenas_vencidas.pack(side="left", padx=(0, 12))
        self.chk_apenas_vencidas.select()
        self.chk_apenas_desatualizadas = self.chk_apenas_vencidas  # retrocompatibilidade

        self.chk_backup_auto = ctk.CTkCheckBox(
            f_controles,
            text="Criar backup preventivo antes de salvar",
            font=("Segoe UI", 10),
            checkbox_width=18,
            checkbox_height=18
        )
        self.chk_backup_auto.pack(side="left")
        self.chk_backup_auto.select()

        # Barra de Progresso
        self.f_progresso = ctk.CTkFrame(card_acoes, fg_color="transparent")
        self.f_progresso.pack(fill="x", padx=14, pady=(0, 10))

        self.barra_progresso = ctk.CTkProgressBar(self.f_progresso, height=12, corner_radius=6)
        self.barra_progresso.pack(fill="x", pady=(2, 4))
        self.barra_progresso.set(0)

        self.lbl_progresso_texto = ctk.CTkLabel(
            self.f_progresso,
            text="Aguardando início da varredura...",
            font=("Segoe UI", 10),
            text_color="#94a3b8"
        )
        self.lbl_progresso_texto.pack(anchor="w")

        # Card 3: Tabela Comparativa (Treeview)
        self.card_tabela = card_tabela = ctk.CTkFrame(self.frame_conteudo, corner_radius=10, fg_color="#172a45")
        card_tabela.pack(fill="both", expand=True, pady=(0, 10))

        f_tit_tab = ctk.CTkFrame(card_tabela, fg_color="transparent")
        f_tit_tab.pack(fill="x", padx=14, pady=(10, 6))

        ctk.CTkLabel(
            f_tit_tab,
            text="📊 RESULTADOS DA VARREDURA NO SAP (IW39 ➡️ CADASTRO):",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        ).pack(side="left")

        self.lbl_contagem_itens = ctk.CTkLabel(
            f_tit_tab,
            text="0 equipes carregadas",
            font=("Segoe UI", 11, "bold"),
            text_color="#cbd5e1"
        )
        self.lbl_contagem_itens.pack(side="right")

        f_tree_container = ctk.CTkFrame(card_tabela, fg_color="transparent")
        f_tree_container.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        colunas = ("item", "equipe", "deposito", "segmento", "ordem_atual", "venc_atual", "ordem_sap", "venc_sap", "status")
        self.tree = ttk.Treeview(f_tree_container, columns=colunas, show="headings", height=10, style="AtualizarOSME.Treeview")

        self.tree.heading("item", text="#")
        self.tree.heading("equipe", text="Equipe (Z2)")
        self.tree.heading("deposito", text="Depósito")
        self.tree.heading("segmento", text="Segmento")
        self.tree.heading("ordem_atual", text="OSME Planilha (H)")
        self.tree.heading("venc_atual", text="Vencimento Planilha (I)")
        self.tree.heading("ordem_sap", text="Nova OSME SAP")
        self.tree.heading("venc_sap", text="Novo Vencimento SAP")
        self.tree.heading("status", text="Situação")

        self.tree.column("item", width=40, anchor="center")
        self.tree.column("equipe", width=100, anchor="center")
        self.tree.column("deposito", width=80, anchor="center")
        self.tree.column("segmento", width=100, anchor="center")
        self.tree.column("ordem_atual", width=130, anchor="center")
        self.tree.column("venc_atual", width=130, anchor="center")
        self.tree.column("ordem_sap", width=130, anchor="center")
        self.tree.column("venc_sap", width=130, anchor="center")
        self.tree.column("status", width=200, anchor="w")

        scrollbar = ttk.Scrollbar(f_tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.tree.tag_configure("atualizar", foreground="#34d399")
        self.tree.tag_configure("vigente", foreground="#94a3b8")
        self.tree.tag_configure("sem_ordem", foreground="#fbbf24")
        self.tree.tag_configure("pendente", foreground="#fcd34d")
        self.tree.tag_configure("erro", foreground="#f87171")

        # Card 4: Terminal de Log de Execução
        self.card_log = card_log = ctk.CTkFrame(self.frame_conteudo, corner_radius=10, fg_color="#172a45")
        card_log.pack(fill="x", pady=(0, 5))

        f_tit_log = ctk.CTkFrame(card_log, fg_color="transparent")
        f_tit_log.pack(fill="x", padx=14, pady=(8, 2))

        ctk.CTkLabel(f_tit_log, text="📜 LOG DE PROCESSAMENTO SAP:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8").pack(side="left")

        self.txt_log = ctk.CTkTextbox(card_log, height=85, font=("Consolas", 10))
        self.txt_log.pack(fill="x", padx=14, pady=(0, 10))
        self.txt_log.configure(state="disabled")

        self._log("Módulo Atualizar OSME inicializado.")

    def _selecionar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha CONTROLE_SERIAIS",
            filetypes=[("Arquivos Excel", "*.xlsx *.xlsm"), ("Todos os Arquivos", "*.*")]
        )
        if caminho:
            self.entry_caminho.delete(0, tk.END)
            self.entry_caminho.insert(0, caminho)
            self._salvar_caminho(caminho)
            self._iniciar_leitura_planilha_thread()

    def _iniciar_leitura_planilha_thread(self):
        caminho = self.entry_caminho.get().strip()
        if not caminho or not os.path.exists(caminho):
            messagebox.showwarning("Atualizar OSME", "Selecione uma planilha Excel válida.")
            return

        if self.carregando_planilha or self.executando:
            return

        self.carregando_planilha = True
        self.btn_recarregar.configure(state="disabled", text="⏳ Lendo...")
        self._status("Lendo aba CADASTRO da planilha...")
        self._log(f"Iniciando leitura da planilha: {os.path.basename(caminho)}")

        threading.Thread(target=self._ler_planilha_background, args=(caminho,), daemon=True).start()

    def _ler_planilha_background(self, caminho):
        try:
            with open(caminho, "rb") as f_raw:
                buf = io.BytesIO(f_raw.read())

            equipes_lidas = []
            try:
                wb = load_workbook(filename=buf, read_only=True, data_only=True)
                if "CADASTRO" not in wb.sheetnames:
                    self.parent_frame.after(0, lambda: self._erro_leitura_planilha("Aba 'CADASTRO' não encontrada na planilha."))
                    return

                ws = wb["CADASTRO"]
                idx_eq, idx_dep, idx_seg, idx_ord, idx_venc = None, None, None, None, None
                cabecalho_row = 1

                for r_num, row_vals in enumerate(ws.iter_rows(max_row=5, values_only=True), start=1):
                    for c_idx, val in enumerate(row_vals):
                        norm = normalizar_texto(val)
                        if norm in ("EQUIPE", "EQUIPES", "RECURSO") and idx_eq is None:
                            idx_eq = c_idx
                        elif norm in ("DEPOSITO", "DEP", "LGORT") and idx_dep is None:
                            idx_dep = c_idx
                        elif norm in ("SEGMENTO", "SEG") and idx_seg is None:
                            idx_seg = c_idx
                        elif norm in ("ORDEM", "OSME", "NUM_ORDEM") and idx_ord is None:
                            idx_ord = c_idx
                        elif norm in ("VENCIMEN", "VENCIMENTO", "VALIDADE", "DATA_FIM") and idx_venc is None:
                            idx_venc = c_idx
                    if idx_eq is not None and idx_ord is not None:
                        cabecalho_row = r_num
                        break

                # Fallbacks baseados na imagem oficial: E=Equipes (4), F=Depósito (5), G=Segmento (6), H=Ordem (7), I=Vencimento (8)
                if idx_eq is None: idx_eq = 4
                if idx_dep is None: idx_dep = 5
                if idx_seg is None: idx_seg = 6
                if idx_ord is None: idx_ord = 7
                if idx_venc is None: idx_venc = 8

                self.col_equipe_idx = idx_eq
                self.col_dep_idx = idx_dep
                self.col_seg_idx = idx_seg
                self.col_ordem_idx = idx_ord
                self.col_venc_idx = idx_venc

                for r_idx, row in enumerate(ws.iter_rows(min_row=cabecalho_row + 1, values_only=True), start=cabecalho_row + 1):
                    if not row or len(row) <= idx_eq or not row[idx_eq]:
                        continue
                    eq_val = str(row[idx_eq]).strip()
                    if not eq_val or eq_val.upper() in ("EQUIPES", "EQUIPE", "TOTAL"):
                        continue

                    dep_val = str(row[idx_dep]).strip() if len(row) > idx_dep and row[idx_dep] else "—"
                    seg_val = str(row[idx_seg]).strip() if len(row) > idx_seg and row[idx_seg] else "—"
                    ord_val = str(row[idx_ord]).strip() if len(row) > idx_ord and row[idx_ord] else "—"
                    venc_val = formatar_data_para_exibicao(row[idx_venc] if len(row) > idx_venc and row[idx_venc] else "—")

                    equipes_lidas.append({
                        "row_idx": r_idx,
                        "equipe": eq_val.upper(),
                        "deposito": dep_val,
                        "segmento": seg_val,
                        "ordem_atual": ord_val,
                        "venc_atual": venc_val,
                        "ordem_sap": "",
                        "venc_sap": "",
                        "status": "Aguardando varredura",
                        "tag": ""
                    })

                self.parent_frame.after(0, lambda: self._concluir_leitura_planilha(equipes_lidas))
            finally:
                try:
                    wb.close()
                except Exception:
                    pass
                try:
                    buf.close()
                except Exception:
                    pass
                try:
                    import gc
                    gc.collect()
                except Exception:
                    pass

        except Exception as e:
            self.parent_frame.after(0, lambda e=e: self._erro_leitura_planilha(str(e)))

    def _concluir_leitura_planilha(self, equipes):
        self.carregando_planilha = False
        self.btn_recarregar.configure(state="normal", text="🔄 Recarregar Equipes")
        self.dados_equipes = equipes

        self.badge_equipes.configure(text=f"📋 Equipes: {len(equipes)}")
        self.badge_atualizadas.configure(text="🔄 Novas OSMEs: 0")
        self.badge_sem_ordem.configure(text="⚠️ Sem Ordem SAP: 0")

        self._renderizar_tabela()
        self._status(f"Leitura concluída: {len(equipes)} equipes carregadas da aba CADASTRO.")
        self._log(f"✅ {len(equipes)} equipes identificadas (Col H: Ordem | Col I: Vencimento).")

    def _erro_leitura_planilha(self, erro_msg):
        self.carregando_planilha = False
        self.btn_recarregar.configure(state="normal", text="🔄 Recarregar Equipes")
        self._status("Erro ao carregar planilha.")
        self._log(f"❌ Erro ao ler planilha: {erro_msg}")
        messagebox.showerror("Atualizar OSME — Erro", f"Não foi possível ler a planilha:\n\n{erro_msg}")

    def _renderizar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        total_novas = 0
        total_sem_ordem = 0

        for i, item in enumerate(self.dados_equipes, start=1):
            tag = item.get("tag", "")
            if tag == "atualizar":
                total_novas += 1
            elif tag == "sem_ordem":
                total_sem_ordem += 1

            self.tree.insert(
                "",
                tk.END,
                iid=str(i - 1),
                values=(
                    i,
                    item["equipe"],
                    item["deposito"],
                    item["segmento"],
                    item["ordem_atual"] or "—",
                    item["venc_atual"] or "—",
                    item["ordem_sap"] or "—",
                    item["venc_sap"] or "—",
                    item["status"]
                ),
                tags=(tag,)
            )

        self.lbl_contagem_itens.configure(text=f"{len(self.dados_equipes)} equipes")
        self.badge_atualizadas.configure(text=f"🔄 Novas OSMEs: {total_novas}")
        self.badge_sem_ordem.configure(text=f"⚠️ Sem Ordem SAP: {total_sem_ordem}")
        self.lbl_metricas_rodape.configure(
            text=f"Status: {len(self.dados_equipes)} equipes | {total_novas} novas OSMEs para gravar na planilha."
        )

    def iniciar_varredura_acao(self):
        if self.executando:
            return

        if not self.dados_equipes:
            messagebox.showwarning("Atualizar OSME", "Carregue a planilha antes de iniciar a varredura.")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("Atualizar OSME", "O SAP GUI precisa estar aberto e conectado para consultar a IW39.")
            return

        self.executando = True
        self.cancelar_execucao = False
        self.btn_iniciar.configure(state="disabled", text="⏳ Varrendo SAP...", fg_color="#f59e0b")
        self.btn_parar.configure(state="normal")
        self.barra_progresso.set(0)
        self._status("Iniciando varredura no SAP IW39...")
        self._log(f"🚀 Iniciando varredura IW39 para {len(self.dados_equipes)} equipes no SAP...")

        threading.Thread(target=self._thread_varredura_sap, daemon=True).start()

    def parar_varredura_acao(self):
        if not self.executando:
            return
        self.cancelar_execucao = True
        self._log("⚠️ Solicitação de cancelamento enviada. Finalizando equipe atual...")

    def _gerar_script_varredura_vbs(self, equipes_lista):
        """
        Gera um script VBScript único que processa a lista inteira de equipes sem reiniciar o SAP,
        tratando os dois cenários (ordem única e grade ALV com ordenação decrescente por GLTRP).
        """
        linhas_eq = ", ".join([f'"{eq.lower()}"' for eq in equipes_lista])
        return f'''
On Error Resume Next
Dim SapGuiAuto, app, connection, session
Set SapGuiAuto = GetObject("SAPGUI")
If Not IsObject(SapGuiAuto) Then
    WScript.Echo "ERRO: SAP GUI não encontrado."
    WScript.Quit 1
End If
Set app = SapGuiAuto.GetScriptingEngine
Set connection = app.Children(0)
Set session = connection.Children(0)
session.findById("wnd[0]").maximize

' Função auxiliar para busca recursiva de controle por padrão de ID
Function FindControlByPattern(parentObj, patternStr)
    Set FindControlByPattern = Nothing
    If parentObj Is Nothing Then Exit Function
    On Error Resume Next
    Dim idx, c, res
    For idx = 0 To parentObj.Children.Count - 1
        Set c = parentObj.Children(CInt(idx))
        If InStr(UCase(c.Id), UCase(patternStr)) > 0 Then
            Set FindControlByPattern = c
            Exit Function
        ElseIf c.Children.Count > 0 Then
            Set res = FindControlByPattern(c, patternStr)
            If Not res Is Nothing Then
                Set FindControlByPattern = res
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function

' Função auxiliar para busca de controles de VALOR (GuiCTextField, GuiTextField), IGNORANDO GuiLabel e títulos
Function FindValueControlByPattern(parentObj, patternStr)
    Set FindValueControlByPattern = Nothing
    If parentObj Is Nothing Then Exit Function
    On Error Resume Next
    Dim idx, c, res
    For idx = 0 To parentObj.Children.Count - 1
        Set c = parentObj.Children(CInt(idx))
        If c.Type <> "GuiLabel" And c.Type <> "GuiTitlebar" And InStr(UCase(c.Id), "/LBL") = 0 Then
            If InStr(UCase(c.Id), UCase(patternStr)) > 0 Then
                Dim valTxt
                valTxt = Trim(c.Text)
                If InStr(LCase(valTxt), "fim-base") = 0 And InStr(LCase(valTxt), "iníciobase") = 0 And InStr(LCase(valTxt), "iniciobase") = 0 Then
                    Set FindValueControlByPattern = c
                    Exit Function
                End If
            End If
        End If
        If c.Children.Count > 0 Then
            Set res = FindValueControlByPattern(c, patternStr)
            If Not res Is Nothing Then
                Set FindValueControlByPattern = res
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function

' Função auxiliar para busca de qualquer GridView ALV na tela
Function FindFirstGridView(parentObj)
    Set FindFirstGridView = Nothing
    If parentObj Is Nothing Then Exit Function
    On Error Resume Next
    Dim idx, c, res
    For idx = 0 To parentObj.Children.Count - 1
        Set c = parentObj.Children(CInt(idx))
        If c.Type = "GuiGridView" Or InStr(c.Id, "shellcont/shell") > 0 Then
            Set FindFirstGridView = c
            Exit Function
        ElseIf c.Children.Count > 0 Then
            Set res = FindFirstGridView(c)
            If Not res Is Nothing Then
                Set FindFirstGridView = res
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function

' Navega para a IW39
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

' Garante parâmetros iniciais padrão
session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/chkDY_OFN").selected = True
session.findById("wnd[0]/usr/chkDY_IAR").selected = True
session.findById("wnd[0]/usr/chkDY_MAB").selected = False
session.findById("wnd[0]/usr/chkDY_HIS").selected = False

Dim lista
lista = Array({linhas_eq})

Dim i, eq
For i = 0 To UBound(lista)
    eq = Trim(lista(i))
    If eq <> "" Then
        session.findById("wnd[0]/usr/ctxtDY_PARNR").text = eq
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        WScript.Sleep 650

        Do While session.Children.Count > 1
            session.findById("wnd[1]").sendVKey 0
            WScript.Sleep 200
        Loop

        Dim ordemEncontrada, vencEncontrado
        ordemEncontrada = ""
        vencEncontrado = ""

        ' =====================================================================
        ' CENÁRIO 1: Tela de Cabeçalho Individual (ex: CDC301 / CDC400)
        ' =====================================================================
        ' 1.1 Tenta extrair a ordem de 10 a 12 dígitos do título da janela ativa
        On Error Resume Next
        Dim wTitle, reOrd, mOrd
        wTitle = session.ActiveWindow.Text
        Set reOrd = New RegExp
        reOrd.Pattern = "\\b(1000\\d{{8}}|\\d{{10,12}})\\b"
        Set mOrd = reOrd.Execute(wTitle)
        If mOrd.Count > 0 Then
            ordemEncontrada = Trim(mOrd(0).Value)
        End If
        On Error GoTo 0

        ' 1.2 Tenta achar pelo controle CAUFVD-AUFNR (apenas campos de valor, não rótulos)
        If ordemEncontrada = "" Then
            On Error Resume Next
            Dim txtOrd
            Set txtOrd = FindValueControlByPattern(session.findById("wnd[0]/usr"), "ctxtCAUFVD-AUFNR")
            If txtOrd Is Nothing Then
                Set txtOrd = FindValueControlByPattern(session.findById("wnd[0]/usr"), "txtCAUFVD-AUFNR")
            End If
            If txtOrd Is Nothing Then
                Set txtOrd = FindValueControlByPattern(session.findById("wnd[0]/usr"), "CAUFVD-AUFNR")
            End If
            If Not txtOrd Is Nothing Then
                Dim rawOrd
                rawOrd = Trim(txtOrd.text)
                If IsNumeric(rawOrd) And Len(rawOrd) >= 8 Then
                    ordemEncontrada = rawOrd
                End If
            End If
            On Error GoTo 0
        End If

        ' 1.3 Se identificou que está no cabeçalho da ordem, busca o vencimento (Fim-base)
        If ordemEncontrada <> "" Then
            On Error Resume Next
            Dim txtGLTRP
            ' 1.3.1 Procura o campo de data ctxtCAUFVD-GLTRP (NUNCA o rótulo lblCAUFVD-GLTRP)
            Set txtGLTRP = FindValueControlByPattern(session.findById("wnd[0]/usr"), "ctxtCAUFVD-GLTRP")
            If txtGLTRP Is Nothing Then
                Set txtGLTRP = FindValueControlByPattern(session.findById("wnd[0]/usr"), "CAUFVD-GLTRP")
            End If
            If txtGLTRP Is Nothing Then
                Set txtGLTRP = FindValueControlByPattern(session.findById("wnd[0]/usr"), "GLTRP")
            End If
            If Not txtGLTRP Is Nothing Then
                vencEncontrado = Trim(txtGLTRP.text)
            End If

            ' 1.3.2 Fallback para caminhos absolutos conhecidos da tela da ordem no SAP GUI
            If InStr(LCase(vencEncontrado), "fim") > 0 Or InStr(LCase(vencEncontrado), "base") > 0 Or vencEncontrado = "" Then
                Dim cPath, directCtrl
                Set directCtrl = Nothing
                For Each cPath In Array( _
                    "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP", _
                    "wnd[0]/usr/subSUB_ALL:SAPLCOIH:0101/subSUB_HEADER:SAPLCOIH:0105/ctxtCAUFVD-GLTRP", _
                    "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/txtCAUFVD-GLTRP" _
                )
                    Set directCtrl = session.findById(cPath)
                    If Not directCtrl Is Nothing Then
                        If directCtrl.text <> "" Then
                            vencEncontrado = Trim(directCtrl.text)
                            Exit For
                        End If
                    End If
                Next
            End If

            ' 1.3.3 Sanitização rigorosa no VBScript: se não contiver números ou contiver rótulo, descarta
            If InStr(LCase(vencEncontrado), "fim") > 0 Or InStr(LCase(vencEncontrado), "base") > 0 Or InStr(LCase(vencEncontrado), "data") > 0 Then
                vencEncontrado = ""
            End If
            On Error GoTo 0

            ' Volta para a tela de pesquisa da IW39
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            WScript.Sleep 300
            If session.findById("wnd[0]/usr/ctxtDY_PARNR") Is Nothing Then
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                WScript.Sleep 250
            End If
        End If

        ' =====================================================================
        ' CENÁRIO 2: Entrou na grade ALV (Múltiplas ordens, ex: CDC300)
        ' =====================================================================
        If ordemEncontrada = "" Then
            Dim grid
            Set grid = FindFirstGridView(session.findById("wnd[0]"))
            If Not grid Is Nothing Then
                Dim totalLinhas, r, cOrd, cVenc, bestOrd, bestVenc
                totalLinhas = 0
                On Error Resume Next
                totalLinhas = grid.RowCount
                On Error GoTo 0
                bestOrd = ""
                bestVenc = ""
                For r = 0 To totalLinhas - 1
                    cOrd = ""
                    cVenc = ""
                    On Error Resume Next
                    cOrd = Trim(grid.GetCellValue(r, "AUFNR"))
                    If cOrd = "" Then cOrd = Trim(grid.GetCellValue(r, "ORDERID"))
                    cVenc = Trim(grid.GetCellValue(r, "GLTRP"))
                    If cVenc = "" Then cVenc = Trim(grid.GetCellValue(r, "GSTRP"))
                    On Error GoTo 0
                    
                    If bestOrd = "" Then
                        bestOrd = cOrd
                        bestVenc = cVenc
                    Else
                        ' Normaliza datas (DD.MM.YYYY para YYYYMMDD) para pegar a OSME mais recente
                        Dim d1, d2
                        d1 = cVenc: d2 = bestVenc
                        If Len(d1) = 10 And Mid(d1, 3, 1) = "." Then d1 = Mid(d1, 7, 4) & Mid(d1, 4, 2) & Mid(d1, 1, 2)
                        If Len(d2) = 10 And Mid(d2, 3, 1) = "." Then d2 = Mid(d2, 7, 4) & Mid(d2, 4, 2) & Mid(d2, 1, 2)
                        If d1 > d2 Then
                            bestOrd = cOrd
                            bestVenc = cVenc
                        End If
                    End If
                Next

                If InStr(LCase(bestVenc), "fim") > 0 Or InStr(LCase(bestVenc), "base") > 0 Then
                    bestVenc = ""
                End If

                ordemEncontrada = bestOrd
                vencEncontrado = bestVenc

                ' Volta para a tela de pesquisa da IW39
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                WScript.Sleep 300
                If session.findById("wnd[0]/usr/ctxtDY_PARNR") Is Nothing Then
                    session.findById("wnd[0]/tbar[0]/btn[3]").press
                    WScript.Sleep 250
                End If
            End If
        End If

        ' Emite resultado estruturado
        If ordemEncontrada <> "" Then
            WScript.Echo "RESULTADO_OSME:" & UCase(eq) & "|" & ordemEncontrada & "|" & vencEncontrado & "|OK"
        Else
            WScript.Echo "RESULTADO_OSME:" & UCase(eq) & "|||SEM_ORDEM"
        End If

        ' Garante que estamos prontos para a próxima equipe na tela da IW39
        Dim chkParnr
        Set chkParnr = Nothing
        On Error Resume Next
        Set chkParnr = session.findById("wnd[0]/usr/ctxtDY_PARNR")
        If chkParnr Is Nothing Then
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            WScript.Sleep 250
            Set chkParnr = session.findById("wnd[0]/usr/ctxtDY_PARNR")
            If chkParnr Is Nothing Then
                session.findById("wnd[0]/tbar[0]/btn[12]").press
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
                session.findById("wnd[0]").sendVKey 0
                WScript.Sleep 400
                session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
                session.findById("wnd[0]/usr/chkDY_OFN").selected = True
                session.findById("wnd[0]/usr/chkDY_IAR").selected = True
                session.findById("wnd[0]/usr/chkDY_MAB").selected = False
                session.findById("wnd[0]/usr/chkDY_HIS").selected = False
            End If
        End If
        On Error GoTo 0
    End If
Next
'''

    def _thread_varredura_sap(self):
        hoje = datetime.now().date()
        apenas_vencidas = getattr(self, "chk_apenas_vencidas", None) and self.chk_apenas_vencidas.get()

        equipes_totais = []
        qtd_vigentes_puladas = 0

        for item in self.dados_equipes:
            dt_venc = converter_para_datetime(item.get("venc_atual"))
            tem_ordem = item.get("ordem_atual") and str(item.get("ordem_atual")).strip() not in ("—", "-", "", "None")

            # Se o filtro inteligente estiver ativo e a OSME ainda estiver no prazo de validade:
            if apenas_vencidas and dt_venc is not None and dt_venc.date() > hoje and tem_ordem:
                item["ordem_sap"] = item["ordem_atual"]
                item["venc_sap"] = item["venc_atual"]
                item["status"] = f"✅ OSME já vigente ({item['venc_atual']})"
                item["tag"] = "vigente"
                qtd_vigentes_puladas += 1
            else:
                equipes_totais.append(item["equipe"])

        # Atualiza a tabela imediatamente na interface com as equipes vigentes já sinalizadas
        if qtd_vigentes_puladas > 0:
            self.parent_frame.after(0, self._renderizar_tabela)
            self._log(f"⚡ Filtro Inteligente: {qtd_vigentes_puladas} equipes com OSME vigente puladas!")
            self._log(f"🔎 Varrendo no SAP IW39 apenas {len(equipes_totais)} equipes com OSME vencida ou pendente...")

        total = len(equipes_totais)
        if total == 0:
            self._log("🏁 Todas as equipes já estão com OSME vigente! Nenhuma consulta necessária no SAP.")
            self.parent_frame.after(0, self._concluir_varredura)
            return

        script_vbs = self._gerar_script_varredura_vbs(equipes_totais)
        with tempfile.NamedTemporaryFile(mode="w", suffix=".vbs", delete=False, encoding="latin-1", errors="replace") as tf:
            tf.write(script_vbs)
            caminho_vbs = tf.name

        try:
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            proc = subprocess.Popen(
                ["cscript.exe", "//nologo", caminho_vbs],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=creationflags,
                bufsize=1
            )

            idx_processados = 0
            for linha in proc.stdout:
                if self.cancelar_execucao:
                    proc.terminate()
                    break

                linha_limpa = linha.strip()
                if "RESULTADO_OSME:" in linha_limpa:
                    partes = linha_limpa.split("RESULTADO_OSME:")[1].split("|")
                    eq_ret = partes[0].strip().upper()
                    ord_ret = partes[1].strip() if len(partes) > 1 else ""
                    venc_ret = partes[2].strip() if len(partes) > 2 else ""
                    st_ret = partes[3].strip() if len(partes) > 3 else ""

                    venc_formatado = formatar_data_para_exibicao(venc_ret)

                    idx_processados += 1
                    self.parent_frame.after(
                        0,
                        lambda e=eq_ret, o=ord_ret, v=venc_formatado, s=st_ret, p=idx_processados, tot=total:
                        self._atualizar_linha_varredura(e, o, v, s, p, tot)
                    )

            proc.wait()

        except Exception as ex:
            self.parent_frame.after(0, lambda ex=ex: self._log(f"❌ Erro na varredura SAP: {ex}"))
        finally:
            try:
                os.unlink(caminho_vbs)
            except Exception:
                pass
            self.parent_frame.after(0, self._concluir_varredura)

    def _atualizar_linha_varredura(self, equipe, nova_ordem, novo_venc, status_sap, progresso, total):
        pct = progresso / max(total, 1)
        self.barra_progresso.set(pct)
        self.lbl_progresso_texto.configure(text=f"Processando: {progresso}/{total} equipes ({int(pct*100)}%) — Equipe atual: {equipe}")

        # Encontra o item nos dados
        item_alvo = None
        item_idx = None
        for i, it in enumerate(self.dados_equipes):
            if it["equipe"] == equipe:
                item_alvo = it
                item_idx = i
                break

        if not item_alvo:
            return

        item_alvo["ordem_sap"] = nova_ordem
        item_alvo["venc_sap"] = novo_venc

        if status_sap == "SEM_ORDEM" or not nova_ordem:
            item_alvo["status"] = "⚠️ Nenhuma ordem ativa na IW39"
            item_alvo["tag"] = "sem_ordem"
            self._log(f"⚠️ {equipe}: Nenhuma ordem ativa encontrada na IW39.")
        else:
            # Compara com a OSME atual da planilha
            ord_antiga = str(item_alvo["ordem_atual"]).strip()
            venc_antigo = str(item_alvo["venc_atual"]).strip()

            if ord_antiga != nova_ordem:
                item_alvo["status"] = f"🔄 Nova OSME detectada ({ord_antiga} ➡️ {nova_ordem})"
                item_alvo["tag"] = "atualizar"
                self._log(f"🔄 {equipe}: Nova OSME {nova_ordem} (Venc: {novo_venc}) — Anterior: {ord_antiga}")
            elif venc_antigo != novo_venc and is_data_valida(novo_venc):
                item_alvo["status"] = f"📅 Novo Vencimento ({venc_antigo} ➡️ {novo_venc})"
                item_alvo["tag"] = "atualizar"
                self._log(f"📅 {equipe}: Vencimento atualizado para {novo_venc}")
            else:
                item_alvo["status"] = "✅ OSME e Vencimento já vigentes"
                item_alvo["tag"] = "vigente"
                self._log(f"✅ {equipe}: OSME {nova_ordem} já atualizada.")

        # Atualiza a linha no Treeview
        if item_idx is not None and self.tree.exists(str(item_idx)):
            self.tree.item(
                str(item_idx),
                values=(
                    item_idx + 1,
                    item_alvo["equipe"],
                    item_alvo["deposito"],
                    item_alvo["segmento"],
                    item_alvo["ordem_atual"] or "—",
                    item_alvo["venc_atual"] or "—",
                    item_alvo["ordem_sap"] or "—",
                    item_alvo["venc_sap"] or "—",
                    item_alvo["status"]
                ),
                tags=(item_alvo["tag"],)
            )

        # Atualiza badges
        novas = sum(1 for x in self.dados_equipes if x.get("tag") == "atualizar")
        sem_ord = sum(1 for x in self.dados_equipes if x.get("tag") == "sem_ordem")
        self.badge_atualizadas.configure(text=f"🔄 Novas OSMEs: {novas}")
        self.badge_sem_ordem.configure(text=f"⚠️ Sem Ordem SAP: {sem_ord}")
        self.lbl_metricas_rodape.configure(
            text=f"Status: {progresso}/{total} equipes | {novas} novas OSMEs para gravar na planilha."
        )

    def _concluir_varredura(self):
        self.executando = False
        self.btn_iniciar.configure(state="normal", text="▶ Iniciar Varredura SAP (IW39)", fg_color="#f57c00")
        self.btn_parar.configure(state="disabled")

        novas = sum(1 for x in self.dados_equipes if x.get("tag") == "atualizar")
        sem_ord = sum(1 for x in self.dados_equipes if x.get("tag") == "sem_ordem")

        self.barra_progresso.set(1.0)
        self.lbl_progresso_texto.configure(text=f"Varredura concluída! {novas} novas OSMEs identificadas no SAP.")
        self._status(f"Varredura finalizada. {novas} atualizações disponíveis.")
        self._log(f"🏁 Varredura concluída! Novas OSMEs: {novas} | Sem ordem ativa: {sem_ord}.")

        if novas > 0:
            if messagebox.askyesno(
                "Varredura Concluída — Gravar na Planilha?",
                f"Foram identificadas {novas} equipes com novas OSMEs no SAP!\n\nDeseja gravar as alterações agora nas colunas H (Ordem) e I (Vencimento) da planilha original?",
                icon="question"
            ):
                self.gravar_alteracoes_planilha_acao()
        else:
            messagebox.showinfo("Atualizar OSME", "Varredura concluída! Todas as equipes já estão com as OSMEs mais recentes.")

    def _iniciar_timer_pendencias(self):
        """Inicia o ciclo de verificação em segundo plano das pendências da planilha."""
        self._verificar_timer_pendencias()

    def _verificar_timer_pendencias(self):
        """Verifica a cada 20 segundos se a planilha foi liberada para gravar pendências."""
        try:
            qtd = obter_quantidade_pendencias()
            self._atualizar_banner_pendencias(qtd)
            if qtd > 0:
                caminho = self.entry_caminho.get().strip()
                if caminho and os.path.exists(caminho) and is_file_writable(caminho):
                    sucessos, restantes, msg = processar_todas_pendencias()
                    if sucessos > 0:
                        self._log(f"🎉 Sincronização Automática: {sucessos} operação(ões) gravada(s) na planilha com sucesso!")
                        self._status(f"🎉 {sucessos} operação(ões) gravada(s) na planilha automaticamente!")
                        self._atualizar_banner_pendencias(restantes)
                        self._iniciar_leitura_planilha_thread()
        except Exception:
            pass
        self.parent_frame.after(20000, self._verificar_timer_pendencias)

    def _atualizar_banner_pendencias(self, qtd=None):
        """Exibe ou oculta o banner de pendências com a contagem atual."""
        if qtd is None:
            qtd = obter_quantidade_pendencias()
        if not hasattr(self, "f_banner_pendencias"):
            return
        if qtd > 0:
            self.lbl_banner_pendencias.configure(
                text=f"⏳ Há {qtd} operação(ões) aguardando liberação da planilha pelo outro usuário."
            )
            if not self.f_banner_pendencias.winfo_ismapped():
                self.f_banner_pendencias.pack(fill="x", padx=14, pady=(0, 8))
        else:
            if self.f_banner_pendencias.winfo_ismapped():
                self.f_banner_pendencias.pack_forget()

    def _sincronizar_pendencias_manual(self):
        """Acionamento manual do botão de sincronização com feedback visual."""
        self.btn_sinc_banner.configure(text="⏳ Gravando...", state="disabled")
        self.parent_frame.update_idletasks()
        sucessos, restantes, msg = processar_todas_pendencias()
        self._atualizar_banner_pendencias(restantes)
        self.btn_sinc_banner.configure(text="🔄 Sincronizar Agora", state="normal")
        if sucessos > 0 and restantes == 0:
            messagebox.showinfo(
                "Sincronização Concluída",
                f"🎉 Sucesso total!\n\nForam gravados {sucessos} registro(s) pendente(s) na planilha com êxito!"
            )
            self._iniciar_leitura_planilha_thread()
        elif sucessos > 0 and restantes > 0:
            messagebox.showinfo(
                "Sincronização Parcial",
                f"✓ {sucessos} gravados com sucesso.\nRestam {restantes} que continuam em uso na rede."
            )
            self._iniciar_leitura_planilha_thread()
        else:
            messagebox.showwarning(
                "Planilha em Uso",
                "A planilha continua aberta por outro usuário na rede ou em outra janela do Excel.\n\n"
                "O N.E.X.U.S continuará tentando gravar automaticamente a cada 20 segundos em segundo plano."
            )

    def _abrir_modal_pendencias(self):
        """Abre o modal para inspecionar, copiar dados ou descartar pendências."""
        ModalVisualizarPendencias(
            self.parent_frame,
            callback_apos_sincronizacao=lambda: [self._atualizar_banner_pendencias(), self._iniciar_leitura_planilha_thread()],
            paletas=self.paletas,
            obter_tema_atual=self.obter_tema_atual
        )

    def gravar_alteracoes_planilha_acao(self):
        itens_para_gravar = [
            it for it in self.dados_equipes
            if it.get("tag") in ("atualizar", "pendente") and (it.get("ordem_sap") or it.get("ordem_atual"))
        ]
        if not itens_para_gravar:
            messagebox.showinfo("Gravar na Planilha", "Não há novas OSMEs ou alterações para gravar na planilha.")
            return

        caminho = self.entry_caminho.get().strip()
        if not caminho or not os.path.exists(caminho):
            messagebox.showwarning("Gravar na Planilha", "Caminho da planilha inválido ou inacessível.")
            return

        col_dep_1based = self.col_dep_idx + 1
        col_seg_1based = self.col_seg_idx + 1
        col_ordem_1based = self.col_ordem_idx + 1
        col_venc_1based = self.col_venc_idx + 1

        # 1. Verifica se a planilha está aberta no Excel ativo do usuário
        wb_com = None
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            xl_ativo = win32com.client.GetActiveObject("Excel.Application")
            if xl_ativo is not None:
                for wb_inst in xl_ativo.Workbooks:
                    try:
                        if os.path.abspath(wb_inst.FullName).lower() == os.path.abspath(caminho).lower():
                            wb_com = wb_inst
                            break
                    except Exception:
                        pass
        except Exception:
            wb_com = None

        # 1.1 Se o usuário está com ela aberta no Excel de forma editável:
        if wb_com is not None and not getattr(wb_com, "ReadOnly", False):
            try:
                ws_com = None
                for sh in wb_com.Sheets:
                    if sh.Name.upper() == "CADASTRO":
                        ws_com = sh
                        break

                if ws_com is not None:
                    # Backup preventivo antes de gravar
                    if self.chk_backup_auto.get():
                        try:
                            pasta_backup = os.path.join(os.path.dirname(caminho), "Backup_NEXUS")
                            os.makedirs(pasta_backup, exist_ok=True)
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            nome_base = os.path.basename(caminho)
                            caminho_bak = os.path.join(pasta_backup, f"BACKUP_PRE_ATUALIZAR_OSME_{timestamp}_{nome_base}")
                            shutil.copy2(caminho, caminho_bak)
                            self._log(f"🛡️ Backup preventivo criado: {os.path.basename(caminho_bak)}")
                        except Exception as e_bak:
                            self._log(f"Aviso ao criar backup: {e_bak}")

                    # Mapeamento dinâmico em tempo real de equipes por linha
                    # para nunca gravar na linha errada caso a planilha tenha sido classificada ou reordenada
                    col_equipe_1based = self.col_equipe_idx + 1
                    mapa_linhas_equipes = {}
                    try:
                        max_r = ws_com.UsedRange.Rows.Count
                        for r in range(2, max_r + 2):
                            v_eq = ws_com.Cells(r, col_equipe_1based).Value
                            if v_eq:
                                norm_eq = normalizar_texto(v_eq)
                                if norm_eq and norm_eq not in mapa_linhas_equipes:
                                    mapa_linhas_equipes[norm_eq] = r
                    except Exception as e_map:
                        self._log(f"Aviso mapeamento dinâmico Excel: {e_map}")

                    for item in itens_para_gravar:
                        norm_it = normalizar_texto(item["equipe"])
                        r_idx = mapa_linhas_equipes.get(norm_it, item.get("row_idx"))

                        # Validação dupla: verifica se a célula de fato contém a equipe
                        val_confere = normalizar_texto(ws_com.Cells(r_idx, col_equipe_1based).Value)
                        if val_confere != norm_it:
                            # Busca varrendo a coluna se a posição tiver mudado
                            achou_linha = False
                            for r_tentativa in range(2, max(300, ws_com.UsedRange.Rows.Count + 2)):
                                if normalizar_texto(ws_com.Cells(r_tentativa, col_equipe_1based).Value) == norm_it:
                                    r_idx = r_tentativa
                                    achou_linha = True
                                    break
                            if not achou_linha:
                                self._log(f"⚠️ Alerta: Equipe {item['equipe']} não encontrada no CADASTRO. Pulando para segurança.")
                                continue

                        nova_ordem = str(item.get("ordem_sap") or item.get("ordem_atual") or "").strip()
                        novo_venc = item.get("venc_sap") or item.get("venc_atual")

                        if nova_ordem and nova_ordem not in ("—", "-"):
                            ws_com.Cells(r_idx, col_ordem_1based).Value = nova_ordem

                        dt_venc = converter_para_datetime(novo_venc)
                        if dt_venc is not None:
                            ws_com.Cells(r_idx, col_venc_1based).Value = dt_venc
                            try:
                                ws_com.Cells(r_idx, col_venc_1based).NumberFormat = "dd/mm/yyyy"
                            except Exception:
                                pass
                        elif novo_venc and is_data_valida(novo_venc):
                            ws_com.Cells(r_idx, col_venc_1based).Value = str(novo_venc).strip()
                            try:
                                ws_com.Cells(r_idx, col_venc_1based).NumberFormat = "dd/mm/yyyy"
                            except Exception:
                                pass

                    wb_com.Save()

                    for item in itens_para_gravar:
                        if item.get("ordem_sap"):
                            item["ordem_atual"] = item["ordem_sap"]
                        if item.get("venc_sap"):
                            item["venc_atual"] = item["venc_sap"]
                        item["status"] = "💾 Gravado com sucesso na planilha"
                        item["tag"] = "vigente"

                    self._renderizar_tabela()
                    self._status(f"Sucesso! {len(itens_para_gravar)} equipes atualizadas na planilha.")
                    self._log(f"💾 {len(itens_para_gravar)} equipes atualizadas com sucesso via Excel ativo!")

                    messagebox.showinfo(
                        "Planilha Atualizada com Sucesso!",
                        f"🎉 Foram gravadas {len(itens_para_gravar)} OSMEs e datas de vencimento na aba CADASTRO do seu Excel aberto!\n\nArquivo: {os.path.basename(caminho)}"
                    )
                    return
            except Exception as ex_com:
                self._log(f"Aviso ao gravar via Excel aberto: {ex_com}. Tentando gravação direta...")

        # 2. Gravação direta via openpyxl (arquivo fechado ou processo em segundo plano)
        limpar_processos_excel_fantasmas()
        time.sleep(0.3)

        if not is_file_writable(caminho):
            # A PLANILHA ESTÁ ABERTA / EM USO NA REDE:
            # Fallback estilo Saída OSME -> adiciona à Fila de Pendências!
            adicionar_pendencia(
                tipo="atualizar_osme",
                caminho_planilha=caminho,
                dados_operacao={
                    "itens": [
                        {
                            "equipe": it["equipe"],
                            "row_idx": it["row_idx"],
                            "ordem_sap": it.get("ordem_sap") or it.get("ordem_atual"),
                            "venc_sap": it.get("venc_sap") or it.get("venc_atual"),
                            "deposito": it.get("deposito"),
                            "segmento": it.get("segmento")
                        }
                        for it in itens_para_gravar
                    ],
                    "col_equipe_idx": self.col_equipe_idx,
                    "col_dep_idx": self.col_dep_idx,
                    "col_seg_idx": self.col_seg_idx,
                    "col_ordem_idx": self.col_ordem_idx,
                    "col_venc_idx": self.col_venc_idx
                },
                motivo="Planilha aberta no Excel / bloqueada na rede"
            )

            for item in itens_para_gravar:
                item["status"] = "⏳ Na Fila de Pendências (gravará automaticamente)"
                item["tag"] = "pendente"

            self._renderizar_tabela()
            self._atualizar_banner_pendencias()
            self._log(f"⏳ {len(itens_para_gravar)} atualizações adicionadas à Fila de Pendências! O N.E.X.U.S gravará automaticamente assim que a planilha for liberada.")
            self._status(f"⏳ {len(itens_para_gravar)} atualizações adicionadas à Fila de Pendências.")

            messagebox.showinfo(
                "Adicionado à Fila de Pendências",
                f"⏳ A planilha está aberta no Excel ou bloqueada na rede:\n\n{caminho}\n\n"
                f"As alterações de {len(itens_para_gravar)} equipes foram adicionadas à Fila de Pendências do N.E.X.U.S!\n\n"
                "O sistema tentará gravar automaticamente em segundo plano a cada 20 segundos assim que o arquivo for liberado.\n"
                "Você também pode acompanhar pelo banner no topo ou clicar em '🔄 Sincronizar Agora'."
            )
            return

        try:
            # Backup preventivo
            if self.chk_backup_auto.get():
                pasta_backup = os.path.join(os.path.dirname(caminho), "Backup_NEXUS")
                os.makedirs(pasta_backup, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_base = os.path.basename(caminho)
                caminho_bak = os.path.join(pasta_backup, f"BACKUP_PRE_ATUALIZAR_OSME_{timestamp}_{nome_base}")
                shutil.copy2(caminho, caminho_bak)
                self._log(f"🛡️ Backup preventivo criado: {os.path.basename(caminho_bak)}")

            wb = load_workbook(caminho, keep_vba=caminho.lower().endswith(".xlsm"))
            if "CADASTRO" not in wb.sheetnames:
                raise ValueError("A aba 'CADASTRO' não foi encontrada na planilha.")

            ws = wb["CADASTRO"]
            col_equipe_1based = self.col_equipe_idx + 1

            # Mapeamento dinâmico em tempo real de equipes por linha
            mapa_linhas_equipes = {}
            for r in range(2, ws.max_row + 1):
                v_eq = ws.cell(row=r, column=col_equipe_1based).value
                if v_eq:
                    norm_eq = normalizar_texto(v_eq)
                    if norm_eq and norm_eq not in mapa_linhas_equipes:
                        mapa_linhas_equipes[norm_eq] = r

            for item in itens_para_gravar:
                norm_it = normalizar_texto(item["equipe"])
                r_idx = mapa_linhas_equipes.get(norm_it, item.get("row_idx"))

                # Validação dupla: checa se a célula contém a equipe
                val_confere = normalizar_texto(ws.cell(row=r_idx, column=col_equipe_1based).value)
                if val_confere != norm_it:
                    achou_linha = False
                    for r_tentativa in range(2, ws.max_row + 1):
                        if normalizar_texto(ws.cell(row=r_tentativa, column=col_equipe_1based).value) == norm_it:
                            r_idx = r_tentativa
                            achou_linha = True
                            break
                    if not achou_linha:
                        self._log(f"⚠️ Alerta: Equipe {item['equipe']} não encontrada no CADASTRO. Pulando para segurança.")
                        continue

                nova_ordem = str(item.get("ordem_sap") or item.get("ordem_atual") or "").strip()
                novo_venc = item.get("venc_sap") or item.get("venc_atual")

                # Atualiza Coluna H (Ordem)
                if nova_ordem and nova_ordem not in ("—", "-"):
                    ws.cell(row=r_idx, column=col_ordem_1based, value=nova_ordem)

                # Atualiza Coluna I (Vencimento) com datetime nativo e formato dd/mm/yyyy
                dt_venc = converter_para_datetime(novo_venc)
                if dt_venc is not None:
                    c_v = ws.cell(row=r_idx, column=col_venc_1based, value=dt_venc)
                    c_v.number_format = "dd/mm/yyyy"
                elif novo_venc and is_data_valida(novo_venc):
                    c_v = ws.cell(row=r_idx, column=col_venc_1based, value=str(novo_venc).strip())
                    c_v.number_format = "dd/mm/yyyy"

            # Tenta salvar no caminho oficial
            salvo_sucesso = False
            caminho_final = caminho
            try:
                try:
                    wb.save(caminho)
                    salvo_sucesso = True
                except PermissionError:
                    # Se ocorrer PermissionError no momento da gravação, adiciona à fila de pendências!
                    adicionar_pendencia(
                        tipo="atualizar_osme",
                        caminho_planilha=caminho,
                        dados_operacao={
                            "itens": [
                                {
                                    "equipe": it["equipe"],
                                    "row_idx": it["row_idx"],
                                    "ordem_sap": it.get("ordem_sap") or it.get("ordem_atual"),
                                    "venc_sap": it.get("venc_sap") or it.get("venc_atual"),
                                    "deposito": it.get("deposito"),
                                    "segmento": it.get("segmento")
                                }
                                for it in itens_para_gravar
                            ],
                            "col_equipe_idx": self.col_equipe_idx,
                            "col_dep_idx": self.col_dep_idx,
                            "col_seg_idx": self.col_seg_idx,
                            "col_ordem_idx": self.col_ordem_idx,
                            "col_venc_idx": self.col_venc_idx
                        },
                        motivo="Bloqueio de escrita detectado ao salvar"
                    )

                    for item in itens_para_gravar:
                        item["status"] = "⏳ Na Fila de Pendências (gravará automaticamente)"
                        item["tag"] = "pendente"

                    self._renderizar_tabela()
                    self._atualizar_banner_pendencias()
                    self._log(f"⏳ Bloqueio ao salvar. {len(itens_para_gravar)} equipes adicionadas à Fila de Pendências!")

                    messagebox.showinfo(
                        "Adicionado à Fila de Pendências",
                        f"A planilha foi bloqueada durante a gravação.\n\n"
                        f"As alterações de {len(itens_para_gravar)} equipes foram adicionadas à Fila de Pendências e serão gravadas automaticamente assim que o Excel for liberado!"
                    )
                    return
            finally:
                try:
                    wb.close()
                except Exception:
                    pass

            if salvo_sucesso:
                for item in itens_para_gravar:
                    if item.get("ordem_sap"):
                        item["ordem_atual"] = item["ordem_sap"]
                    if item.get("venc_sap"):
                        item["venc_atual"] = item["venc_sap"]
                    item["status"] = "💾 Gravado com sucesso na planilha"
                    item["tag"] = "vigente"

                self._renderizar_tabela()
                self._status(f"Sucesso! {len(itens_para_gravar)} equipes atualizadas na planilha.")
                self._log(f"💾 {len(itens_para_gravar)} equipes atualizadas com sucesso em {os.path.basename(caminho_final)}!")

                messagebox.showinfo(
                    "Planilha Atualizada com Sucesso!",
                    f"🎉 Foram gravadas {len(itens_para_gravar)} OSMEs e datas de vencimento na aba CADASTRO!\n\nArquivo: {os.path.basename(caminho_final)}"
                )

        except Exception as ex:
            self._log(f"❌ Erro ao gravar planilha: {ex}")
            messagebox.showerror("Erro de Gravação", f"Falha ao salvar alterações na planilha:\n\n{ex}")


def construir_aba_atualizar_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaAtualizarOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
