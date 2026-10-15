import tkinter as tk
import customtkinter as ctk
from utils_metrics import ler_metricas

DEFAULT_PALETAS = {
    "Nexus": {
        "bg": "#090d16", "container": "#111827", "texto": "#f8fafc",
        "accent": "#00d2ff", "hover": "#0284c7", "input": "#1a2436"
    },
    "Azul Delray": {
        "bg": "#0a192f", "container": "#172a45", "texto": "#ffffff",
        "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"
    },
    "Conecta Blue": {
        "bg": "#0a192f", "container": "#172a45", "texto": "#ffffff",
        "accent": "#38bdf8", "hover": "#0284c7", "input": "#1f3a60"
    }
}

class AbaDashboard:
    def __init__(self, parent_frame, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.cards_info = []

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self.atualizar_metricas()

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

    def _criar_card(self, parent, titulo, valor_inicial="0", icone="📊"):
        frame = ctk.CTkFrame(parent, corner_radius=12)

        lbl_icone = ctk.CTkLabel(frame, text=icone, font=("Segoe UI", 32))
        lbl_icone.pack(pady=(14, 4))

        lbl_titulo = ctk.CTkLabel(frame, text=titulo, font=("Segoe UI", 12, "bold"))
        lbl_titulo.pack(pady=3)

        lbl_valor = ctk.CTkLabel(frame, text=str(valor_inicial), font=("Segoe UI", 26, "bold"))
        lbl_valor.pack(pady=(0, 14))

        return frame, lbl_valor, lbl_titulo

    def _construir_interface(self):
        # Cabeçalho Principal
        self.frame_header = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.frame_header.pack(fill="x", pady=(24, 6))

        self.lbl_welcome = ctk.CTkLabel(
            self.frame_header,
            text="✨ Bem-vindo ao N.E.X.U.S",
            font=("Segoe UI", 24, "bold")
        )
        self.lbl_welcome.pack(pady=(0, 4))

        self.lbl_subtitle = ctk.CTkLabel(
            self.frame_header,
            text="Visão Geral de Desempenho e Operações em Tempo Real",
            font=("Segoe UI", 13),
            text_color="#94a3b8"
        )
        self.lbl_subtitle.pack(pady=(0, 16))

        # Grade balanceada de 6 colunas para dispor 5 cards (3 na linha 1, 2 centralizados na linha 2)
        self.grid_frame = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.grid_frame.pack(fill="both", expand=True, padx=30, pady=(0, 10))

        for i in range(6):
            self.grid_frame.grid_columnconfigure(i, weight=1, uniform="col")

        # 1. Bandeiradas Executadas (Linha 0, Colunas 0-1)
        card_band, lbl_band, tit_band = self._criar_card(self.grid_frame, "Bandeiradas Realizadas", icone="🚩")
        card_band.grid(row=0, column=0, columnspan=2, padx=8, pady=8, sticky="nsew")

        # 2. Saídas de OSME (Linha 0, Colunas 2-3)
        card_saidas, lbl_saidas, tit_saidas = self._criar_card(self.grid_frame, "Saídas de OSME", icone="📤")
        card_saidas.grid(row=0, column=2, columnspan=2, padx=8, pady=8, sticky="nsew")

        # 3. OSMEs Transferidas (Linha 0, Colunas 4-5)
        card_osme, lbl_osme, tit_osme = self._criar_card(self.grid_frame, "OSMEs Transferidas", icone="📦")
        card_osme.grid(row=0, column=4, columnspan=2, padx=8, pady=8, sticky="nsew")

        # 4. Etiquetas Geradas (Linha 1, Colunas 1-2 - centralizado)
        card_etiq, lbl_etiq, tit_etiq = self._criar_card(self.grid_frame, "Etiquetas Geradas", icone="🎫")
        card_etiq.grid(row=1, column=1, columnspan=2, padx=8, pady=8, sticky="nsew")

        # 5. Fichas WorkForce (Linha 1, Colunas 3-4 - centralizado)
        card_fichas, lbl_fichas, tit_fichas = self._criar_card(self.grid_frame, "Fichas WorkForce", icone="📋")
        card_fichas.grid(row=1, column=3, columnspan=2, padx=8, pady=8, sticky="nsew")

        self.cards_info = [
            {"card": card_band, "lbl": lbl_band, "tit": tit_band, "chave": "bandeiradas_executadas"},
            {"card": card_saidas, "lbl": lbl_saidas, "tit": tit_saidas, "chave": "saidas_osme"},
            {"card": card_osme, "lbl": lbl_osme, "tit": tit_osme, "chave": "osmes_transferidos"},
            {"card": card_etiq, "lbl": lbl_etiq, "tit": tit_etiq, "chave": "etiquetas_geradas"},
            {"card": card_fichas, "lbl": lbl_fichas, "tit": tit_fichas, "chave": "fichas_abertas"},
        ]

        # Botão Atualizar Métricas
        self.btn_atualizar = ctk.CTkButton(
            self.parent_frame,
            text="🔄 Atualizar Métricas",
            command=self.atualizar_metricas,
            font=("Segoe UI", 12, "bold"),
            width=220,
            height=40,
            corner_radius=8
        )
        self.btn_atualizar.pack(pady=(10, 24))

    def atualizar_metricas(self):
        metricas = ler_metricas()
        for item in self.cards_info:
            valor = metricas.get(item["chave"], 0)
            item["lbl"].configure(text=str(valor))

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        bg = c.get("bg", "#090d16")
        container = c.get("container", "#111827")
        accent = c.get("accent", "#00d2ff")
        hover = c.get("hover", "#0284c7")
        texto = c.get("texto", "#f8fafc")

        self.parent_frame.configure(fg_color=bg)
        self.lbl_welcome.configure(text_color=texto)

        for item in self.cards_info:
            item["card"].configure(fg_color=container)
            item["tit"].configure(text_color=texto)
            item["lbl"].configure(text_color=accent)

        self.btn_atualizar.configure(fg_color=accent, hover_color=hover, text_color="#090d16" if "Light" not in tema_nome else "#ffffff")

def construir_aba_dashboard(parent_frame, paletas=None, obter_tema_atual=None):
    return AbaDashboard(parent_frame, paletas, obter_tema_atual)
