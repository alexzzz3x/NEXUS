import tkinter as tk
import customtkinter as ctk
from utils_metrics import ler_metricas

DEFAULT_PALETAS = {
    "Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#f57c00", "hover": "#e65100", "input": "#1f3a60"}
}

class AbaDashboard:
    def __init__(self, parent_frame, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")
        self.cards = {}
        
        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self.atualizar_metricas()

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Conecta Blue"))

    def _criar_card(self, parent, titulo, valor_inicial="0", icone="📊", color="container"):
        frame = ctk.CTkFrame(parent, corner_radius=12)
        
        lbl_icone = ctk.CTkLabel(frame, text=icone, font=("Segoe UI", 36))
        lbl_icone.pack(pady=(15, 5))
        
        lbl_titulo = ctk.CTkLabel(frame, text=titulo, font=("Segoe UI", 13, "bold"))
        lbl_titulo.pack(pady=5)
        
        lbl_valor = ctk.CTkLabel(frame, text=str(valor_inicial), font=("Segoe UI", 28, "bold"))
        lbl_valor.pack(pady=(0, 15))
        
        return frame, lbl_valor

    def _construir_interface(self):
        self.lbl_welcome = ctk.CTkLabel(self.parent_frame, text="✨ Bem-vindo ao Conecta Hub 9.0", font=("Segoe UI", 24, "bold"))
        self.lbl_welcome.pack(pady=(30, 10))
        
        self.lbl_subtitle = ctk.CTkLabel(self.parent_frame, text="Visão Geral de Desempenho e Ferramentas", font=("Segoe UI", 14))
        self.lbl_subtitle.pack(pady=(0, 30))

        self.grid_frame = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.grid_frame.pack(fill="both", expand=True, padx=40)
        
        # Configure grid columns
        self.grid_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="col")

        # Cards
        self.card_osme, self.lbl_val_osme = self._criar_card(self.grid_frame, "OSMEs Transferidos", icone="📦")
        self.card_osme.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.card_etiquetas, self.lbl_val_etiquetas = self._criar_card(self.grid_frame, "Etiquetas Geradas", icone="🎫")
        self.card_etiquetas.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.card_seriais, self.lbl_val_seriais = self._criar_card(self.grid_frame, "Seriais Extraídos", icone="🔢")
        self.card_seriais.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")

        self.btn_atualizar = ctk.CTkButton(self.parent_frame, text="🔄 Atualizar Métricas", command=self.atualizar_metricas, font=("Segoe UI", 12, "bold"), width=200, height=40)
        self.btn_atualizar.pack(pady=30)

    def atualizar_metricas(self):
        metricas = ler_metricas()
        
        val_osme = metricas.get("osmes_transferidos", 0)
        val_etiquetas = metricas.get("etiquetas_geradas", 0)
        val_seriais = metricas.get("seriais_extraidos", 0)
        
        self.lbl_val_osme.configure(text=str(val_osme))
        self.lbl_val_etiquetas.configure(text=str(val_etiquetas))
        self.lbl_val_seriais.configure(text=str(val_seriais))

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))
        self.lbl_welcome.configure(text_color=c["texto"])
        self.lbl_subtitle.configure(text_color=c["texto"])
        
        for card in [self.card_osme, self.card_etiquetas, self.card_seriais]:
            card.configure(fg_color=c["container"])
            
        for lbl in [self.lbl_val_osme, self.lbl_val_etiquetas, self.lbl_val_seriais]:
            lbl.configure(text_color=c["accent"])
            
        self.btn_atualizar.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")

def construir_aba_dashboard(parent_frame, paletas=None, obter_tema_atual=None):
    return AbaDashboard(parent_frame, paletas, obter_tema_atual)
