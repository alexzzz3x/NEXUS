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
    "Nexus": {
        "bg": "#090d16",
        "container": "#111827",
        "texto": "#f8fafc",
        "accent": "#00d2ff",
        "hover": "#0284c7",
        "input": "#1a2436",
    },
    "Azul Delray": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#38bdf8",
        "hover": "#0284c7",
        "input": "#1f3a60",
    },
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#38bdf8",
        "hover": "#0284c7",
        "input": "#1f3a60",
    }
}

EQUIPES_CLANDESTINAS = [
    "RCL001", "RCL002", "RCL003", "RCL004",
    "RCL005", "RCL006", "RCL007", "RCL008"
]


def _converter_qtd_para_float(qtd_str):
    """Converte quantidade de string do SAP para float de forma segura."""
    if not qtd_str:
        return 0.0
    q = str(qtd_str).strip()
    if not q:
        return 0.0
    import re
    q = re.sub(r'[,.]0+$', '', q)
    q = q.replace(',', '.')
    try:
        return float(q)
    except Exception:
        limpo = re.sub(r"[^\d.]", "", q)
        try:
            return float(limpo)
        except Exception:
            return 0.0

def _formatar_qtd_float(valor):
    """Formata valor float em string limpa sem zeros redundantes."""
    if valor == int(valor):
        return str(int(valor))
    s = f"{valor:.3f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


class JanelaSaldoRCL(ctk.CTkToplevel):
    """Pop-up modal dedicado para exibir APENAS os itens em verde (considerados) da MIGO, somando duplicados."""
    def __init__(self, parent, equipe, ordem, itens, paletas=None, obter_tema_atual=None, callback_abrir_sap=None):
        super().__init__(parent)
        self.equipe = equipe or ""
        self.ordem = ordem or ""
        self.itens = itens or []
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.callback_abrir_sap = callback_abrir_sap

        self.title(f"📦 Saldo Disponível — {self.equipe}")
        self.geometry("540x580")
        self.minsize(500, 520)
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))

        self._centralizar(parent)
        self._construir_interface()
        self.focus_force()

    def _centralizar(self, parent):
        try:
            self.update_idletasks()
            pw = parent.winfo_width() if parent else 1200
            ph = parent.winfo_height() if parent else 800
            px = parent.winfo_rootx() if parent else 100
            py = parent.winfo_rooty() if parent else 100
            w, h = 540, 580
            x = max(50, px + (pw - w) // 2)
            y = max(50, py + (ph - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry("540x580+200+150")

    def _eh_item_verde(self, recebedor):
        """Verifica se o recebedor corresponde ao padrão de item considerado (em verde)."""
        import re
        rec = (recebedor or "").strip().upper()
        rec_norm = re.sub(r'\s+', ' ', rec)
        m = re.search(r'(\d+)', self.equipe)
        sufixo = m.group(1)[-2:].zfill(2) if m else "04"
        esperado = f"EMPR 1CL{sufixo}".upper()
        return (esperado in rec_norm) or ("EMPR 1CL04" in rec_norm)

    def _obter_itens_verdes_agrupados(self):
        """Filtra apenas os itens em verde e agrupa por material somando as quantidades."""
        agrupados = {}
        for it in self.itens:
            recebedor = it.get("recebedor", "")
            if not self._eh_item_verde(recebedor):
                continue

            mat = str(it.get("material", "")).strip()
            if not mat:
                continue

            qtd = _converter_qtd_para_float(it.get("qtd", 0))
            un_raw = str(it.get("unidade", "")).strip()
            if un_raw.upper() in ["ST", "PC", "PÇ", "UN", "UND"]:
                un_limpa = "Pçs"
            elif un_raw.upper() in ["M", "MT", "MTR"]:
                un_limpa = "M"
            else:
                un_limpa = un_raw or "Pçs"

            if mat not in agrupados:
                agrupados[mat] = {
                    "material": mat,
                    "descricao": it.get("descricao", "").strip(),
                    "qtd_total": qtd,
                    "unidade": un_limpa,
                    "recebedor": it.get("recebedor", "").strip(),
                    "centro": it.get("centro", "EMPR").strip() or "EMPR",
                    "deposito": it.get("deposito", "CL03").strip() or "CL03",
                    "lancamentos": [qtd]
                }
            else:
                agrupados[mat]["qtd_total"] += qtd
                agrupados[mat]["lancamentos"].append(qtd)
                if not agrupados[mat]["descricao"] and it.get("descricao"):
                    agrupados[mat]["descricao"] = it.get("descricao").strip()

        resultado = []
        for mat, d in agrupados.items():
            d["qtd_formatada"] = _formatar_qtd_float(d["qtd_total"])
            resultado.append(d)
        return resultado

    def _construir_interface(self):
        c = self.paletas.get(self.obter_tema_atual(), self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))
        self.configure(fg_color=c.get("bg", "#0a192f"))

        container = ctk.CTkFrame(self, fg_color=c.get("container", "#172a45"), corner_radius=12, border_width=1, border_color="#334155")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        # 1. Cabeçalho
        f_header = ctk.CTkFrame(container, fg_color="transparent")
        f_header.pack(fill="x", padx=16, pady=(16, 10))

        lbl_icon = ctk.CTkLabel(f_header, text="📦", font=("Segoe UI", 26))
        lbl_icon.pack(side="left", padx=(0, 10))

        f_titles = ctk.CTkFrame(f_header, fg_color="transparent")
        f_titles.pack(side="left", fill="both", expand=True)

        lbl_title = ctk.CTkLabel(
            f_titles,
            text=f"SALDO DISPONÍVEL — {self.equipe}",
            font=("Segoe UI", 16, "bold"),
            text_color=c.get("texto", "white")
        )
        lbl_title.pack(anchor="w")

        ordem_txt = f"Ordem IW39: {self.ordem}" if self.ordem else "Consulta direta na MIGO"
        lbl_sub = ctk.CTkLabel(f_titles, text=f"{ordem_txt}  •  🟢 Somente Itens em Verde", font=("Segoe UI", 11), text_color="#34d399")
        lbl_sub.pack(anchor="w")

        # 2. Obter itens verdes agrupados e somados
        itens_verdes = self._obter_itens_verdes_agrupados()

        # Frame de Ações (Fixo no rodapé)
        f_acoes = ctk.CTkFrame(container, fg_color="transparent")
        f_acoes.pack(fill="x", padx=16, pady=(8, 14), side="bottom")

        self.btn_copiar = ctk.CTkButton(
            f_acoes,
            text="📋 Copiar Saldo",
            height=36,
            font=("Segoe UI", 12, "bold"),
            fg_color=c.get("accent", "#f57c00"),
            hover_color=c.get("hover", "#e65100"),
            command=self._copiar_saldo
        )
        self.btn_copiar.pack(side="left", fill="x", expand=True, padx=(0, 6))

        if self.callback_abrir_sap:
            btn_sap = ctk.CTkButton(
                f_acoes,
                text="🖥️ Abrir SAP",
                height=36,
                width=110,
                font=("Segoe UI", 11, "bold"),
                fg_color="#0284c7",
                hover_color="#0369a1",
                command=self.callback_abrir_sap
            )
            btn_sap.pack(side="left", padx=3)

        btn_fechar = ctk.CTkButton(
            f_acoes,
            text="❌ Fechar",
            height=36,
            width=90,
            font=("Segoe UI", 12),
            fg_color="#334155",
            hover_color="#475569",
            command=self.destroy
        )
        btn_fechar.pack(side="right", padx=(6, 0))

        # 3. Conteúdo Central (Apenas itens verdes)
        if not itens_verdes:
            f_aviso = ctk.CTkFrame(container, fg_color="#451a03", corner_radius=10, border_width=1, border_color="#d97706")
            f_aviso.pack(fill="both", expand=True, padx=16, pady=(10, 10))
            ctk.CTkLabel(
                f_aviso,
                text="⚠️ Nenhum saldo / item considerado em verde (EMPR 1CL...)\nencontrado para esta ordem na MIGO.",
                font=("Segoe UI", 13, "bold"),
                text_color="#fde68a",
                justify="center"
            ).pack(expand=True, padx=20, pady=20)
            self.texto_copiar = f"Equipe {self.equipe} (Ordem {self.ordem}): Sem saldo disponível em verde"
            return

        # Monta textos para cópia
        linhas_copiar = []
        for it in itens_verdes:
            linhas_copiar.append(f"{it['material']}  {it['qtd_formatada']} {it['unidade']}  {it['recebedor']}  ({it['descricao']})")
        self.texto_copiar = "\n".join(linhas_copiar)

        # Se houver apenas 1 material considerado
        if len(itens_verdes) == 1:
            it = itens_verdes[0]
            f_hero = ctk.CTkFrame(container, fg_color="#064e3b", corner_radius=10, border_width=1, border_color="#059669")
            f_hero.pack(fill="both", expand=True, padx=16, pady=(0, 10))

            lbl_tag = ctk.CTkLabel(f_hero, text="🎯 SALDO DISPONÍVEL (TOTAL CONSOLIDADO)", font=("Segoe UI", 11, "bold"), text_color="#a7f3d0")
            lbl_tag.pack(anchor="w", padx=16, pady=(14, 4))

            f_qtd_row = ctk.CTkFrame(f_hero, fg_color="transparent")
            f_qtd_row.pack(anchor="w", padx=16, pady=(0, 4))

            lbl_qtd_val = ctk.CTkLabel(f_qtd_row, text=it['qtd_formatada'], font=("Segoe UI", 36, "bold"), text_color="#34d399")
            lbl_qtd_val.pack(side="left")

            lbl_un_val = ctk.CTkLabel(f_qtd_row, text=f" {it['unidade']}", font=("Segoe UI", 18, "bold"), text_color="#a7f3d0")
            lbl_un_val.pack(side="left", padx=(6, 0), pady=(10, 0))

            # Se houve soma de múltiplos lançamentos duplicados (ex: 15 M + 50 M = 65 M)
            if len(it["lancamentos"]) > 1:
                detalhe_soma = " + ".join(_formatar_qtd_float(q) for q in it["lancamentos"])
                lbl_soma = ctk.CTkLabel(
                    f_hero,
                    text=f"∑ Somado de {len(it['lancamentos'])} lançamentos: {detalhe_soma} = {it['qtd_formatada']} {it['unidade']}",
                    font=("Segoe UI", 11, "bold"),
                    text_color="#6ee7b7"
                )
                lbl_soma.pack(anchor="w", padx=16, pady=(0, 8))

            lbl_mat_info = ctk.CTkLabel(
                f_hero,
                text=f"Material: {it['material']}  •  {it['descricao']}",
                font=("Segoe UI", 12, "bold"),
                text_color="white",
                wraplength=460,
                justify="left"
            )
            lbl_mat_info.pack(anchor="w", padx=16, pady=(4, 10))

            # Badges
            f_badges = ctk.CTkFrame(f_hero, fg_color="transparent")
            f_badges.pack(anchor="w", padx=16, pady=(0, 16))

            f_b_rec = ctk.CTkFrame(f_badges, fg_color="#0284c7", corner_radius=6)
            f_b_rec.pack(side="left", padx=(0, 6))
            ctk.CTkLabel(f_b_rec, text=f"Recebedor: {it['recebedor']}", font=("Segoe UI", 10, "bold"), text_color="white").pack(padx=8, pady=3)

            f_b_cen = ctk.CTkFrame(f_badges, fg_color="#1e293b", corner_radius=6)
            f_b_cen.pack(side="left", padx=(0, 6))
            ctk.CTkLabel(f_b_cen, text=f"Centro: {it['centro']}", font=("Segoe UI", 10), text_color="#cbd5e1").pack(padx=8, pady=3)

            f_b_dep = ctk.CTkFrame(f_badges, fg_color="#1e293b", corner_radius=6)
            f_b_dep.pack(side="left")
            ctk.CTkLabel(f_b_dep, text=f"Depósito: {it['deposito']}", font=("Segoe UI", 10), text_color="#fde047").pack(padx=8, pady=3)

        else:
            # Mais de um material considerado diferente
            lbl_tag = ctk.CTkLabel(
                container,
                text=f"🎯 {len(itens_verdes)} MATERIAIS EM VERDE (COM QUANTIDADES SOMADAS):",
                font=("Segoe UI", 11, "bold"),
                text_color="#34d399"
            )
            lbl_tag.pack(anchor="w", padx=16, pady=(0, 6))

            f_scroll = ctk.CTkScrollableFrame(container, fg_color="transparent")
            f_scroll.pack(fill="both", expand=True, padx=12, pady=(0, 10))

            for it in itens_verdes:
                card = ctk.CTkFrame(f_scroll, fg_color="#064e3b", corner_radius=8, border_width=1, border_color="#059669")
                card.pack(fill="x", pady=4, padx=2)

                f_topo_card = ctk.CTkFrame(card, fg_color="transparent")
                f_topo_card.pack(fill="x", padx=12, pady=(8, 2))

                lbl_m = ctk.CTkLabel(f_topo_card, text=f"Material {it['material']}", font=("Consolas", 13, "bold"), text_color="#a7f3d0")
                lbl_m.pack(side="left")

                lbl_q = ctk.CTkLabel(f_topo_card, text=f"{it['qtd_formatada']} {it['unidade']}", font=("Segoe UI", 16, "bold"), text_color="#34d399")
                lbl_q.pack(side="right")

                lbl_d = ctk.CTkLabel(card, text=it['descricao'], font=("Segoe UI", 11), text_color="white", wraplength=450, justify="left")
                lbl_d.pack(anchor="w", padx=12, pady=(0, 4))

                if len(it["lancamentos"]) > 1:
                    detalhe_soma = " + ".join(_formatar_qtd_float(q) for q in it["lancamentos"])
                    lbl_soma = ctk.CTkLabel(
                        card,
                        text=f"∑ Somado de {len(it['lancamentos'])} itens: {detalhe_soma} = {it['qtd_formatada']} {it['unidade']}",
                        font=("Segoe UI", 10, "bold"),
                        text_color="#6ee7b7"
                    )
                    lbl_soma.pack(anchor="w", padx=12, pady=(0, 4))

                f_sub_b = ctk.CTkFrame(card, fg_color="transparent")
                f_sub_b.pack(anchor="w", padx=12, pady=(0, 8))

                lbl_rec_tag = ctk.CTkLabel(f_sub_b, text=f"{it['recebedor']} • Centro {it['centro']} • Dep {it['deposito']}", font=("Segoe UI", 10), text_color="#94a3b8")
                lbl_rec_tag.pack(side="left")

    def _copiar_saldo(self):
        self.clipboard_clear()
        self.clipboard_append(self.texto_copiar)
        self.update()
        self.btn_copiar.configure(text="✓ Saldo Copiado!")
        self.after(2000, lambda: self.btn_copiar.configure(text="📋 Copiar Saldo"))


class AbaConsultarMigo:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando = False
        self.popup_saldo_ref = None

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("🔍 Módulo Consultar MIGO (Clandestinas) pronto.")

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

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
        f_in2.pack(fill="x", padx=15, pady=(2, 4))

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

        # Linha 3: Destaque do Item Considerado (Filtro EMPR 1CL...)
        self.f_in3 = ctk.CTkFrame(self.frame_inputs, fg_color="#0b1e36", corner_radius=6, border_width=1, border_color="#0284c7")
        self.f_in3.pack(fill="x", padx=15, pady=(4, 8))

        lbl_ret_title = ctk.CTkLabel(
            self.f_in3,
            text="🎯 Item Considerado:",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        )
        lbl_ret_title.pack(side="left", padx=(10, 8), pady=5)

        self.entry_retorno = ctk.CTkEntry(
            self.f_in3,
            placeholder_text="Aguardando consulta (Ex: 510388  27 Pçs EMPR 1CL04)...",
            font=("Consolas", 12, "bold"),
            text_color="#4ade80",
            fg_color="#051329",
            border_color="#1e3a5f",
            height=28
        )
        self.entry_retorno.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=5)

        self.btn_copiar_retorno = ctk.CTkButton(
            self.f_in3,
            text="📋 Copiar Retorno",
            width=125,
            height=28,
            font=("Segoe UI", 11, "bold"),
            fg_color="#f57c00",
            hover_color="#e65100",
            text_color="white",
            command=self.copiar_retorno_destaque
        )
        self.btn_copiar_retorno.pack(side="right", padx=(0, 8), pady=5)

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
        colunas = ("equipe", "ordem", "item", "material", "descricao", "qtd", "un", "recebedor", "centro", "deposito")
        self.tree = ttk.Treeview(self.frame_tabela, columns=colunas, show="headings", selectmode="browse")
        
        self.tree.heading("equipe", text="Equipe")
        self.tree.heading("ordem", text="Ordem (IW39)")
        self.tree.heading("item", text="Item")
        self.tree.heading("material", text="Material")
        self.tree.heading("descricao", text="Descrição do Material")
        self.tree.heading("qtd", text="Qtd")
        self.tree.heading("un", text="UN")
        self.tree.heading("recebedor", text="Recebedor (WEMPF)")
        self.tree.heading("centro", text="Centro")
        self.tree.heading("deposito", text="Depósito")

        self.tree.column("equipe", width=65, anchor="center")
        self.tree.column("ordem", width=95, anchor="center")
        self.tree.column("item", width=40, anchor="center")
        self.tree.column("material", width=90, anchor="center")
        self.tree.column("descricao", width=190, anchor="w")
        self.tree.column("qtd", width=50, anchor="center")
        self.tree.column("un", width=40, anchor="center")
        self.tree.column("recebedor", width=115, anchor="center")
        self.tree.column("centro", width=70, anchor="center")
        self.tree.column("deposito", width=70, anchor="center")

        self.tree.tag_configure("destaque", background="#064e3b", foreground="#a7f3d0")

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
        self.tree.tag_configure("destaque", background="#064e3b", foreground="#a7f3d0")

    def _obter_recebedor_esperado(self, equipe):
        """Retorna o padrão do recebedor esperado com base na equipe."""
        eq = (equipe or "").strip().upper()
        import re
        m = re.search(r'(\d+)', eq)
        if m:
            num = m.group(1)
            sufixo = num[-2:].zfill(2)
            return f"EMPR 1CL{sufixo}"
        return "EMPR 1CL04"

    def _eh_item_considerado(self, recebedor, equipe):
        """
        Verifica se o recebedor do item deve ser considerado.
        Regra: deve conter 'EMPR 1CL04' ou o padrão da equipe (ex: 'EMPR 1CL01' para RCL001).
        """
        import re
        rec = (recebedor or "").strip().upper()
        rec_norm = re.sub(r'\s+', ' ', rec)
        esperado = self._obter_recebedor_esperado(equipe).upper()
        esperado_norm = re.sub(r'\s+', ' ', esperado)
        return (esperado_norm in rec_norm) or ("EMPR 1CL04" in rec_norm)

    def _formatar_qtd(self, qtd_str):
        """Formata quantidade removendo decimais zerados (ex: '27,000' -> '27')."""
        import re
        q = (qtd_str or "").strip()
        if not q:
            return "0"
        q = re.sub(r'[,.]0+$', '', q)
        return q

    def copiar_retorno_destaque(self):
        """Copia o texto do retorno considerado para o Clipboard."""
        txt = self.entry_retorno.get().strip()
        if not txt:
            messagebox.showinfo("Aviso", "Nenhum retorno considerado gerado ainda.")
            return
        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(txt)
        show_toast(self.parent_frame, f"Copiado: {txt}", duration=2000, type="success")
        play_sound("done")

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

    def limpar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.lbl_qtd_itens.configure(text="Itens Encontrados na MIGO: 0")
        if hasattr(self, "entry_retorno"):
            self.entry_retorno.delete(0, tk.END)

    def copiar_dados_tabela(self):
        linhas = []
        linhas.append("Equipe\tOrdem\tItem\tMaterial\tDescrição\tQtd\tUN\tRecebedor\tCentro\tDepósito")
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

    def exibir_popup_saldo(self, equipe, ordem, itens):
        """Abre o Pop-up moderno de Saldo Disponível para a equipe RCL."""
        if not itens and not ordem:
            return
        try:
            if self.popup_saldo_ref and self.popup_saldo_ref.winfo_exists():
                self.popup_saldo_ref.destroy()
            top = self.parent_frame.winfo_toplevel()
            self.popup_saldo_ref = JanelaSaldoRCL(
                top, equipe, ordem, itens,
                paletas=self.paletas,
                obter_tema_atual=self.obter_tema_atual,
                callback_abrir_sap=self.abrir_migo_sap_gui
            )
        except Exception as e:
            log_error(f"Erro ao abrir pop-up de saldo RCL: {e}")

    def _worker_consulta_migo(self, equipe, ordem, callback_ao_concluir=None):
        set_sap_ocupado(True)
        try:
            itens = self._sap_consultar_itens_migo(ordem)
            self.parent_frame.after(0, lambda: self._inserir_itens_tabela(equipe, ordem, itens))
            self.parent_frame.after(0, lambda: self.exibir_popup_saldo(equipe, ordem, itens))
            if callback_ao_concluir:
                self.parent_frame.after(0, lambda: callback_ao_concluir(equipe, ordem, itens))
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

    def iniciar_fluxo_completo(self, callback_ao_concluir=None):
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
                self.parent_frame.after(0, lambda: self.exibir_popup_saldo(equipe, ordem, itens))
                if callback_ao_concluir:
                    self.parent_frame.after(0, lambda: callback_ao_concluir(equipe, ordem, itens))
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
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   On Error Resume Next
   Set session = connection.Children(0)
   If Err.Number <> 0 Or session Is Nothing Then
      Err.Clear
      Dim c, s, connCand, sessCand
      For c = 0 To application.Children.Count - 1
         Set connCand = application.Children(c)
         If Not connCand Is Nothing Then
            For s = 0 To connCand.Children.Count - 1
               Set sessCand = connCand.Children(s)
               If Not sessCand Is Nothing Then
                  Set connection = connCand
                  Set session = sessCand
                  Exit For
               End If
            Next
         End If
         If Not session Is Nothing Then Exit For
      Next
   End If
   On Error GoTo 0
End If
If session Is Nothing Then
   WScript.Echo "ERRO:SEM_SESSAO"
   WScript.Quit 1
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
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
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_idx & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R08"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_idx & "/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = ordem
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_idx & "/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/btnMIGO_OK_GO").press
    If Err.Number = 0 Then
        Exit For
    End If
Next

session.findById("wnd[0]").maximize
WScript.Echo "OK"
'''
        rc, out, err = self._executar_script_vbs(vbs, timeout=None)
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

    def _buscar_descricao_local(self, codigo):
        """Busca a descrição oficial do material no banco de dados Excel local."""
        cod = str(codigo).strip()
        if not cod:
            return ""
        try:
            import pandas as pd
            from utils_os import obter_diretorio_executavel, obter_caminho_recurso
            caminho_banco = os.path.join(obter_diretorio_executavel(), "data", "banco_de_dados.xlsx")
            if not os.path.exists(caminho_banco):
                caminho_banco = obter_caminho_recurso("data", "banco_de_dados.xlsx")
            if os.path.exists(caminho_banco):
                df = pd.read_excel(caminho_banco)
                df.columns = df.columns.str.strip()
                if "Material" in df.columns:
                    df["Material"] = df["Material"].astype(str).str.strip()
                    res = df[df["Material"] == cod]
                    for col in ["Texto breve material", "Descricao", "DESCRICAO", "Texto Breve", "Descrição"]:
                        if col in df.columns and not res.empty:
                            val = str(res.iloc[0][col]).strip()
                            if val and val.lower() != "nan":
                                return val
        except Exception:
            pass
        return ""

    def _inserir_itens_tabela(self, equipe, ordem, itens):
        itens_considerados = []
        for item in itens:
            mat = item.get("material", "").strip()
            desc = item.get("descricao", "").strip()
            if not desc and mat:
                desc = self._buscar_descricao_local(mat)
                item["descricao"] = desc

            eh_destaque = self._eh_item_considerado(item.get("recebedor", ""), equipe)
            tags = ("destaque",) if eh_destaque else ()
            if eh_destaque:
                itens_considerados.append(item)

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
                    item.get("recebedor", ""),
                    item.get("centro", ""),
                    item.get("deposito", "")
                ),
                tags=tags
            )
        total = len(self.tree.get_children())
        self.lbl_qtd_itens.configure(text=f"Itens Encontrados na MIGO: {total}")

        if itens_considerados:
            # Agrupa por material somando as quantidades dos itens em verde
            itens_agrupados = {}
            for it in itens_considerados:
                mat = it.get("material", "").strip()
                if not mat:
                    continue
                q = _converter_qtd_para_float(it.get("qtd", 0))
                un_raw = it.get("unidade", "").strip()
                if un_raw.upper() in ["ST", "PC", "PÇ", "UN", "UND"]:
                    un_l = "Pçs"
                elif un_raw.upper() in ["M", "MT", "MTR"]:
                    un_l = "M"
                else:
                    un_l = un_raw or "Pçs"

                if mat not in itens_agrupados:
                    itens_agrupados[mat] = {
                        "material": mat,
                        "qtd_total": q,
                        "unidade": un_l,
                        "recebedor": it.get("recebedor", ""),
                        "itens": [q]
                    }
                else:
                    itens_agrupados[mat]["qtd_total"] += q
                    itens_agrupados[mat]["itens"].append(q)

            partes_ret = []
            for mat, dados in itens_agrupados.items():
                qtd_somada = _formatar_qtd_float(dados["qtd_total"])
                partes_ret.append(f"{mat}  {qtd_somada} {dados['unidade']}  {dados['recebedor']}")

            retorno_str = " | ".join(partes_ret)
            if hasattr(self, "entry_retorno"):
                self.entry_retorno.delete(0, tk.END)
                self.entry_retorno.insert(0, retorno_str)
            self._log_threadsafe(f"🎯 [ITENS CONSIDERADOS SOMADOS]: {retorno_str}")
        else:
            padrao_esperado = self._obter_recebedor_esperado(equipe)
            self._log_threadsafe(f"ℹ️ Nenhum item com recebedor '{padrao_esperado}' encontrado para a ordem {ordem}.")

    # =========================================================================
    # FUNÇÕES DE EXECUÇÃO VIA CSCRIPT / VBSCRIPT NATIVO (100% COMPATÍVEL COM SAP)
    # =========================================================================

    def _executar_script_vbs(self, conteudo, timeout=None):
        import subprocess, tempfile
        fd, caminho = tempfile.mkstemp(prefix="nexus_migo_", suffix=".vbs", text=True)
        try:
            with os.fdopen(fd, "w", encoding="mbcs", errors="ignore") as f:
                f.write(conteudo)
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            
            # Prioriza cscript.exe padrão (System32 / 64-bit), exatamente igual a OSME, Bandeirada e IQ09
            cscripts = ["cscript.exe"]
            cscript_sys32 = r"C:\Windows\System32\cscript.exe"
            cscript_32 = r"C:\Windows\SysWOW64\cscript.exe"
            if os.path.exists(cscript_sys32) and cscript_sys32 not in cscripts:
                cscripts.append(cscript_sys32)
            if os.path.exists(cscript_32) and cscript_32 not in cscripts:
                cscripts.append(cscript_32)

            last_rc, last_out, last_err = -1, "", ""
            for cs in cscripts:
                try:
                    res = subprocess.run(
                        [cs, "//nologo", caminho],
                        capture_output=True, text=True, timeout=timeout, creationflags=creationflags
                    )
                    last_rc, last_out, last_err = res.returncode, res.stdout or "", res.stderr or ""
                    # Se executou com sucesso (código 0), retorna de imediato
                    if res.returncode == 0:
                        return last_rc, last_out, last_err
                    # Se falhou por causa de incompatibilidade de bitness / ROT (GetObject), tenta o próximo executável
                    if any(termo in (last_out + last_err) for termo in ["800401E4", "Sintaxe inválida", "ActiveX", "GetObject", "ERRO:SAP_DESCONECTADO"]):
                        continue
                    # Outro erro (ex: lógica da tela do SAP), retorna
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
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   On Error Resume Next
   Set session = connection.Children(0)
   If Err.Number <> 0 Or session Is Nothing Then
      Err.Clear
      Dim c, s, connCand, sessCand
      For c = 0 To application.Children.Count - 1
         Set connCand = application.Children(c)
         If Not connCand Is Nothing Then
            For s = 0 To connCand.Children.Count - 1
               Set sessCand = connCand.Children(s)
               If Not sessCand Is Nothing Then
                  Set connection = connCand
                  Set session = sessCand
                  Exit For
               End If
            Next
         End If
         If Not session Is Nothing Then Exit For
      Next
   End If
   On Error GoTo 0
End If
If session Is Nothing Then
   WScript.Echo "ERRO:SEM_SESSAO"
   WScript.Quit 1
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 400

' 2. Preenche Parceiro Z2 e codigo da equipe
session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = LCase(recurso)
session.findById("wnd[0]/usr/ctxtDY_PARNR").setFocus
session.findById("wnd[0]/usr/ctxtDY_PARNR").caretPosition = Len(recurso)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 200

' 3. Executa F8
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 700

' Fecha popup informativo se surgir
If session.Children.Count > 1 Then
    session.findById("wnd[1]").sendVKey 0
    WScript.Sleep 200
End If

Dim ordemEncontrada
ordemEncontrada = ""

' Caso A: Detalhe individual direto (conforme gravado pelo usuario)
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

' Caso B: Grade ALV (se retornar lista de ordens)
If ordemEncontrada = "" Then
    Dim grid
    Set grid = Nothing
    On Error Resume Next
    Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    If Not grid Is Nothing Then
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
        rc, out, err = self._executar_script_vbs(vbs, timeout=None)
        for linha in out.splitlines():
            linha = linha.strip()
            if linha.startswith("ORDEM_ENCONTRADA:"):
                return linha.split(":", 1)[1].strip()
            elif "ERRO:SAP_DESCONECTADO" in linha or "ERRO:SAP_FECHADO" in linha:
                raise RuntimeError("O SAP GUI não está aberto ou não possui conexão ativa com o sistema. Conecte ao ambiente (ex: Produção) antes de consultar.")
            elif "ERRO:SEM_SESSAO" in linha or "ERRO:SEM_CONEXAO" in linha:
                raise RuntimeError("Nenhuma sessão ativa aberta no SAP. Conecte ao ambiente no SAP Logon.")
            elif "ERRO:SCRIPTING_DESABILITADO" in linha or "ERRO:SCRIPTING_ENGINE" in linha:
                raise RuntimeError("Scripting do SAP GUI está desabilitado. Habilite nas Opções do SAP GUI (Acessibilidade e Scripting).")
            elif "ERRO:" in linha:
                raise RuntimeError(linha)
        if rc != 0:
            erro_det = (err or out or "Falha ao executar script SAP").strip()
            if "GetObject" in erro_det or "ActiveX" in erro_det:
                raise RuntimeError("O SAP GUI não está conectado. Abra o SAP Logon e conecte em uma sessão.")
            raise RuntimeError(erro_det)
        return ""

    def _sap_consultar_itens_migo(self, ordem):
        vbs = f'''
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   On Error Resume Next
   Set session = connection.Children(0)
   If Err.Number <> 0 Or session Is Nothing Then
      Err.Clear
      Dim c, s, connCand, sessCand
      For c = 0 To application.Children.Count - 1
         Set connCand = application.Children(c)
         If Not connCand Is Nothing Then
            For s = 0 To connCand.Children.Count - 1
               Set sessCand = connCand.Children(s)
               If Not sessCand Is Nothing Then
                  Set connection = connCand
                  Set session = sessCand
                  Exit For
               End If
            Next
         End If
         If Not session Is Nothing Then Exit For
      Next
   End If
   On Error GoTo 0
End If
If session Is Nothing Then
   WScript.Echo "ERRO:SEM_SESSAO"
   WScript.Quit 1
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
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

' Preenche R08 (Ordem) e clica em Executar
Dim subs, sub_item
subs = Array("0003", "0002")
For Each sub_item In subs
    Err.Clear
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R08"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = ordem
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & sub_item & "/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/btnMIGO_OK_GO").press
    If Err.Number = 0 Then
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
    Dim totalLinhas, r, maxRows, visRows
    totalLinhas = tbl.RowCount
    If totalLinhas > 100 Then maxRows = 100 Else maxRows = totalLinhas
    If maxRows = 0 Then maxRows = tbl.VisibleRowCount
    visRows = tbl.VisibleRowCount
    If visRows <= 0 Then visRows = 10

    Dim colIdxMat, colIdxDesc, colIdxQtd, colIdxUn, colIdxDep, colIdxCentro, colIdxWempf
    colIdxMat = 1
    colIdxDesc = 2
    colIdxQtd = 4
    colIdxUn = 5
    colIdxDep = 6
    colIdxCentro = 7
    colIdxWempf = 11

    On Error Resume Next
    Dim cIdx, cObj, cName
    For cIdx = 0 To tbl.Columns.Count - 1
        Set cObj = tbl.Columns.Item(cIdx)
        If Not cObj Is Nothing Then
            cName = UCase(cObj.Name)
            If InStr(cName, "MATNR") > 0 Then colIdxMat = cIdx
            If InStr(cName, "MAKTX") > 0 Then colIdxDesc = cIdx
            If InStr(cName, "ERFMG") > 0 Then colIdxQtd = cIdx
            If InStr(cName, "ERFME") > 0 Then colIdxUn = cIdx
            If InStr(cName, "LGORT") > 0 Then colIdxDep = cIdx
            If InStr(cName, "WERKS") > 0 Then colIdxCentro = cIdx
            If InStr(cName, "WEMPF") > 0 Then colIdxWempf = cIdx
        End If
    Next
    On Error GoTo 0

    For r = 0 To maxRows - 1
        On Error Resume Next
        Dim visRow
        visRow = r
        If r >= visRows Then
            tbl.VerticalScrollbar.Position = r
            WScript.Sleep 60
            visRow = r - tbl.VerticalScrollbar.Position
            If visRow < 0 Or visRow >= visRows Then visRow = 0
        End If

        Dim mat, desc, qtd, un, centro, dep, wempf
        mat = ""
        desc = ""
        qtd = ""
        un = ""
        centro = ""
        dep = ""
        wempf = ""

        ' --- QUANTIDADE (ERFMG) ---
        qtd = Trim(tbl.GetCellValue(r, "GOITEM-ERFMG"))
        If qtd = "" Then qtd = Trim(tbl.GetCellValue(r, "ERFMG"))
        If qtd = "" Then qtd = Trim(tbl.GetCellValue(r, colIdxQtd))
        If qtd = "" Then qtd = Trim(session.findById(tbl.Id & "/txtGOITEM-ERFMG[" & colIdxQtd & "," & visRow & "]").text)
        If qtd = "" Then qtd = Trim(session.findById(tbl.Id & "/txtGOITEM-ERFMG[4," & visRow & "]").text)
        If qtd = "" Then qtd = Trim(session.findById(tbl.Id & "/ctxtGOITEM-ERFMG[" & colIdxQtd & "," & visRow & "]").text)
        If qtd = "" Then qtd = Trim(session.findById(tbl.Id & "/ctxtGOITEM-ERFMG[4," & visRow & "]").text)
        If qtd = "" Then qtd = Trim(session.findById(tbl.Id & "/txtGOITEM-ERFMG[" & colIdxQtd & "," & r & "]").text)
        If qtd = "" Then qtd = Trim(session.findById(tbl.Id & "/txtGOITEM-ERFMG[4," & r & "]").text)

        ' --- UNIDADE (ERFME) ---
        un = Trim(tbl.GetCellValue(r, "GOITEM-ERFME"))
        If un = "" Then un = Trim(tbl.GetCellValue(r, "ERFME"))
        If un = "" Then un = Trim(tbl.GetCellValue(r, colIdxUn))
        If un = "" Then un = Trim(session.findById(tbl.Id & "/ctxtGOITEM-ERFME[" & colIdxUn & "," & visRow & "]").text)
        If un = "" Then un = Trim(session.findById(tbl.Id & "/txtGOITEM-ERFME[" & colIdxUn & "," & visRow & "]").text)
        If un = "" Then un = Trim(session.findById(tbl.Id & "/ctxtGOITEM-ERFME[5," & visRow & "]").text)
        If un = "" Then un = Trim(session.findById(tbl.Id & "/txtGOITEM-ERFME[5," & visRow & "]").text)
        If un = "" Then un = Trim(session.findById(tbl.Id & "/ctxtGOITEM-ERFME[" & colIdxUn & "," & r & "]").text)
        If un = "" Then un = Trim(session.findById(tbl.Id & "/ctxtGOITEM-ERFME[5," & r & "]").text)

        ' --- DESCRIÇÃO (MAKTX) ---
        desc = Trim(tbl.GetCellValue(r, "GOITEM-MAKTX"))
        If desc = "" Then desc = Trim(tbl.GetCellValue(r, "MAKTX"))
        If desc = "" Then desc = Trim(tbl.GetCellValue(r, colIdxDesc))
        If desc = "" Then desc = Trim(session.findById(tbl.Id & "/txtGOITEM-MAKTX[" & colIdxDesc & "," & visRow & "]").text)
        If desc = "" Then desc = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MAKTX[" & colIdxDesc & "," & visRow & "]").text)
        If desc = "" Then desc = Trim(session.findById(tbl.Id & "/txtGOITEM-MAKTX[2," & visRow & "]").text)
        If desc = "" Then desc = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MAKTX[2," & visRow & "]").text)
        If desc = "" Then desc = Trim(session.findById(tbl.Id & "/txtGOITEM-MAKTX[" & colIdxDesc & "," & r & "]").text)
        If desc = "" Then desc = Trim(session.findById(tbl.Id & "/txtGOITEM-MAKTX[2," & r & "]").text)

        ' --- MATERIAL (MATNR) ---
        mat = Trim(tbl.GetCellValue(r, "GOITEM-MATNR"))
        If mat = "" Then mat = Trim(tbl.GetCellValue(r, "MATNR"))
        If mat = "" Then mat = Trim(tbl.GetCellValue(r, colIdxMat))
        If mat = "" Then mat = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MATNR[" & colIdxMat & "," & visRow & "]").text)
        If mat = "" Then mat = Trim(session.findById(tbl.Id & "/txtGOITEM-MATNR[" & colIdxMat & "," & visRow & "]").text)
        If mat = "" Then mat = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MATNR[1," & visRow & "]").text)
        If mat = "" Then mat = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MATNR[" & colIdxMat & "," & r & "]").text)
        If mat = "" Then mat = Trim(session.findById(tbl.Id & "/ctxtGOITEM-MATNR[1," & r & "]").text)

        ' --- RECEBEDOR (WEMPF) ---
        wempf = Trim(tbl.GetCellValue(r, "GOITEM-WEMPF"))
        If wempf = "" Then wempf = Trim(tbl.GetCellValue(r, "WEMPF"))
        If wempf = "" Then wempf = Trim(tbl.GetCellValue(r, colIdxWempf))
        If wempf = "" Then wempf = Trim(session.findById(tbl.Id & "/txtGOITEM-WEMPF[" & colIdxWempf & "," & visRow & "]").text)
        If wempf = "" Then wempf = Trim(session.findById(tbl.Id & "/ctxtGOITEM-WEMPF[" & colIdxWempf & "," & visRow & "]").text)
        If wempf = "" Then wempf = Trim(session.findById(tbl.Id & "/txtGOITEM-WEMPF[11," & visRow & "]").text)
        If wempf = "" Then wempf = Trim(session.findById(tbl.Id & "/txtGOITEM-WEMPF[10," & visRow & "]").text)
        If wempf = "" Then wempf = Trim(session.findById(tbl.Id & "/txtGOITEM-WEMPF[" & colIdxWempf & "," & r & "]").text)
        If wempf = "" Then wempf = Trim(session.findById(tbl.Id & "/txtGOITEM-WEMPF[11," & r & "]").text)

        ' --- CENTRO (WERKS) ---
        centro = Trim(tbl.GetCellValue(r, "GOITEM-WERKS"))
        If centro = "" Then centro = Trim(tbl.GetCellValue(r, "WERKS"))
        If centro = "" Then centro = Trim(tbl.GetCellValue(r, "NAME1"))
        If centro = "" Then centro = Trim(tbl.GetCellValue(r, colIdxCentro))
        If centro = "" Then centro = Trim(session.findById(tbl.Id & "/ctxtGOITEM-WERKS[" & colIdxCentro & "," & visRow & "]").text)
        If centro = "" Then centro = Trim(session.findById(tbl.Id & "/txtGOITEM-WERKS[" & colIdxCentro & "," & visRow & "]").text)
        If centro = "" Then centro = Trim(session.findById(tbl.Id & "/ctxtGOITEM-WERKS[7," & visRow & "]").text)
        If centro = "" Then centro = "EMPR"

        ' --- DEPÓSITO (LGORT) ---
        dep = Trim(tbl.GetCellValue(r, "GOITEM-LGORT"))
        If dep = "" Then dep = Trim(tbl.GetCellValue(r, "LGORT"))
        If dep = "" Then dep = Trim(tbl.GetCellValue(r, "LGOBE"))
        If dep = "" Then dep = Trim(tbl.GetCellValue(r, colIdxDep))
        If dep = "" Then dep = Trim(session.findById(tbl.Id & "/ctxtGOITEM-LGORT[" & colIdxDep & "," & visRow & "]").text)
        If dep = "" Then dep = Trim(session.findById(tbl.Id & "/txtGOITEM-LGORT[" & colIdxDep & "," & visRow & "]").text)
        If dep = "" Then dep = Trim(session.findById(tbl.Id & "/ctxtGOITEM-LGORT[6," & visRow & "]").text)
        If dep = "" Then dep = Trim(session.findById(tbl.Id & "/ctxtGOITEM-LGORT[5," & visRow & "]").text)
        If dep = "" Then dep = Trim(session.findById(tbl.Id & "/ctxtGOITEM-LGORT[" & colIdxDep & "," & r & "]").text)
        If dep = "" Then dep = Trim(session.findById(tbl.Id & "/ctxtGOITEM-LGORT[6," & r & "]").text)

        If mat <> "" Or desc <> "" Or qtd <> "" Or wempf <> "" Then
            WScript.Echo "ITEM:" & (r + 1) & "|" & mat & "|" & desc & "|" & qtd & "|" & un & "|" & wempf & "|" & centro & "|" & dep
        End If
    Next
Else
    ' Caso de item unico detalhado
    On Error Resume Next
    Dim txtMat
    Set txtMat = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX")
    If Not txtMat Is Nothing Then
        WScript.Echo "ITEM:1|Item Ordem|" & Trim(txtMat.text) & "|1|UN|EMPR 1CL04|EMPR|CL03"
    End If
End If

' Sai da MIGO tratando popup de confirmacao sem salvar
session.findById("wnd[0]/tbar[0]/btn[15]").press
WScript.Sleep 250
If session.Children.Count > 1 Then
    On Error Resume Next
    session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[1]").sendVKey 0
    End If
    On Error GoTo 0
End If

WScript.Echo "FIM_MIGO"
'''
        rc, out, err = self._executar_script_vbs(vbs, timeout=None)
        itens = []
        for linha in out.splitlines():
            linha = linha.strip()
            if linha.startswith("ITEM:"):
                partes = linha.split(":", 1)[1].split("|")
                if len(partes) >= 8:
                    itens.append({
                        "item": partes[0],
                        "material": partes[1],
                        "descricao": partes[2],
                        "qtd": partes[3],
                        "unidade": partes[4],
                        "recebedor": partes[5],
                        "centro": partes[6],
                        "deposito": partes[7]
                    })
                elif len(partes) >= 7:
                    itens.append({
                        "item": partes[0],
                        "material": partes[1],
                        "descricao": partes[2],
                        "qtd": partes[3],
                        "unidade": partes[4],
                        "recebedor": "",
                        "centro": partes[5],
                        "deposito": partes[6]
                    })
            elif "ERRO:SAP_DESCONECTADO" in linha or "ERRO:SAP_FECHADO" in linha:
                raise RuntimeError("O SAP GUI não está aberto ou não possui conexão ativa com o sistema. Conecte ao ambiente (ex: Produção) antes de consultar.")
            elif "ERRO:SEM_SESSAO" in linha or "ERRO:SEM_CONEXAO" in linha:
                raise RuntimeError("Nenhuma sessão ativa aberta no SAP. Conecte ao ambiente no SAP Logon.")
            elif "ERRO:SCRIPTING_DESABILITADO" in linha or "ERRO:SCRIPTING_ENGINE" in linha:
                raise RuntimeError("Scripting do SAP GUI está desabilitado. Habilite nas Opções do SAP GUI (Acessibilidade e Scripting).")
            elif "ERRO:" in linha:
                raise RuntimeError(linha)
        if rc != 0 and not itens:
            erro_det = (err or out or "Falha ao executar script SAP").strip()
            if "GetObject" in erro_det or "ActiveX" in erro_det:
                raise RuntimeError("O SAP GUI não está conectado. Abra o SAP Logon e conecte em uma sessão.")
            raise RuntimeError(erro_det)
        return itens


def construir_aba_consultar_migo(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaConsultarMigo(parent_frame, label_status_global, paletas, obter_tema_atual)
