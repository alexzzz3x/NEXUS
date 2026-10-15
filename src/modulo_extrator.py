import os
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox
from collections import Counter
import customtkinter as ctk
from modulo_Sap import verificar_sap_aberto
from utils_metrics import incrementar_metrica
from utils_os import obter_caminho_script, obter_diretorio_executavel, obter_caminho_recurso

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

def buscar_descricao_material_local(codigo):
    """Busca a descrição do material no banco de dados Excel local."""
    cod = str(codigo).strip()
    if not cod:
        return ""
    try:
        import pandas as pd
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

class JanelaDetalheMedidor(ctk.CTkToplevel):
    """Pop-up moderno para exibir detalhes capturados de um medidor na IQ09."""
    def __init__(self, parent, dados, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.dados = dados or {}
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")

        self.title("⚡ Detalhes do Medidor — SAP IQ09")
        self.geometry("520x570")
        self.minsize(480, 520)
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
            w, h = 520, 570
            x = max(50, px + (pw - w) // 2)
            y = max(50, py + (ph - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry("520x570+200+150")

    def _construir_interface(self):
        c = self.paletas.get(self.obter_tema_atual(), self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))
        self.configure(fg_color=c.get("bg", "#0a192f"))

        container = ctk.CTkFrame(self, fg_color=c.get("container", "#172a45"), corner_radius=12, border_width=1, border_color="#334155")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        # 1. Cabeçalho
        f_header = ctk.CTkFrame(container, fg_color="transparent")
        f_header.pack(fill="x", padx=16, pady=(16, 10))

        lbl_icone = ctk.CTkLabel(f_header, text="⚡", font=("Segoe UI", 26))
        lbl_icone.pack(side="left", padx=(0, 10))

        f_titulos = ctk.CTkFrame(f_header, fg_color="transparent")
        f_titulos.pack(side="left", fill="both", expand=True)

        lbl_titulo = ctk.CTkLabel(f_titulos, text="DETALHES DO MEDIDOR", font=("Segoe UI", 16, "bold"), text_color=c.get("texto", "white"))
        lbl_titulo.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(f_titulos, text="Informações extraídas da transação IQ09 no SAP GUI", font=("Segoe UI", 11), text_color="#94a3b8")
        lbl_sub.pack(anchor="w")

        # 2. Badge de Status (ex: LIDI / DISP ou LIDI / ELAB)
        raw_status = self.dados.get("status", "").strip() or "NÃO INFORMADO"
        # Normalizar separador se veio separado por espaço
        if "/" not in raw_status and " " in raw_status:
            partes = [p.strip() for p in raw_status.split() if p.strip()]
            status_txt = " / ".join(partes)
        else:
            status_txt = raw_status

        st_upper = status_txt.upper()
        eh_disp = "DISP" in st_upper
        eh_elab = "ELAB" in st_upper
        eh_alerta = eh_elab or ("BLOQ" in st_upper) or (not eh_disp and raw_status != "NÃO INFORMADO")

        if eh_disp and not eh_elab:
            disp_bg = "#064e3b"
            disp_border = "#059669"
            disp_fg = "#34d399"
            st_prefix = "✓ "
            aviso_txt = "Disponível no SAP"
        elif eh_alerta:
            disp_bg = "#451a03"
            disp_border = "#f59e0b"
            disp_fg = "#fef08a"
            st_prefix = "⚠️ "
            aviso_txt = "Atenção: Não disponível (ELAB/Bloqueado)"
        else:
            disp_bg = "#1e3a8a"
            disp_border = "#2563eb"
            disp_fg = "#93c5fd"
            st_prefix = "ℹ️ "
            aviso_txt = ""

        f_status = ctk.CTkFrame(container, fg_color=disp_bg, corner_radius=8, border_width=1, border_color=disp_border)
        f_status.pack(fill="x", padx=16, pady=(0, 12))

        lbl_st_tag = ctk.CTkLabel(f_status, text="STATUS:", font=("Segoe UI", 11, "bold"), text_color="#cbd5e1")
        lbl_st_tag.pack(side="left", padx=(14, 6), pady=8)

        lbl_st_val = ctk.CTkLabel(f_status, text=f"{st_prefix}{status_txt}", font=("Segoe UI", 14, "bold"), text_color=disp_fg)
        lbl_st_val.pack(side="left", pady=8)

        if aviso_txt:
            lbl_st_aviso = ctk.CTkLabel(f_status, text=aviso_txt, font=("Segoe UI", 10, "italic"), text_color=disp_fg)
            lbl_st_aviso.pack(side="right", padx=(0, 14), pady=8)

        # 3. Card do Número de Série / Medidor
        f_serial = ctk.CTkFrame(container, fg_color=c.get("input", "#1f3a60"), corner_radius=8, border_width=1, border_color="#334155")
        f_serial.pack(fill="x", padx=16, pady=(0, 10))

        lbl_ser_tit = ctk.CTkLabel(f_serial, text="🔢 NÚMERO DE SÉRIE / MEDIDOR", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_ser_tit.pack(anchor="w", padx=14, pady=(8, 2))

        serial_val = self.dados.get("serial", "") or "—"
        lbl_ser_val = ctk.CTkLabel(f_serial, text=serial_val, font=("Consolas", 20, "bold"), text_color=c.get("accent", "#f57c00"))
        lbl_ser_val.pack(anchor="w", padx=14, pady=(0, 10))

        # 4. Card Denominação
        f_desc = ctk.CTkFrame(container, fg_color=c.get("input", "#1f3a60"), corner_radius=8, border_width=1, border_color="#334155")
        f_desc.pack(fill="x", padx=16, pady=(0, 10))

        lbl_desc_tit = ctk.CTkLabel(f_desc, text="📝 DENOMINAÇÃO DO EQUIPAMENTO", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_desc_tit.pack(anchor="w", padx=14, pady=(8, 2))

        desc_val = self.dados.get("denominacao", "") or "—"
        lbl_desc_val = ctk.CTkLabel(f_desc, text=desc_val, font=("Segoe UI", 12, "bold"), text_color="white", wraplength=450, justify="left")
        lbl_desc_val.pack(anchor="w", padx=14, pady=(0, 10))

        # 5. Grid de Informações Técnicas (Material, Equipamento, Centro, Depósito)
        f_grid = ctk.CTkFrame(container, fg_color="transparent")
        f_grid.pack(fill="x", padx=16, pady=(0, 10))
        f_grid.columnconfigure(0, weight=1)
        f_grid.columnconfigure(1, weight=1)

        # Card Material & Equipamento (Esquerda)
        f_col_esq = ctk.CTkFrame(f_grid, fg_color=c.get("input", "#1f3a60"), corner_radius=8, border_width=1, border_color="#334155")
        f_col_esq.grid(row=0, column=0, padx=(0, 6), sticky="nsew")

        lbl_m1 = ctk.CTkLabel(f_col_esq, text="⚙️ Material:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_m1.pack(anchor="w", padx=10, pady=(8, 1))
        lbl_m2 = ctk.CTkLabel(f_col_esq, text=self.dados.get("material", "—") or "—", font=("Consolas", 12, "bold"), text_color="white")
        lbl_m2.pack(anchor="w", padx=10, pady=(0, 6))

        lbl_e1 = ctk.CTkLabel(f_col_esq, text="🆔 Equipamento SAP:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_e1.pack(anchor="w", padx=10, pady=(0, 1))
        lbl_e2 = ctk.CTkLabel(f_col_esq, text=self.dados.get("equipamento", "—") or "—", font=("Consolas", 12, "bold"), text_color="#38bdf8")
        lbl_e2.pack(anchor="w", padx=10, pady=(0, 8))

        # Card Centro & Depósito (Direita)
        f_col_dir = ctk.CTkFrame(f_grid, fg_color=c.get("input", "#1f3a60"), corner_radius=8, border_width=1, border_color="#334155")
        f_col_dir.grid(row=0, column=1, padx=(6, 0), sticky="nsew")

        val_centro = str(self.dados.get("centro", "") or "").strip()
        if not val_centro or val_centro in ["—", "-", "None", "Não informado"]:
            val_centro = "Sem informação"
            cor_centro = "#94a3b8"
        else:
            cor_centro = "white"

        val_dep = str(self.dados.get("deposito", "") or "").strip()
        if not val_dep or val_dep in ["—", "-", "None", "Não informado"]:
            val_dep = "Sem informação"
            cor_dep = "#94a3b8"
        else:
            cor_dep = "#fde047"

        lbl_c1 = ctk.CTkLabel(f_col_dir, text="🏢 Centro:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_c1.pack(anchor="w", padx=10, pady=(8, 1))
        lbl_c2 = ctk.CTkLabel(f_col_dir, text=val_centro, font=("Segoe UI", 12, "bold"), text_color=cor_centro)
        lbl_c2.pack(anchor="w", padx=10, pady=(0, 6))

        lbl_d1 = ctk.CTkLabel(f_col_dir, text="📦 Depósito:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_d1.pack(anchor="w", padx=10, pady=(0, 1))
        lbl_d2 = ctk.CTkLabel(f_col_dir, text=val_dep, font=("Segoe UI", 12, "bold"), text_color=cor_dep)
        lbl_d2.pack(anchor="w", padx=10, pady=(0, 8))

        # Tipo de estoque se houver
        tipo_estq = self.dados.get("tipo_estoque", "").strip()
        if tipo_estq:
            lbl_tipo = ctk.CTkLabel(container, text=f"🏷️ Tipo de Estoque: {tipo_estq}", font=("Segoe UI", 10), text_color="#94a3b8")
            lbl_tipo.pack(anchor="w", padx=20, pady=(0, 8))

        # 6. Botões de Ação
        f_acoes = ctk.CTkFrame(container, fg_color="transparent")
        f_acoes.pack(fill="x", padx=16, pady=(10, 14), side="bottom")

        self.btn_copiar = ctk.CTkButton(
            f_acoes,
            text="📋 Copiar Informações",
            height=36,
            font=("Segoe UI", 12, "bold"),
            fg_color=c.get("accent", "#f57c00"),
            hover_color=c.get("hover", "#e65100"),
            command=self._copiar_dados
        )
        self.btn_copiar.pack(side="left", fill="x", expand=True, padx=(0, 6))

        btn_fechar = ctk.CTkButton(
            f_acoes,
            text="❌ Fechar",
            height=36,
            width=100,
            font=("Segoe UI", 12),
            fg_color="#334155",
            hover_color="#475569",
            command=self.destroy
        )
        btn_fechar.pack(side="right", padx=(6, 0))

    def _copiar_dados(self):
        val_centro = str(self.dados.get("centro", "") or "").strip() or "Sem informação"
        val_dep = str(self.dados.get("deposito", "") or "").strip() or "Sem informação"
        resumo = (
            f"⚡ DETALHES DO MEDIDOR (SAP IQ09)\n"
            f"• Medidor/Serial: {self.dados.get('serial', '—')}\n"
            f"• Status: {self.dados.get('status', '—')}\n"
            f"• Denominação: {self.dados.get('denominacao', '—')}\n"
            f"• Material: {self.dados.get('material', '—')}\n"
            f"• Equipamento: {self.dados.get('equipamento', '—')}\n"
            f"• Centro: {val_centro}\n"
            f"• Depósito: {val_dep}"
        )
        self.clipboard_clear()
        self.clipboard_append(resumo)
        self.update()
        self.btn_copiar.configure(text="✓ Informações Copiadas!")
        self.after(2000, lambda: self.btn_copiar.configure(text="📋 Copiar Informações"))

class JanelaMultiplosMedidores(ctk.CTkToplevel):
    """Pop-up moderno para exibir lista com múltiplos medidores capturados na IQ09."""
    def __init__(self, parent, lista_medidores, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.lista_medidores = lista_medidores or []
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")

        qtd = len(self.lista_medidores)
        self.title(f"⚡ Múltiplos Medidores — SAP IQ09 ({qtd} encontrados)")
        self.geometry("860x600")
        self.minsize(740, 480)
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
            w, h = 860, 600
            x = max(50, px + (pw - w) // 2)
            y = max(50, py + (ph - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry("860x600+180+120")

    def _construir_interface(self):
        c = self.paletas.get(self.obter_tema_atual(), self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))
        self.configure(fg_color=c.get("bg", "#0a192f"))

        container = ctk.CTkFrame(self, fg_color=c.get("container", "#172a45"), corner_radius=12, border_width=1, border_color="#334155")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        # 1. Cabeçalho
        f_header = ctk.CTkFrame(container, fg_color="transparent")
        f_header.pack(fill="x", padx=16, pady=(14, 10))

        lbl_icone = ctk.CTkLabel(f_header, text="⚡", font=("Segoe UI", 26))
        lbl_icone.pack(side="left", padx=(0, 10))

        f_titulos = ctk.CTkFrame(f_header, fg_color="transparent")
        f_titulos.pack(side="left", fill="both", expand=True)

        qtd = len(self.lista_medidores)
        lbl_titulo = ctk.CTkLabel(f_titulos, text="CONSULTA MÚLTIPLA DE MEDIDORES", font=("Segoe UI", 16, "bold"), text_color=c.get("texto", "white"))
        lbl_titulo.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(f_titulos, text=f"Total de {qtd} medidores carregados da transação IQ09 no SAP GUI", font=("Segoe UI", 11), text_color="#94a3b8")
        lbl_sub.pack(anchor="w")

        # 2. Mini-Cards de Estatísticas (Total, DISP, Outros)
        def _eh_disp_valido(st_str):
            u = str(st_str).upper()
            return ("DISP" in u) and ("ELAB" not in u) and ("BLOQ" not in u)

        self.qtd_disp = sum(1 for m in self.lista_medidores if _eh_disp_valido(m.get("status", "")))
        self.qtd_outros = qtd - self.qtd_disp
        self.filtro_atual = "todos"

        f_stats = ctk.CTkFrame(container, fg_color="transparent")
        f_stats.pack(fill="x", padx=16, pady=(0, 10))

        def _criar_card(pai, titulo, valor, cor_val, alerta=False):
            b_color = "#f59e0b" if alerta else "#334155"
            bg_c = "#341805" if alerta else c.get("input", "#1f3a60")
            card = ctk.CTkFrame(pai, fg_color=bg_c, corner_radius=8, border_width=2 if alerta else 1, border_color=b_color)
            card.pack(side="left", fill="x", expand=True, padx=4)
            lbl_t = ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 10, "bold"), text_color="#fde047" if alerta else "#94a3b8")
            lbl_t.pack(anchor="w", padx=10, pady=(6, 1))
            lbl_v = ctk.CTkLabel(card, text=str(valor), font=("Segoe UI", 14, "bold"), text_color=cor_val)
            lbl_v.pack(anchor="w", padx=10, pady=(0, 6))

        _criar_card(f_stats, "TOTAL ENCONTRADOS", f"{qtd} medidores", "white")
        _criar_card(f_stats, "✓ STATUS DISP (DISPONÍVEL)", f"{self.qtd_disp} medidores", "#34d399")
        _criar_card(
            f_stats,
            "⚠️ ATENÇÃO / NÃO DISP" if self.qtd_outros > 0 else "OUTROS STATUS",
            f"{self.qtd_outros} medidor(es)" if self.qtd_outros > 0 else "0",
            "#fbbf24" if self.qtd_outros > 0 else "#94a3b8",
            alerta=(self.qtd_outros > 0)
        )

        # 3. Barra de Filtro Rápido
        if self.qtd_outros > 0:
            f_filtros = ctk.CTkFrame(container, fg_color="transparent")
            f_filtros.pack(fill="x", padx=16, pady=(0, 6))

            self.btn_f_todos = ctk.CTkButton(
                f_filtros,
                text=f"👁️ Ver Todos ({qtd})",
                height=28,
                width=110,
                font=("Segoe UI", 10, "bold"),
                fg_color="#0284c7",
                hover_color="#0369a1",
                command=lambda: self._aplicar_filtro("todos")
            )
            self.btn_f_todos.pack(side="left", padx=(0, 6))

            self.btn_f_atencao = ctk.CTkButton(
                f_filtros,
                text=f"⚠️ Apenas com Atenção ({self.qtd_outros})",
                height=28,
                width=180,
                font=("Segoe UI", 10, "bold"),
                fg_color="#b45309",
                hover_color="#92400e",
                command=lambda: self._aplicar_filtro("atencao")
            )
            self.btn_f_atencao.pack(side="left", padx=(0, 6))

        # 4. Cabeçalho das Colunas
        f_colunas = ctk.CTkFrame(container, fg_color="#1e293b", corner_radius=6, height=32)
        f_colunas.pack(fill="x", padx=16, pady=(0, 6))
        f_colunas.pack_propagate(False)

        colunas = [
            ("SERIAL / MEDIDOR", 130, (12, 0)),
            ("STATUS (SISTEMA / USUÁRIO)", 150, (6, 0)),
            ("MATERIAL", 110, (6, 0)),
            ("CENTRO / DEP", 110, (6, 0)),
            ("DENOMINAÇÃO", 0, (6, 6))
        ]
        for nome, larg, pd in colunas:
            lbl_c = ctk.CTkLabel(f_colunas, text=nome, font=("Segoe UI", 10, "bold"), text_color="#cbd5e1")
            if larg > 0:
                lbl_c.pack(side="left", padx=pd)
                lbl_c.configure(width=larg)
            else:
                lbl_c.pack(side="left", fill="x", expand=True, padx=pd)

        # 5. Barra de Ações Inferior (fixada no rodapé)
        f_acoes = ctk.CTkFrame(container, fg_color="transparent")
        f_acoes.pack(fill="x", padx=16, pady=(4, 12), side="bottom")

        self.btn_copiar_tab = ctk.CTkButton(
            f_acoes,
            text="📋 Copiar Tabela Completa (Excel)",
            height=36,
            font=("Segoe UI", 12, "bold"),
            fg_color=c.get("accent", "#f57c00"),
            hover_color=c.get("hover", "#e65100"),
            command=self._copiar_tabela
        )
        self.btn_copiar_tab.pack(side="left", padx=(0, 8))

        self.btn_copiar_ser = ctk.CTkButton(
            f_acoes,
            text="🔢 Copiar Seriais",
            height=36,
            font=("Segoe UI", 12),
            fg_color="#334155",
            hover_color="#475569",
            command=self._copiar_apenas_seriais
        )
        self.btn_copiar_ser.pack(side="left", padx=4)

        btn_fechar = ctk.CTkButton(
            f_acoes,
            text="❌ Fechar",
            height=36,
            width=100,
            font=("Segoe UI", 12),
            fg_color="#1e293b",
            hover_color="#334155",
            command=self.destroy
        )
        btn_fechar.pack(side="right", padx=(8, 0))

        # 6. Lista Rolável de Medidores (preenche o centro)
        self.scroll_frame = ctk.CTkScrollableFrame(container, fg_color="transparent", corner_radius=6)
        self.scroll_frame.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        self._renderizar_linhas()

    def _aplicar_filtro(self, tipo):
        self.filtro_atual = tipo
        if hasattr(self, "btn_f_todos") and hasattr(self, "btn_f_atencao"):
            if tipo == "todos":
                self.btn_f_todos.configure(fg_color="#0284c7")
                self.btn_f_atencao.configure(fg_color="#475569")
            else:
                self.btn_f_todos.configure(fg_color="#475569")
                self.btn_f_atencao.configure(fg_color="#b45309")
        self._renderizar_linhas()

    def _renderizar_linhas(self):
        c = self.paletas.get(self.obter_tema_atual(), self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))
        for w in self.scroll_frame.winfo_children():
            w.destroy()

        def _eh_disp(st_str):
            u = str(st_str).upper()
            return ("DISP" in u) and ("ELAB" not in u) and ("BLOQ" not in u)

        itens = self.lista_medidores
        if self.filtro_atual == "atencao":
            itens = [m for m in self.lista_medidores if not _eh_disp(m.get("status", ""))]

        for idx, m in enumerate(itens):
            st_txt = m.get("status", "—") or "—"
            disp = _eh_disp(st_txt)

            if disp:
                bg_row = c.get("input", "#1f3a60") if idx % 2 == 0 else "#162740"
                border_color = "#334155"
                border_w = 1
                badge_bg = "#064e3b"
                badge_fg = "#34d399"
                badge_txt = f"✓ {st_txt}"
            else:
                # Destaque de ATENÇÃO (ex: ELAB, MONT, ESTO, BLOQ)
                bg_row = "#2e1808" if idx % 2 == 0 else "#251206"
                border_color = "#f59e0b"
                border_w = 2
                badge_bg = "#78350f"
                badge_fg = "#fde047"
                badge_txt = f"⚠️ {st_txt}"

            f_row = ctk.CTkFrame(self.scroll_frame, fg_color=bg_row, corner_radius=6, border_width=border_w, border_color=border_color)
            f_row.pack(fill="x", pady=2)

            ser_txt = m.get("serial", "—") or "—"
            mat_txt = m.get("material", "—") or "—"
            centro_txt = m.get("centro", "—") or "—"
            dep_txt = m.get("deposito", "—") or "—"
            desc_txt = m.get("denominacao", "—") or "—"

            # Serial
            lbl_s = ctk.CTkLabel(f_row, text=ser_txt, font=("Consolas", 12, "bold"), text_color=c.get("accent", "#f57c00"), width=130, anchor="w")
            lbl_s.pack(side="left", padx=(12, 6), pady=6)

            # Status Badge Destacado
            f_badge = ctk.CTkFrame(f_row, fg_color=badge_bg, corner_radius=4)
            f_badge.pack(side="left", padx=6, pady=4)
            lbl_st = ctk.CTkLabel(f_badge, text=f" {badge_txt} ", font=("Segoe UI", 10, "bold"), text_color=badge_fg, width=138, anchor="center")
            lbl_st.pack(padx=4, pady=2)

            # Material
            lbl_m = ctk.CTkLabel(f_row, text=mat_txt, font=("Consolas", 11), text_color="white", width=110, anchor="w")
            lbl_m.pack(side="left", padx=6, pady=6)

            # Centro / Depósito
            c_dep = f"{centro_txt} / {dep_txt}" if centro_txt != "Sem informação" or dep_txt != "Sem informação" else "Sem info"
            lbl_cd = ctk.CTkLabel(f_row, text=c_dep, font=("Segoe UI", 11), text_color="#fde047" if "Sem" not in c_dep else "#94a3b8", width=110, anchor="w")
            lbl_cd.pack(side="left", padx=6, pady=6)

            # Denominação
            lbl_d = ctk.CTkLabel(f_row, text=desc_txt, font=("Segoe UI", 11), text_color="#e2e8f0", anchor="w")
            lbl_d.pack(side="left", fill="x", expand=True, padx=(6, 12), pady=6)

    def _copiar_tabela(self):
        linhas = ["SERIAL\tSTATUS\tMATERIAL\tCENTRO\tDEPÓSITO\tDENOMINAÇÃO"]
        for m in self.lista_medidores:
            linhas.append(
                f"{m.get('serial', '')}\t{m.get('status', '')}\t{m.get('material', '')}\t"
                f"{m.get('centro', '')}\t{m.get('deposito', '')}\t{m.get('denominacao', '')}"
            )
        texto = "\r\n".join(linhas)
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(texto)
            win32clipboard.CloseClipboard()
        except Exception:
            self.clipboard_clear()
            self.clipboard_append(texto)
            self.update()

        txt_antigo = self.btn_copiar_tab.cget("text")
        self.btn_copiar_tab.configure(text="✅ Tabela Copiada!")
        self.after(2000, lambda: self.btn_copiar_tab.configure(text=txt_antigo))

    def _copiar_apenas_seriais(self):
        seriais = [str(m.get("serial", "")).strip() for m in self.lista_medidores if m.get("serial")]
        texto = "\r\n".join(seriais)
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(texto)
            win32clipboard.CloseClipboard()
        except Exception:
            self.clipboard_clear()
            self.clipboard_append(texto)
            self.update()

        txt_antigo = self.btn_copiar_ser.cget("text")
        self.btn_copiar_ser.configure(text="✅ Seriais Copiados!")
        self.after(2000, lambda: self.btn_copiar_ser.configure(text=txt_antigo))


class JanelaSeriaisDuplicados(ctk.CTkToplevel):
    """Pop-up moderno para exibir seriais duplicados encontrados durante a extração."""
    def __init__(self, parent, dados_analise, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.dados_analise = dados_analise or {}
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")

        duplicados = self.dados_analise.get("duplicados", {})
        total_dup = self.dados_analise.get("total_duplicados", 0)
        qtd_seriais_dup = len(duplicados)

        self.title(f"⚠️ Seriais Duplicados — Detector ({qtd_seriais_dup} seriais, {total_dup} repetições)")
        self.geometry("620x520")
        self.minsize(520, 420)
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
            w, h = 620, 520
            x = max(50, px + (pw - w) // 2)
            y = max(50, py + (ph - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry("620x520+200+150")

    def _construir_interface(self):
        c = self.paletas.get(self.obter_tema_atual(), self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))
        self.configure(fg_color=c.get("bg", "#0a192f"))

        container = ctk.CTkFrame(self, fg_color=c.get("container", "#172a45"), corner_radius=12, border_width=1, border_color="#334155")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        # 1. Cabeçalho
        f_header = ctk.CTkFrame(container, fg_color="transparent")
        f_header.pack(fill="x", padx=16, pady=(14, 10))

        lbl_icone = ctk.CTkLabel(f_header, text="⚠️", font=("Segoe UI", 26))
        lbl_icone.pack(side="left", padx=(0, 10))

        f_titulos = ctk.CTkFrame(f_header, fg_color="transparent")
        f_titulos.pack(side="left", fill="both", expand=True)

        lbl_titulo = ctk.CTkLabel(f_titulos, text="SERIAIS DUPLICADOS DETECTADOS", font=("Segoe UI", 16, "bold"), text_color="#f59e0b")
        lbl_titulo.pack(anchor="w")

        lbl_sub = ctk.CTkLabel(f_titulos, text="Estes seriais apareceram repetidos no texto e foram consolidados na lista única.", font=("Segoe UI", 11), text_color="#94a3b8")
        lbl_sub.pack(anchor="w")

        # 2. Mini-Cards de Estatísticas (Bruto, Únicos, Repetições)
        f_stats = ctk.CTkFrame(container, fg_color="transparent")
        f_stats.pack(fill="x", padx=16, pady=(0, 10))

        def _criar_card(pai, titulo, valor, cor_val):
            card = ctk.CTkFrame(pai, fg_color=c.get("input", "#1f3a60"), corner_radius=8, border_width=1, border_color="#334155")
            card.pack(side="left", fill="x", expand=True, padx=4)
            lbl_t = ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
            lbl_t.pack(anchor="w", padx=10, pady=(6, 1))
            lbl_v = ctk.CTkLabel(card, text=str(valor), font=("Segoe UI", 14, "bold"), text_color=cor_val)
            lbl_v.pack(anchor="w", padx=10, pady=(0, 6))

        total_bruto = self.dados_analise.get("total_bruto", 0)
        total_unicos = self.dados_analise.get("total_unicos", 0)
        total_dup = self.dados_analise.get("total_duplicados", 0)
        duplicados = self.dados_analise.get("duplicados", {})

        _criar_card(f_stats, "TOTAL BRUTO", f"{total_bruto}", "white")
        _criar_card(f_stats, "SERIAIS ÚNICOS", f"{total_unicos}", "#34d399")
        _criar_card(f_stats, "REPETIÇÕES", f"{total_dup}", "#f59e0b")

        # 3. Cabeçalho das Colunas
        f_colunas = ctk.CTkFrame(container, fg_color="#1e293b", corner_radius=6, height=32)
        f_colunas.pack(fill="x", padx=16, pady=(0, 6))
        f_colunas.pack_propagate(False)

        lbl_c1 = ctk.CTkLabel(f_colunas, text="SERIAL / MEDIDOR", font=("Segoe UI", 10, "bold"), text_color="#cbd5e1", width=220, anchor="w")
        lbl_c1.pack(side="left", padx=(14, 0))

        lbl_c2 = ctk.CTkLabel(f_colunas, text="OCORRÊNCIAS NO TEXTO", font=("Segoe UI", 10, "bold"), text_color="#cbd5e1", anchor="w")
        lbl_c2.pack(side="left", padx=(20, 0))

        # 4. Lista Rolável de Duplicados
        scroll_frame = ctk.CTkScrollableFrame(container, fg_color="transparent", corner_radius=6)
        scroll_frame.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        itens_ordenados = sorted(duplicados.items(), key=lambda x: (-x[1], x[0]))

        for idx, (serial, contagem) in enumerate(itens_ordenados):
            bg_row = c.get("input", "#1f3a60") if idx % 2 == 0 else "#162740"
            f_row = ctk.CTkFrame(scroll_frame, fg_color=bg_row, corner_radius=6, border_width=1, border_color="#334155")
            f_row.pack(fill="x", pady=2)

            lbl_s = ctk.CTkLabel(f_row, text=str(serial), font=("Consolas", 13, "bold"), text_color=c.get("accent", "#f57c00"), width=220, anchor="w")
            lbl_s.pack(side="left", padx=(14, 6), pady=6)

            badge_bg = "#78350f" if contagem > 2 else "#451a03"
            f_badge = ctk.CTkFrame(f_row, fg_color=badge_bg, corner_radius=4)
            f_badge.pack(side="left", padx=14, pady=4)
            lbl_badge = ctk.CTkLabel(f_badge, text=f" Repetiu {contagem} vezes ", font=("Segoe UI", 11, "bold"), text_color="#fcd34d")
            lbl_badge.pack(padx=8, pady=3)

        # 5. Barra de Ações Inferior
        f_acoes = ctk.CTkFrame(container, fg_color="transparent")
        f_acoes.pack(fill="x", padx=16, pady=(4, 12), side="bottom")

        self.btn_copiar_dup = ctk.CTkButton(
            f_acoes,
            text="📋 Copiar Seriais Duplicados",
            height=36,
            font=("Segoe UI", 12, "bold"),
            fg_color="#f57c00",
            hover_color="#e65100",
            command=self._copiar_seriais_duplicados
        )
        self.btn_copiar_dup.pack(side="left", padx=(0, 8))

        self.btn_copiar_detalhado = ctk.CTkButton(
            f_acoes,
            text="📊 Copiar com Contagem",
            height=36,
            font=("Segoe UI", 12),
            fg_color="#334155",
            hover_color="#475569",
            command=self._copiar_detalhes
        )
        self.btn_copiar_detalhado.pack(side="left", padx=4)

        btn_fechar = ctk.CTkButton(
            f_acoes,
            text="❌ Fechar",
            height=36,
            width=100,
            font=("Segoe UI", 12),
            fg_color="#1e293b",
            hover_color="#334155",
            command=self.destroy
        )
        btn_fechar.pack(side="right", padx=(8, 0))

    def _copiar_seriais_duplicados(self):
        duplicados = self.dados_analise.get("duplicados", {})
        texto = "\r\n".join(str(s) for s in duplicados.keys())
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(texto)
            win32clipboard.CloseClipboard()
        except Exception:
            self.clipboard_clear()
            self.clipboard_append(texto)
            self.update()

        txt_antigo = self.btn_copiar_dup.cget("text")
        self.btn_copiar_dup.configure(text="✅ Seriais Duplicados Copiados!")
        self.after(2000, lambda: self.btn_copiar_dup.configure(text=txt_antigo))

    def _copiar_detalhes(self):
        duplicados = self.dados_analise.get("duplicados", {})
        linhas = [f"{s}\t{count}x" for s, count in duplicados.items()]
        texto = "\r\n".join(linhas)
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(texto)
            win32clipboard.CloseClipboard()
        except Exception:
            self.clipboard_clear()
            self.clipboard_append(texto)
            self.update()

        txt_antigo = self.btn_copiar_detalhado.cget("text")
        self.btn_copiar_detalhado.configure(text="✅ Detalhes Copiados!")
        self.after(2000, lambda: self.btn_copiar_detalhado.configure(text=txt_antigo))


def obter_script_iq09():
    return obter_caminho_script("SCRIPT DA IQ09.vbs")

SCRIPT_IQ09 = obter_script_iq09()

def decodificar_saida_cscript(stdout_data):
    """Decodifica a saída do cscript.exe tratando a codificação OEM (cp850) do Windows para evitar caracteres corrompidos."""
    if not stdout_data:
        return ""
    if isinstance(stdout_data, str):
        stdout_data = stdout_data.encode("latin1", errors="replace")
    for enc in ["cp850", "cp1252", "utf-8"]:
        try:
            return stdout_data.decode(enc)
        except Exception:
            continue
    return stdout_data.decode("utf-8", errors="replace")

def extrair_seriais_com_analise(texto_ou_lista) -> dict:
    """
    Extrai e expande números de série / medidores (numéricos ou alfanuméricos como 'UQM9HDN26L01533345'),
    processando intervalos (ex: '18400025 a 18400028') e listas separadas por vírgula, espaço ou quebra de linha.
    Retorna métricas completas com detecção de seriais duplicados.
    """
    if isinstance(texto_ou_lista, (list, tuple, set)):
        texto = "\n".join(str(x) for x in texto_ou_lista)
    else:
        texto = str(texto_ou_lista or "")

    seriais_brutos = []
    # 1. Intervalos numéricos: ex: 18400025 a 18400028
    padrao_intervalo = re.compile(r"\b(\d{4,18})\s*(?:a|A|-|ate|at[eé]|à)\s*(\d{4,18})\b", re.IGNORECASE)

    for match in padrao_intervalo.finditer(texto):
        ini_str, fim_str = match.groups()
        ini, fim = int(ini_str), int(fim_str)
        if abs(fim - ini) <= 1000:
            passo = 1 if fim >= ini else -1
            largura = max(len(ini_str), len(fim_str))
            preservar = ini_str.startswith("0") or fim_str.startswith("0")
            for n in range(ini, fim + passo, passo):
                seriais_brutos.append(str(n).zfill(largura) if preservar else str(n))

    texto_sem_intervalo = padrao_intervalo.sub(" ", texto)

    # 2. Medidores Alfanuméricos (ex: UQM9HDN26L01533345, NANSEN, LANDIS, etc.)
    # Tokens de 7 a 22 caracteres que contenham letras e dígitos (pelo menos 2 dígitos)
    tokens_alfa = re.findall(r"\b(?=[A-Za-z0-9_\-]*[A-Za-z])(?=[A-Za-z0-9_\-]*\d)[A-Za-z0-9_\-]{7,22}\b", texto_sem_intervalo)
    for t in tokens_alfa:
        t_limpo = t.strip().upper()
        if sum(c.isdigit() for c in t_limpo) >= 2 and t_limpo not in ["CONECTA", "WORKFORCE", "ASSISTANT", "BANDEIRADA", "TRANSFERENCIA", "CONFERENCIA"]:
            seriais_brutos.append(t_limpo)

    for t in tokens_alfa:
        texto_sem_intervalo = texto_sem_intervalo.replace(t, " ")

    # 3. Numéricos puros isolados (mínimo 4 dígitos para medidores/seriais)
    isolados = re.findall(r"\b\d{4,18}\b", texto_sem_intervalo)
    seriais_brutos.extend(isolados)

    # Contagem e detecção de duplicados
    contagem = Counter(seriais_brutos)

    unicos = []
    vistos = set()
    for s in seriais_brutos:
        if s not in vistos:
            vistos.add(s)
            unicos.append(s)

    duplicados = {s: count for s, count in contagem.items() if count > 1}
    total_bruto = len(seriais_brutos)
    total_unicos = len(unicos)
    total_duplicados = total_bruto - total_unicos

    return {
        "todos": seriais_brutos,
        "unicos": unicos,
        "duplicados": duplicados,
        "total_bruto": total_bruto,
        "total_unicos": total_unicos,
        "total_duplicados": total_duplicados,
    }

def extrair_lista_seriais(texto_ou_lista) -> list:
    """Extrai lista de seriais únicos (preserva compatibilidade retroativa com todo o projeto)."""
    return extrair_seriais_com_analise(texto_ou_lista)["unicos"]

def parse_todos_medidores(stdout_text) -> list:
    """Extrai todas as informações de medidores retornadas pelo script VBS."""
    if not stdout_text:
        return []
    lista = []
    for linha in stdout_text.splitlines():
        linha = linha.strip()
        if not linha.startswith("MEDIDOR_INFO:"):
            continue
        dados_raw = linha.split(":", 1)[1]
        partes = dict(item.split("=", 1) for item in dados_raw.split("|") if "=" in item)

        st = partes.get("STATUS", "").strip()
        st = re.sub(r"\s+", " ", st)
        if not st or st.lower() == "status":
            st = "Sem informação"
        else:
            # Normaliza múltiplos status separados por espaço ou barra (ex: "LIDI DISP" -> "LIDI / DISP")
            tokens = [t.strip() for t in st.replace("/", " ").split() if t.strip()]
            if len(tokens) > 1:
                st = " / ".join(tokens)
            elif len(tokens) == 1:
                st = tokens[0]

        eq = partes.get("EQ", "").strip()
        if not eq or eq.lower() == "equipamento":
            eq = "—"

        desc = partes.get("DESC", "").strip()
        if "denomina" in desc.lower():
            desc = ""

        mat = partes.get("MAT", "").strip()
        if not mat or mat.lower() == "material":
            mat = "—"

        ser = partes.get("SERIAL", "").strip()
        if not ser or "série" in ser.lower() or "serie" in ser.lower():
            ser = "—"

        centro = partes.get("CENTRO", "").strip()
        if not centro or centro in ["—", "-", "None"] or "sem info" in centro.lower() or centro.lower() == "centro":
            centro = "Sem informação"

        deposito = partes.get("DEP", "").strip()
        if not deposito or deposito in ["—", "-", "None"] or "sem info" in deposito.lower() or "deposito" in deposito.lower() or "depósito" in deposito.lower():
            deposito = "Sem informação"

        tipo = partes.get("TIPO", "").strip()
        if "tipo" in tipo.lower():
            tipo = ""

        info = {
            "equipamento": eq or "—",
            "denominacao": desc,
            "status": st,
            "material": mat or "—",
            "serial": ser,
            "centro": centro,
            "deposito": deposito,
            "tipo_estoque": tipo
        }
        if not info["denominacao"] and info["material"] and info["material"] != "—":
            desc_loc = buscar_descricao_material_local(info["material"])
            if desc_loc:
                info["denominacao"] = desc_loc

        lista.append(info)
    return lista

def parse_info_medidor(stdout_text):
    """Extrai a primeira informação estruturada do medidor do retorno do script VBS."""
    todos = parse_todos_medidores(stdout_text)
    return todos[0] if todos else None

def copiar_texto_clipboard(texto: str):
    """Copia texto para a área de transferência do Windows de forma resiliente."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(texto)
        win32clipboard.CloseClipboard()
        return True
    except Exception:
        pass
    try:
        import tkinter as tk
        r = tk.Tk()
        r.withdraw()
        r.clipboard_clear()
        r.clipboard_append(texto)
        r.update()
        r.destroy()
        return True
    except Exception:
        pass
    return False

def consultar_medidor_iq09_sap(serial_ou_seriais, callback_ao_concluir=None) -> dict:
    """
    Executa a consulta de medidor/serial diretamente na transação IQ09 do SAP,
    suportando múltiplos seriais e intervalos (ex: '18400025 a 18400028').
    
    Pode ser executado de forma síncrona dentro da thread do Chatbot.
    """
    if not verificar_sap_aberto():
        return {"erro": "O SAP precisa estar aberto para pesquisar na IQ09. Abra o SAP e tente novamente."}

    seriais_extraidos = extrair_lista_seriais(serial_ou_seriais)
    if not seriais_extraidos:
        texto_puro = str(serial_ou_seriais or "").strip()
        if texto_puro:
            seriais_extraidos = [l.strip() for l in texto_puro.splitlines() if l.strip()]

    if not seriais_extraidos:
        return {"erro": "Informe o número do medidor ou serial para pesquisar na IQ09."}

    serial_principal = seriais_extraidos[0]
    copiar_texto_clipboard("\r\n".join(seriais_extraidos))

    script_iq09 = obter_script_iq09()
    if not os.path.exists(script_iq09):
        return {"erro": f"Script SCRIPT DA IQ09.vbs não encontrado em:\n{script_iq09}"}

    script_temp = None
    try:
        with open(script_iq09, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()
        
        # Insere o reset automático se necessário
        if "Reset automatico inserido pelo Conecta Hub" not in conteudo:
            ponto_ancora = 'session.findById("wnd[0]").maximize'
            bloco_reset = (
                "' Reset automatico inserido pelo Conecta Hub\n"
                'On Error Resume Next\n'
                'session.findById("wnd[0]/tbar[0]/btn[12]").press\n'
                'session.findById("wnd[0]/tbar[0]/okcd").text = "/nS000"\n'
                'session.findById("wnd[0]").sendVKey 0\n'
                'WScript.Sleep 500\n'
                'If session.Children.Count > 1 Then\n'
                '    session.findById("wnd[1]/tbar[0]/btn[0]").press\n'
                'End If\n'
                'On Error GoTo 0\n'
            )
            if ponto_ancora in conteudo:
                conteudo = conteudo.replace(ponto_ancora, ponto_ancora + "\n" + bloco_reset, 1)
            else:
                conteudo = bloco_reset + conteudo

        fd, script_temp = tempfile.mkstemp(prefix="iq09_chat_", suffix=".vbs", text=True)
        with os.fdopen(fd, "w", encoding="mbcs") as f:
            f.write(conteudo)

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        resultado = subprocess.run(
            ["cscript.exe", "//nologo", script_temp],
            capture_output=True, timeout=None, creationflags=creationflags
        )

        stdout_str = decodificar_saida_cscript(resultado.stdout)
        if resultado.returncode == 0:
            todos = parse_todos_medidores(stdout_str)
            if len(todos) > 1:
                if callback_ao_concluir:
                    callback_ao_concluir(todos)
                return {
                    "status": f"Consulta na IQ09 concluída com sucesso! ({len(todos)} medidores encontrados)",
                    "lista_medidores": todos,
                    "dados_medidor": todos[0]
                }
            elif len(todos) == 1:
                info = todos[0]
                if callback_ao_concluir:
                    callback_ao_concluir(info)
                return {
                    "status": "Consulta na IQ09 concluída com sucesso!",
                    "dados_medidor": info,
                    "lista_medidores": todos
                }
            else:
                info_vazia = {
                    "equipamento": "—",
                    "denominacao": buscar_descricao_material_local(serial_principal) or "—",
                    "status": "Não localizado no SAP",
                    "material": "—",
                    "serial": serial_principal,
                    "centro": "Sem informação",
                    "deposito": "Sem informação",
                    "tipo_estoque": ""
                }
                return {
                    "status": "Nenhum medidor retornado pelo SAP.",
                    "dados_medidor": info_vazia,
                    "lista_medidores": []
                }
        else:
            stderr_str = decodificar_saida_cscript(resultado.stderr)
            erro = (stderr_str or stdout_str or "Falha ao executar script da IQ09 no SAP.").strip()
            return {"erro": erro}

    except subprocess.TimeoutExpired:
        return {"erro": "Tempo limite excedido ao executar consulta da IQ09 no SAP."}
    except Exception as e:
        return {"erro": f"Erro inesperado ao consultar IQ09: {e}"}
    finally:
        if script_temp:
            try: os.remove(script_temp)
            except OSError: pass

class AbaExtrator:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando_iq09 = False
        self.popup_medidor_ref = None
        self.popup_duplicados_ref = None
        self.dados_analise_atuais = {}

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        for widget in [self.txt_entrada, self.txt_sem_ponto_virgula, self.txt_com_ponto_virgula]:
            widget.configure(fg_color=c["input"], text_color=c["texto"], border_color=c["container"], border_width=1)

        for card in [self.card_total, self.card_unicos, self.card_duplicados]:
            card.configure(fg_color=c["input"], border_color=c["container"])

        self.btn_extrair.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_iq09.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")

        for btn in [self.btn_copiar_sem, self.btn_copiar_com, self.btn_limpar]:
            btn.configure(fg_color=c["container"], hover_color=c["hover"], text_color=c["texto"])

        self.btn_ver_duplicados.configure(fg_color=c.get("accent", "#f57c00"), hover_color=c.get("hover", "#e65100"), text_color="white")

        self.lbl_titulo.configure(text_color=c["texto"])
        self.lbl_sem.configure(text_color=c["texto"])
        self.lbl_com.configure(text_color=c["texto"])
        self.lbl_val_total.configure(text_color=c["texto"])
        self.lbl_val_unicos.configure(text_color=c.get("accent", "#f57c00"))
        self.container.configure(fg_color="transparent")
        self.frame_resultados.configure(fg_color="transparent")
        self.frame_botoes_resultado.configure(fg_color="transparent")
        self.frame_acoes.configure(fg_color="transparent")

    def extrair_seriais(self):
        texto = self.txt_entrada.get("1.0", tk.END)
        analise = extrair_seriais_com_analise(texto)
        seriais = analise["unicos"]
        self.dados_analise_atuais = analise

        if not seriais:
            self._limpar_metricas()
            messagebox.showwarning("Aviso", "Nenhum serial encontrado para extrair.")
            return

        sem_ponto_virgula = "\n".join(seriais)
        com_ponto_virgula = "\n".join(f"{serial};" for serial in seriais)

        self._set_texto(self.txt_sem_ponto_virgula, sem_ponto_virgula)
        self._set_texto(self.txt_com_ponto_virgula, com_ponto_virgula)

        c = self._cores()
        total_b = analise["total_bruto"]
        total_u = analise["total_unicos"]
        total_d = analise["total_duplicados"]
        qtd_ser_dup = len(analise["duplicados"])

        self.lbl_val_total.configure(text=f"{total_b} seriais")
        self.lbl_val_unicos.configure(text=f"{total_u} únicos", text_color=c.get("accent", "#f57c00"))

        if total_d > 0:
            self.lbl_val_duplicados.configure(
                text=f"⚠️ {total_d} repetições ({qtd_ser_dup} seriais)",
                text_color="#f59e0b"
            )
            self.card_duplicados.configure(border_color="#f59e0b")
            self.btn_ver_duplicados.configure(
                text=f"🔍 Ver {qtd_ser_dup} Duplicados",
                fg_color="#f59e0b",
                hover_color="#d97706"
            )
            self.btn_ver_duplicados.pack(side="right", padx=(8, 0))
            self._status(f"{total_u} seriais únicos extraídos ({total_d} repetições eliminadas).")
        else:
            self.lbl_val_duplicados.configure(
                text="✅ Nenhum duplicado",
                text_color="#34d399"
            )
            self.card_duplicados.configure(border_color=c.get("container", "#172a45"))
            self.btn_ver_duplicados.pack_forget()
            self._status(f"{total_u} seriais extraídos com sucesso (sem duplicados).")

        self.lbl_sem.configure(text=f'Tabela IQ09 sem ";"  •  Total: {total_u}')
        self.lbl_com.configure(text=f'Formatado com ";"  •  Total: {total_u}')
        incrementar_metrica("seriais_extraidos", total_u)

    def copiar_sem_ponto_virgula(self):
        self._copiar_texto(self.txt_sem_ponto_virgula, "Lista sem ponto e vírgula copiada.")

    def copiar_com_ponto_virgula(self):
        self._copiar_texto(self.txt_com_ponto_virgula, "Lista com ponto e vírgula copiada.")

    def abrir_na_iq09(self, callback_ao_concluir=None):
        if self.executando_iq09:
            messagebox.showwarning("Aviso", "A IQ09 já está em execução.")
            return

        seriais = self._obter_lista_sem_ponto_virgula()
        if not seriais:
            self.extrair_seriais()
            seriais = self._obter_lista_sem_ponto_virgula()

        if not seriais:
            return

        script_iq09 = obter_script_iq09()
        if not os.path.exists(script_iq09):
            messagebox.showerror("Erro", f"Script não encontrado:\n{script_iq09}")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP precisa estar aberto para abrir a IQ09.\nAbra o SAP e tente novamente.")
            return

        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append("\r\n".join(seriais))
        self.parent_frame.update()

        self.executando_iq09 = True
        self.btn_iq09.configure(state="disabled")
        self._status(f"Abrindo IQ09 com {len(seriais)} seriais sem ponto e vírgula...")

        thread = threading.Thread(target=self._executar_iq09, args=(callback_ao_concluir,), daemon=True)
        thread.start()

    def _limpar_metricas(self):
        c = self._cores()
        self.lbl_val_total.configure(text="0 seriais", text_color=c.get("texto", "white"))
        self.lbl_val_unicos.configure(text="0 únicos", text_color=c.get("texto", "white"))
        self.lbl_val_duplicados.configure(text="Nenhum duplicado", text_color="#34d399")
        self.card_duplicados.configure(border_color=c.get("container", "#172a45"))
        self.btn_ver_duplicados.pack_forget()
        self.lbl_sem.configure(text='Tabela IQ09 sem ";"')
        self.lbl_com.configure(text='Formatado com ";"')
        self.dados_analise_atuais = {}

    def abrir_popup_duplicados(self):
        if not self.dados_analise_atuais or not self.dados_analise_atuais.get("duplicados"):
            messagebox.showinfo("Duplicados", "Nenhum serial duplicado foi detectado nesta extração.")
            return
        try:
            if self.popup_duplicados_ref and self.popup_duplicados_ref.winfo_exists():
                self.popup_duplicados_ref.destroy()
            top = self.parent_frame.winfo_toplevel()
            self.popup_duplicados_ref = JanelaSeriaisDuplicados(
                top, self.dados_analise_atuais, self.paletas, self.obter_tema_atual
            )
        except Exception as e:
            print(f"[EXTRATOR] Erro ao abrir pop-up de duplicados: {e}")

    def limpar(self):
        for widget in [self.txt_entrada, self.txt_sem_ponto_virgula, self.txt_com_ponto_virgula]:
            self._set_texto(widget, "")
        self._limpar_metricas()
        self._status("Extrator limpo.")

    def _parse_info_medidor(self, stdout_text):
        """Extrai as informações estruturadas do medidor do retorno do script VBS."""
        return parse_info_medidor(stdout_text)

    def exibir_popup_medidor(self, info_medidor):
        """Abre o Pop-up moderno de Detalhes do Medidor."""
        if not info_medidor:
            return
        try:
            if self.popup_medidor_ref and self.popup_medidor_ref.winfo_exists():
                self.popup_medidor_ref.destroy()
            top = self.parent_frame.winfo_toplevel()
            self.popup_medidor_ref = JanelaDetalheMedidor(top, info_medidor, self.paletas, self.obter_tema_atual)
        except Exception as e:
            print(f"[EXTRATOR] Erro ao abrir pop-up de detalhes do medidor: {e}")

    def exibir_popup_multiplos_medidores(self, lista_medidores):
        """Abre o Pop-up moderno de Múltiplos Medidores."""
        if not lista_medidores:
            return
        try:
            if self.popup_medidor_ref and self.popup_medidor_ref.winfo_exists():
                self.popup_medidor_ref.destroy()
            top = self.parent_frame.winfo_toplevel()
            self.popup_medidor_ref = JanelaMultiplosMedidores(top, lista_medidores, self.paletas, self.obter_tema_atual)
        except Exception as e:
            print(f"[EXTRATOR] Erro ao abrir pop-up de múltiplos medidores: {e}")

    def _executar_iq09(self, callback_ao_concluir=None):
        script_temp = None
        try:
            script_temp = self._criar_script_iq09_temporario()
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

            resultado = subprocess.run(
                ["cscript.exe", "//nologo", script_temp],
                capture_output=True, timeout=None, creationflags=creationflags,
            )

            stdout_str = decodificar_saida_cscript(resultado.stdout)
            if resultado.returncode == 0:
                self._status_threadsafe("IQ09 aberta com a lista sem ponto e vírgula.")
                todos = parse_todos_medidores(stdout_str)
                if len(todos) > 1:
                    self.parent_frame.after(0, lambda m=todos: self.exibir_popup_multiplos_medidores(m))
                    if callback_ao_concluir:
                        self.parent_frame.after(0, lambda m=todos: callback_ao_concluir(m))
                elif len(todos) == 1:
                    self.parent_frame.after(0, lambda m=todos[0]: self.exibir_popup_medidor(m))
                    if callback_ao_concluir:
                        self.parent_frame.after(0, lambda m=todos[0]: callback_ao_concluir(m))
            else:
                stderr_str = decodificar_saida_cscript(resultado.stderr)
                erro = (stderr_str or stdout_str or "Erro desconhecido.").strip()
                self.parent_frame.after(0, lambda: messagebox.showerror("Erro IQ09", erro))
                self._status_threadsafe("Falha ao abrir IQ09.")
        except subprocess.TimeoutExpired:
            self.parent_frame.after(0, lambda: messagebox.showwarning("Aviso", "Tempo limite atingido ao abrir a IQ09."))
            self._status_threadsafe("IQ09 excedeu o tempo limite.")
        except Exception as e:
            self.parent_frame.after(0, lambda: messagebox.showerror("Erro IQ09", str(e)))
            self._status_threadsafe("Erro ao executar script da IQ09.")
        finally:
            if script_temp:
                try: os.remove(script_temp)
                except OSError: pass
            self.parent_frame.after(0, self._finalizar_iq09)

    def _criar_script_iq09_temporario(self):
        script_iq09 = obter_script_iq09()
        with open(script_iq09, "r", encoding="utf-8", errors="ignore") as f:
            conteudo = f.read()
        conteudo = self._inserir_reset_sap(conteudo)
        fd, caminho = tempfile.mkstemp(prefix="iq09_", suffix=".vbs", text=True)
        with os.fdopen(fd, "w", encoding="mbcs") as f:
            f.write(conteudo)
        return caminho

    def _inserir_reset_sap(self, conteudo):
        if "Reset automatico inserido pelo Conecta Hub" in conteudo: return conteudo
        ponto_ancora = 'session.findById("wnd[0]").maximize'
        bloco_reset = (
            "' Reset automatico inserido pelo Conecta Hub\n"
            'On Error Resume Next\n'
            'session.findById("wnd[0]/tbar[0]/btn[12]").press\n'
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nS000"\n'
            'session.findById("wnd[0]").sendVKey 0\n'
            'WScript.Sleep 500\n'
            'If session.Children.Count > 1 Then\n'
            '    session.findById("wnd[1]/tbar[0]/btn[0]").press\n'
            'End If\n'
            'On Error GoTo 0\n'
        )
        if ponto_ancora in conteudo:
            return conteudo.replace(ponto_ancora, ponto_ancora + "\n" + bloco_reset, 1)
        return bloco_reset + conteudo

    def _extrair_seriais(self, texto):
        return extrair_lista_seriais(texto)

    def _remover_duplicados_mantendo_ordem(self, seriais):
        vistos = set()
        return [s for s in seriais if not (s in vistos or vistos.add(s))]

    def _obter_lista_sem_ponto_virgula(self):
        texto = self.txt_sem_ponto_virgula.get("1.0", tk.END)
        return [linha.strip().replace(";", "") for linha in texto.splitlines() if linha.strip()]

    def _copiar_texto(self, widget, mensagem):
        texto = widget.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Não há conteúdo para copiar.")
            return
        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(texto)
        self.parent_frame.update()
        self._status(mensagem)

    def _construir_interface(self):
        self.container = ctk.CTkFrame(self.parent_frame)
        self.container.pack(fill="both", expand=True, padx=40, pady=20)

        self.lbl_titulo = ctk.CTkLabel(self.container, text="🔢 Cole os intervalos ou seriais SAP abaixo:", font=("Segoe UI", 14, "bold"))
        self.lbl_titulo.pack(anchor="w", pady=(0, 10))

        self.txt_entrada = ctk.CTkTextbox(self.container, height=120, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_entrada.pack(fill="both", expand=True, pady=(0, 15))

        self.btn_extrair = ctk.CTkButton(self.container, text="🔍 EXTRAIR SERIAL", command=self.extrair_seriais, font=("Segoe UI", 13, "bold"), height=40, corner_radius=6)
        self.btn_extrair.pack(fill="x", pady=(0, 12))

        # Painel de Métricas e Detecção de Duplicados
        self.frame_metricas = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frame_metricas.pack(fill="x", pady=(0, 12))

        # Card 1: Total Extraído
        self.card_total = ctk.CTkFrame(self.frame_metricas, corner_radius=8, border_width=1)
        self.card_total.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.lbl_tit_total = ctk.CTkLabel(self.card_total, text="📊 TOTAL COLETADO", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        self.lbl_tit_total.pack(anchor="w", padx=12, pady=(6, 1))

        self.lbl_val_total = ctk.CTkLabel(self.card_total, text="0 seriais", font=("Segoe UI", 14, "bold"))
        self.lbl_val_total.pack(anchor="w", padx=12, pady=(0, 6))

        # Card 2: Seriais Únicos
        self.card_unicos = ctk.CTkFrame(self.frame_metricas, corner_radius=8, border_width=1)
        self.card_unicos.pack(side="left", fill="x", expand=True, padx=4)

        self.lbl_tit_unicos = ctk.CTkLabel(self.card_unicos, text="🔢 SERIAIS ÚNICOS", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        self.lbl_tit_unicos.pack(anchor="w", padx=12, pady=(6, 1))

        self.lbl_val_unicos = ctk.CTkLabel(self.card_unicos, text="0 únicos", font=("Segoe UI", 14, "bold"))
        self.lbl_val_unicos.pack(anchor="w", padx=12, pady=(0, 6))

        # Card 3: Detector de Duplicados
        self.card_duplicados = ctk.CTkFrame(self.frame_metricas, corner_radius=8, border_width=1)
        self.card_duplicados.pack(side="left", fill="x", expand=True, padx=(6, 0))

        f_dup_header = ctk.CTkFrame(self.card_duplicados, fg_color="transparent")
        f_dup_header.pack(fill="x", padx=12, pady=(6, 1))

        self.lbl_tit_duplicados = ctk.CTkLabel(f_dup_header, text="🛡️ DETECTOR DE DUPLICADOS", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        self.lbl_tit_duplicados.pack(side="left")

        f_dup_corpo = ctk.CTkFrame(self.card_duplicados, fg_color="transparent")
        f_dup_corpo.pack(fill="x", padx=12, pady=(0, 6))

        self.lbl_val_duplicados = ctk.CTkLabel(f_dup_corpo, text="Nenhum duplicado", font=("Segoe UI", 14, "bold"), text_color="#34d399")
        self.lbl_val_duplicados.pack(side="left")

        self.btn_ver_duplicados = ctk.CTkButton(
            f_dup_corpo,
            text="🔍 Ver Duplicados",
            font=("Segoe UI", 11, "bold"),
            height=26,
            width=130,
            command=self.abrir_popup_duplicados
        )

        self.frame_resultados = ctk.CTkFrame(self.container)
        self.frame_resultados.pack(fill="both", expand=True)
        self.frame_resultados.columnconfigure(0, weight=1)
        self.frame_resultados.columnconfigure(1, weight=1)
        self.frame_resultados.rowconfigure(1, weight=1)

        self.lbl_sem = ctk.CTkLabel(self.frame_resultados, text='Tabela IQ09 sem ";"')
        self.lbl_sem.grid(row=0, column=0, sticky="w")
        
        self.lbl_com = ctk.CTkLabel(self.frame_resultados, text='Formatado com ";"')
        self.lbl_com.grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.txt_sem_ponto_virgula = ctk.CTkTextbox(self.frame_resultados, height=150, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_sem_ponto_virgula.grid(row=1, column=0, sticky="nsew", pady=(6, 8))

        self.txt_com_ponto_virgula = ctk.CTkTextbox(self.frame_resultados, height=150, font=("Consolas", 13), wrap="none", corner_radius=6)
        self.txt_com_ponto_virgula.grid(row=1, column=1, sticky="nsew", padx=(12, 0), pady=(6, 8))

        self.frame_botoes_resultado = ctk.CTkFrame(self.frame_resultados)
        self.frame_botoes_resultado.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.frame_botoes_resultado.columnconfigure(0, weight=1)
        self.frame_botoes_resultado.columnconfigure(1, weight=1)

        self.btn_copiar_sem = ctk.CTkButton(self.frame_botoes_resultado, text='📋 Copiar tabela sem ";"', command=self.copiar_sem_ponto_virgula, font=("Segoe UI", 12, "bold"), height=35)
        self.btn_copiar_sem.grid(row=0, column=0, sticky="ew")

        self.btn_copiar_com = ctk.CTkButton(self.frame_botoes_resultado, text='📋 Copiar com ";"', command=self.copiar_com_ponto_virgula, font=("Segoe UI", 12, "bold"), height=35)
        self.btn_copiar_com.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        self.frame_acoes = ctk.CTkFrame(self.container)
        self.frame_acoes.pack(fill="x", pady=(15, 0))

        self.btn_iq09 = ctk.CTkButton(self.frame_acoes, text='🖥️ ABRIR NA IQ09 (Tabela sem ";")', command=self.abrir_na_iq09, font=("Segoe UI", 13, "bold"), height=40)
        self.btn_iq09.pack(side="left", fill="x", expand=True)

        self.btn_limpar = ctk.CTkButton(self.frame_acoes, text="🧹 Limpar", command=self.limpar, font=("Segoe UI", 12, "bold"), width=120, height=40)
        self.btn_limpar.pack(side="left", padx=(10, 0))

    def _set_texto(self, widget, texto):
        widget.delete("1.0", tk.END)
        widget.insert("1.0", texto)

    def _finalizar_iq09(self):
        self.executando_iq09 = False
        self.btn_iq09.configure(state="normal")

    def _status(self, texto):
        if self.label_status_global:
            # CustomTkinter label uses 'text' or configure(text=)
            self.label_status_global.configure(text=f" {texto}")

    def _status_threadsafe(self, texto):
        self.parent_frame.after(0, lambda: self._status(texto))

    def _cores(self, tema_nome=None):
        tema = tema_nome or self.obter_tema_atual()
        return self.paletas.get(tema, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

def construir_aba_extrator(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaExtrator(parent_frame, label_status_global, paletas, obter_tema_atual)
