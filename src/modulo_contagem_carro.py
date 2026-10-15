# -*- coding: utf-8 -*-
"""
Módulo de Contagem de Carro — N.E.X.U.S
Executa o fluxo completo de balanço/auditoria e reatribuição de medidores de equipes em campo:
1. Extração/informação dos medidores da OSME da equipe (ZUDWM_237 ou entrada manual/bipagem).
2. Auditoria e validação na ZMCSE01 (ZMMCSE001 — Análise de Equipe).
3. Localização da Ordem Ativa mais recente da equipe na IW39 (Parceiro Z2).
4. Estorno dos medidores da ordem de origem via MIGO Z62.
5. Saída e re-imputação dos medidores na ordem mais recente via MIGO Z61.
6. Captura e exibição dos Documentos SAP gerados com registro de histórico em CSV.
"""

import os
import re
import csv
import sys
import time
import subprocess
import tempfile
import threading
from collections import OrderedDict
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import customtkinter as ctk

from modulo_Sap import verificar_sap_aberto, set_sap_ocupado
from modulo_extrator import extrair_lista_seriais
from utils_os import obter_diretorio_executavel, obter_caminho_recurso
from utils_metrics import incrementar_metrica

try:
    from utils_workforce import executar_fluxo_workforce, carregar_credenciais_workforce, _carregar_banco_nomes
except ImportError:
    try:
        from modulo_fichas import executar_fluxo_workforce, carregar_credenciais_workforce, _carregar_banco_nomes
    except ImportError:
        executar_fluxo_workforce = None
        carregar_credenciais_workforce = None
        _carregar_banco_nomes = None

DEFAULT_PALETAS = {
    "Nexus": {
        "bg": "#090d16",
        "container": "#111827",
        "texto": "#f8fafc",
        "accent": "#00d2ff",
        "hover": "#0284c7",
        "input": "#1a2436",
        "sucesso": "#10b981",
        "alerta": "#ef4444",
        "azul_sap": "#00d2ff"
    },
    "Azul Delray": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#38bdf8",
        "hover": "#0284c7",
        "input": "#1f3a60",
        "sucesso": "#10b981",
        "alerta": "#ef4444",
        "azul_sap": "#0284c7"
    },
    "Conecta Blue": {
        "bg": "#0a192f",
        "container": "#172a45",
        "texto": "#ffffff",
        "accent": "#38bdf8",
        "hover": "#0284c7",
        "input": "#1f3a60",
        "sucesso": "#10b981",
        "alerta": "#ef4444",
        "azul_sap": "#0284c7"
    }
}


def normalizar_texto(valor):
    if valor is None:
        return ""
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor))
    texto_sem_acento = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", texto_sem_acento.strip().upper())


def normalizar_serial_chave(valor):
    """
    Normaliza o número de série para chave de busca.
    Remove zeros à esquerda caso seja puramente numérico (ex: '00000000018349138' -> '18349138').
    Mantém intactos seriais alfanuméricos (ex: 'UQM9HDN26L01533345').
    """
    if valor is None:
        return ""
    import unicodedata
    texto = unicodedata.normalize("NFKD", str(valor))
    texto_sem_acento = "".join(c for c in texto if not unicodedata.combining(c))
    limpo = re.sub(r"[^A-Z0-9]", "", texto_sem_acento.strip().upper())
    if limpo.isdigit():
        return limpo.lstrip("0") or "0"
    return limpo


def validar_status_lidi_disp(sap_info):
    """
    Valida se o medidor retornado pelo SAP ZMMCSE001 / ZMCSE01 está apto para movimentação.
    Critérios acordados com o usuário:
    1. Equipe / Base: Qualquer equipe que NÃO comece com CCS, CCL, CDC, CSC, CVC é de outra base
       e deve ser IGNORADA DE IMEDIATO (divergente).
    2. St. Mov. Mercadoria: Se constar "UTILIZADO", "MONTADO", "INSTALADO", etc., deve ser
       IGNORADO DE IMEDIATO (divergente), mesmo que Sts Sistema/Usuário conste como LIDI / DISP.
    3. Status Geral: Deve indicar que está LIVRE / DISPONÍVEL (LIDI, DISP, LIVRE, DISPONIVEL, ESTQ).
    4. Restrições impeditivas (SUCA, MONT, INST, DESM, RETE, AVAR, DEFI, BLOQ) tornam divergente.
    5. Ter Nota de Serviço associada NÃO invalida o medidor.
    Retorna: (is_apto: bool, motivo: str)
    """
    if not sap_info:
        return False, "Não encontrado na ZMCSE01"

    # Regra 1: Validação de Equipe (Base Local Permitida)
    # Apenas equipes iniciadas com CCS, CCL, CDC, CSC, CVC pertencem à base local.
    PREFIXOS_BASE_PERMITIDOS = ("CCS", "CCL", "CDC", "CSC", "CVC")
    eq_val = sap_info.get("equipe", "").strip().upper()
    if eq_val and not any(eq_val.startswith(pfx) for pfx in PREFIXOS_BASE_PERMITIDOS):
        return False, f"Equipe de outra base ({eq_val}) — apenas CCS, CCL, CDC, CSC, CVC"

    status_str = sap_info.get("status", "").strip().upper()
    status_sis = sap_info.get("status_sis", "").strip().upper()
    status_usr = sap_info.get("status_usr", "").strip().upper()
    status_disp = sap_info.get("status_disp", "").strip().upper()
    status_mov = sap_info.get("status_mov", "").strip().upper()

    # Regra 2: Validação estrita de St. Mov. Mercadoria
    # Mesmo sendo LIDI/DISP no sistema, se o movimento for "Utilizado" ou restrito, NÃO PODE ser movimentado!
    if status_mov:
        if any(b in status_mov for b in ["UTILIZADO", "UTILIZ", "MONTADO", "MONT", "INSTALADO", "INST", "DESMONTADO", "DESM", "RETIRADO", "RETE", "SUCATA", "SUCA", "AVARIADO", "AVAR", "DEFEIT", "BLOQ"]):
            return False, f"Mercadoria {status_mov} (Já utilizada/instalada em campo)"
        if not any(ok in status_mov for ok in ["DISP", "LIVRE", "ESTQ"]):
            return False, f"Mercadoria {status_mov} (Status de movimento não liberado)"

    texto_status = f"{status_str} {status_sis} {status_usr} {status_disp} {status_mov}".strip().upper()

    # Restrições impeditivas absolutas (Sucatas, Avarias, Instalados, Utilizados, etc.)
    restricoes = [
        ("SUCA", "Sucata"),
        ("AVAR", "Avariado"),
        ("MONT", "Montado"),
        ("INST", "Instalado"),
        ("DESM", "Desmontado"),
        ("RETE", "Retirado"),
        ("DEFI", "Defeituoso"),
        ("BLOQ", "Bloqueado"),
        ("INAP", "Inapto"),
        ("BAIX", "Baixado"),
        ("EXTR", "Extraviado"),
        ("CLAN", "Clandestino"),
        ("UTILIZADO", "Utilizado / Montado"),
        ("USADO", "Usado"),
        ("INDISP", "Indisponível"),
    ]
    for cod, desc in restricoes:
        if cod in texto_status:
            return False, f"Restrição: {desc} ({cod})"

    # Verificação de disponibilidade / livre utilização (LIDI ou DISP ou ESTQ)
    tem_lidi = "LIDI" in texto_status or "LIVRE" in texto_status
    tem_disp = (
        "DISP" in texto_status
        or "DISPONIVEL" in texto_status
        or "DISPONÍVEL" in texto_status
        or "DISPON" in texto_status
        or "ESTQ" in texto_status
        or "ESTOQUE" in texto_status
    )

    if tem_lidi or tem_disp:
        if tem_lidi and tem_disp:
            return True, "LIDI / DISP (Disponível)"
        elif tem_disp:
            detalhe = status_sis or status_usr or status_disp or status_str or "DISP"
            return True, f"DISP (Disponível: {detalhe})"
        else:
            detalhe = status_sis or status_usr or status_disp or status_str or "LIDI"
            return True, f"LIDI (Livre: {detalhe})"
    elif not texto_status:
        # Se veio da grade ZMMCSE001 do veículo sem coluna de status informada ou status vazio
        return True, "Presente no veículo (Sem restrição)"
    else:
        return False, f"Status não disponível ({status_str or texto_status})"



def copiar_texto_clipboard_resiliente(texto):
    """Copia texto para a área de transferência do Windows de forma resiliente."""
    if not texto:
        return False
    try:
        import win32clipboard
        win32clipboard.OpenClipboard(0)
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(str(texto), win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return True
    except Exception:
        pass
    try:
        subprocess.run(["clip.exe"], input=str(texto), text=True, check=True, timeout=2)
        return True
    except Exception:
        pass
    return False


def vb_escape_string(s):
    """Escapa strings para uso em VBScript."""
    if s is None:
        return '""'
    escaped = str(s).replace('"', '""')
    return f'"{escaped}"'


def vb_escape_clipboard(s):
    if not s:
        return '""'
    texto = str(s).replace("\r\n", "\n").replace("\r", "\n")
    partes = texto.split("\n")
    exprs = [vb_escape_string(p) for p in partes]
    return " & vbCrLf & ".join(exprs)


class ModalDocumentosContagem(ctk.CTkToplevel):
    """Pop-up modal exibindo os dois Documentos SAP (Z62 e Z61) gerados na Contagem de Carro com atalho WorkForce."""
    def __init__(self, parent, doc_z62, doc_z61, equipe, ordem_origem, ordem_destino, total_medidores, itens_ignorados=None, materiais_lista=None, callback_abrir_workforce=None, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.doc_z62 = str(doc_z62).strip()
        self.doc_z61 = str(doc_z61).strip()
        self.equipe = str(equipe).strip().upper()
        self.ordem_origem = str(ordem_origem).strip()
        self.ordem_destino = str(ordem_destino).strip()
        self.total_medidores = total_medidores
        self.itens_ignorados = itens_ignorados or []
        self.materiais_lista = materiais_lista or []
        self.callback_abrir_workforce = callback_abrir_workforce

        self.title(f"🎉 Contagem Concluída — Equipe {self.equipe}")
        largura = 680
        altura = 620 if self.itens_ignorados else 540
        self.geometry(f"{largura}x{altura}")
        self.minsize(580, 460)
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.configure(fg_color="#0f172a")

        # Topo
        frame_topo = ctk.CTkFrame(self, fg_color="#059669", corner_radius=0, height=60)
        frame_topo.pack(side="top", fill="x")
        frame_topo.pack_propagate(False)

        ctk.CTkLabel(
            frame_topo,
            text=f"✅ CONTAGEM DE CARRO CONCLUÍDA — {self.equipe}",
            font=("Segoe UI", 15, "bold"),
            text_color="#ffffff"
        ).pack(side="left", padx=16, pady=8)

        # Corpo
        frame_conteudo = ctk.CTkFrame(self, fg_color="transparent")
        frame_conteudo.pack(fill="both", expand=True, padx=20, pady=15)

        # Informações da Operação
        card_info = ctk.CTkFrame(frame_conteudo, fg_color="#1e293b", corner_radius=8, border_color="#334155", border_width=1)
        card_info.pack(fill="x", pady=(0, 15))

        f_grid = ctk.CTkFrame(card_info, fg_color="transparent")
        f_grid.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(f_grid, text="Equipe:", font=("Segoe UI", 11, "bold"), text_color="#94a3b8").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ctk.CTkLabel(f_grid, text=self.equipe, font=("Segoe UI", 12, "bold"), text_color="#38bdf8").grid(row=0, column=1, sticky="w", padx=(0, 20))

        ctk.CTkLabel(f_grid, text="Total Medidores:", font=("Segoe UI", 11, "bold"), text_color="#94a3b8").grid(row=0, column=2, sticky="w", padx=(0, 10))
        ctk.CTkLabel(f_grid, text=f"{self.total_medidores} un", font=("Segoe UI", 12, "bold"), text_color="#10b981").grid(row=0, column=3, sticky="w")

        ctk.CTkLabel(f_grid, text="Ordem Origem:", font=("Segoe UI", 11, "bold"), text_color="#94a3b8").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(6, 0))
        ctk.CTkLabel(f_grid, text=self.ordem_origem or "(Múltiplas/Conforme Item)", font=("Segoe UI", 11, "bold"), text_color="#cbd5e1").grid(row=1, column=1, sticky="w", padx=(0, 20), pady=(6, 0))

        ctk.CTkLabel(f_grid, text="Ordem Destino (IW39):", font=("Segoe UI", 11, "bold"), text_color="#94a3b8").grid(row=1, column=2, sticky="w", padx=(0, 10), pady=(6, 0))
        ctk.CTkLabel(f_grid, text=self.ordem_destino or "(Mesma / Atual)", font=("Segoe UI", 11, "bold"), text_color="#38bdf8").grid(row=1, column=3, sticky="w", pady=(6, 0))

        # Cards com os Documentos
        f_docs = ctk.CTkFrame(frame_conteudo, fg_color="transparent")
        f_docs.pack(fill="x", pady=(0, 12))

        # Card Doc Z62
        card_z62 = ctk.CTkFrame(f_docs, fg_color="#1e293b", corner_radius=8, border_color="#38bdf8", border_width=1)
        card_z62.pack(side="left", fill="both", expand=True, padx=(0, 8))

        doc_z62_lbl = "N/A (Sem Estorno)" if self.doc_z62 == "NÃO REQUERIDO (SEM ESTORNO)" else self.doc_z62
        ctk.CTkLabel(card_z62, text="DOC. SAP Z62 (ESTORNO AO DEPÓSITO)", font=("Segoe UI", 10, "bold"), text_color="#94a3b8").pack(pady=(10, 2))
        ctk.CTkLabel(card_z62, text=doc_z62_lbl, font=("Consolas", 14 if "N/A" in doc_z62_lbl else 18, "bold"), text_color="#38bdf8").pack(pady=(0, 8))

        btn_copiar_z62 = ctk.CTkButton(
            card_z62, text="📋 Copiar Z62", font=("Segoe UI", 11, "bold"), height=28,
            fg_color="#0284c7", hover_color="#0369a1",
            state="disabled" if self.doc_z62 == "NÃO REQUERIDO (SEM ESTORNO)" else "normal",
            command=lambda: self._copiar_texto(self.doc_z62)
        )
        btn_copiar_z62.pack(pady=(0, 10), padx=20, fill="x")

        # Card Doc Z61
        card_z61 = ctk.CTkFrame(f_docs, fg_color="#1e293b", corner_radius=8, border_color="#10b981", border_width=1)
        card_z61.pack(side="right", fill="both", expand=True, padx=(8, 0))

        ctk.CTkLabel(card_z61, text="DOC. SAP Z61 (SAÍDA PARA ORDEM)", font=("Segoe UI", 10, "bold"), text_color="#94a3b8").pack(pady=(10, 2))
        ctk.CTkLabel(card_z61, text=self.doc_z61, font=("Consolas", 18, "bold"), text_color="#10b981").pack(pady=(0, 8))

        btn_copiar_z61 = ctk.CTkButton(
            card_z61, text="📋 Copiar Z61", font=("Segoe UI", 11, "bold"), height=28,
            fg_color="#059669", hover_color="#047857", command=lambda: self._copiar_texto(self.doc_z61)
        )
        btn_copiar_z61.pack(pady=(0, 10), padx=20, fill="x")

        # Card WorkForce
        card_wf = ctk.CTkFrame(frame_conteudo, fg_color="#1e293b", corner_radius=8, border_color="#0284c7", border_width=1)
        card_wf.pack(fill="x", pady=(0, 12))

        f_wf_top = ctk.CTkFrame(card_wf, fg_color="transparent")
        f_wf_top.pack(fill="x", padx=12, pady=(8, 4))

        ctk.CTkLabel(
            f_wf_top,
            text="📋 FICHA WORKFORCE (SOLICITAÇÃO DE MATERIAIS)",
            font=("Segoe UI", 10, "bold"),
            text_color="#38bdf8"
        ).pack(side="left")

        resumo_mats = ", ".join([f"{qtd}x {m}" for m, qtd in self.materiais_lista]) or f"{self.total_medidores} medidor(es)"
        lbl_wf_info = ctk.CTkLabel(
            card_wf,
            text=f"Base: SANTO ANDRÉ  •  Resp. Almox: 19595  •  Recurso: {self.equipe}\nMateriais: {resumo_mats}",
            font=("Segoe UI", 10),
            text_color="#94a3b8",
            justify="left"
        )
        lbl_wf_info.pack(anchor="w", padx=12, pady=(0, 6))

        if self.callback_abrir_workforce:
            self.btn_wf = ctk.CTkButton(
                card_wf,
                text="🌐 Abrir Ficha no WorkForce",
                font=("Segoe UI", 11, "bold"),
                height=32,
                fg_color="#0284c7",
                hover_color="#0369a1",
                command=self.callback_abrir_workforce
            )
            self.btn_wf.pack(fill="x", padx=12, pady=(0, 8))

        # Card de Itens Ignorados (se houver)
        if self.itens_ignorados:
            card_ign = ctk.CTkFrame(frame_conteudo, fg_color="#1e1b4b", corner_radius=8, border_color="#f59e0b", border_width=1)
            card_ign.pack(fill="x", pady=(0, 12))

            f_ign_top = ctk.CTkFrame(card_ign, fg_color="transparent")
            f_ign_top.pack(fill="x", padx=12, pady=(8, 4))

            ctk.CTkLabel(
                f_ign_top,
                text=f"⚠️ {len(self.itens_ignorados)} MEDIDOR(ES) IGNORADO(S) (NÃO LIDI/DISP):",
                font=("Segoe UI", 10, "bold"),
                text_color="#fbbf24"
            ).pack(side="left")

            txt_ign = ctk.CTkTextbox(card_ign, height=65, font=("Consolas", 10), fg_color="#0f172a", text_color="#fde68a")
            txt_ign.pack(fill="x", padx=12, pady=(0, 8))
            linhas_ign = [f"• Serial {it.get('serial', '—')}: {it.get('motivo', 'Não LIDI/DISP')}" for it in self.itens_ignorados]
            txt_ign.insert("1.0", "\n".join(linhas_ign))
            txt_ign.configure(state="disabled")

        # Botão Fechar
        btn_fechar = ctk.CTkButton(
            frame_conteudo, text="Concluir e Fechar", font=("Segoe UI", 12, "bold"), height=36,
            fg_color="#334155", hover_color="#475569", command=self.destroy
        )
        btn_fechar.pack(side="bottom", fill="x")

    def _copiar_texto(self, texto):
        if copiar_texto_clipboard_resiliente(texto):
            messagebox.showinfo("Copiado!", f"Documento SAP copiado para a área de transferência:\n\n{texto}")


class AbaContagemCarro:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")
        self.executando = False

        # Dados em memória
        self.itens_carregados = []  # Lista de dicts {serial, material, ordem_origem, status_val, equipe_val}
        self.ordem_mais_recente_iw39 = ""
        self.mapa_nomes = {}
        self.itens_aptos_auditados = []
        self.itens_divergentes_auditados = []

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self._log("🚗 Módulo de Contagem de Carro inicializado e pronto.")

        threading.Thread(target=self._carregar_banco_nomes_async, daemon=True).start()

    def _carregar_banco_nomes_async(self):
        try:
            if _carregar_banco_nomes:
                mapa = _carregar_banco_nomes()
                self.mapa_nomes = mapa or {}
                self.parent_frame.after(0, self._ao_digitar_re_eletricista)
                self.parent_frame.after(0, self._ao_digitar_re_eletricista_2)
        except Exception:
            self.mapa_nomes = {}

    def _ao_digitar_re_eletricista(self, event=None):
        re_txt = self.entry_re_eletricista.get().strip() if hasattr(self, "entry_re_eletricista") else ""
        if not re_txt:
            if hasattr(self, "lbl_nome_eletricista"):
                self.lbl_nome_eletricista.configure(text="👤 Eletricista 1: não informado", text_color="#94a3b8")
            return
        if hasattr(self, "mapa_nomes") and re_txt in self.mapa_nomes:
            nome = self.mapa_nomes[re_txt]
            if hasattr(self, "lbl_nome_eletricista"):
                self.lbl_nome_eletricista.configure(text=f"👤 {nome}", text_color="#34d399")
        else:
            if hasattr(self, "lbl_nome_eletricista"):
                self.lbl_nome_eletricista.configure(text="👤 RE não localizado no Banco de Nomes", text_color="#fbbf24")

    def _ao_digitar_re_eletricista_2(self, event=None):
        re_txt = self.entry_re_eletricista_2.get().strip() if hasattr(self, "entry_re_eletricista_2") else ""
        if not re_txt:
            if hasattr(self, "lbl_nome_eletricista_2"):
                self.lbl_nome_eletricista_2.configure(text="👤 Eletricista 2: não informado", text_color="#94a3b8")
            return
        if hasattr(self, "mapa_nomes") and re_txt in self.mapa_nomes:
            nome = self.mapa_nomes[re_txt]
            if hasattr(self, "lbl_nome_eletricista_2"):
                self.lbl_nome_eletricista_2.configure(text=f"👤 {nome}", text_color="#34d399")
        else:
            if hasattr(self, "lbl_nome_eletricista_2"):
                self.lbl_nome_eletricista_2.configure(text="👤 RE não localizado no Banco de Nomes", text_color="#fbbf24")

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

    def _status(self, msg):
        if self.label_status_global:
            self.label_status_global.configure(text=f" {msg}")

    def _log(self, msg):
        self.txt_log.configure(state="normal")
        self.txt_log.insert(tk.END, msg)
        if not msg.endswith("\n"):
            self.txt_log.insert(tk.END, "\n")
        self.txt_log.see(tk.END)
        self.txt_log.configure(state="disabled")

    def _log_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._log(msg))

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        bg = c.get("bg", "#090d16")
        container = c.get("container", "#111827")
        inp = c.get("input", "#1a2436")
        texto = c.get("texto", "#f8fafc")
        accent = c.get("accent", "#00d2ff")

        self.parent_frame.configure(fg_color=bg)
        if hasattr(self, "container"):
            self.container.configure(fg_color=container)
        if hasattr(self, "txt_seriais"):
            self.txt_seriais.configure(fg_color=inp, text_color=texto)
        if hasattr(self, "txt_log"):
            self.txt_log.configure(fg_color=inp, text_color=texto)

        try:
            estilo = ttk.Style()
            estilo.configure(
                "Treeview",
                background=container,
                foreground=texto,
                fieldbackground=inp,
                bordercolor=container
            )
            estilo.map(
                "Treeview",
                background=[("selected", accent)],
                foreground=[("selected", "white")]
            )
        except Exception:
            pass

    def _construir_interface(self):
        # Container Principal
        self.container = ctk.CTkFrame(self.parent_frame, corner_radius=10)
        self.container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # Topo
        frame_header = ctk.CTkFrame(self.container, fg_color="transparent")
        frame_header.pack(fill=tk.X, padx=15, pady=(12, 6))

        lbl_titulo = ctk.CTkLabel(
            frame_header,
            text="🚗 Contagem de Carro — Balanço & Reatribuição de Medidores",
            font=("Segoe UI", 16, "bold")
        )
        lbl_titulo.pack(side=tk.LEFT)

        self.btn_wf_topo = ctk.CTkButton(
            frame_header,
            text="🌐 Abrir Ficha no WorkForce",
            font=("Segoe UI", 11, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=30,
            command=self.abrir_ficha_workforce_acao
        )
        self.btn_wf_topo.pack(side=tk.RIGHT, padx=(10, 0))

        lbl_sub = ctk.CTkLabel(
            frame_header,
            text="ZMCSE01 (Auditoria dos Bipados) ➡️ IW39 (Ordem Ativa) ➡️ MIGO Z62 (Estorno) ➡️ MIGO Z61 (Reatribuição)",
            font=("Segoe UI", 11),
            text_color="#94a3b8"
        )
        lbl_sub.pack(side=tk.RIGHT, pady=(3, 0))

        # Divisão em duas colunas principais: Esquerda (Entrada e Ações) e Direita (Grid e Logs)
        frame_corpo = ctk.CTkFrame(self.container, fg_color="transparent")
        frame_corpo.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        # ==========================================
        # COLUNA ESQUERDA: ENTRADAS & AÇÕES (Rolável)
        # ==========================================
        col_esq = ctk.CTkScrollableFrame(frame_corpo, fg_color="transparent", width=430)
        col_esq.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        # Card 1: Identificação da Equipe
        card_eq = ctk.CTkFrame(col_esq, corner_radius=8, border_color="#334155", border_width=1)
        card_eq.pack(fill=tk.X, pady=(0, 10))

        ctk.CTkLabel(
            card_eq,
            text="1. EQUIPE / VEÍCULO A AUDITAR:",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        ).pack(anchor="w", padx=12, pady=(10, 4))

        f_eq_linha = ctk.CTkFrame(card_eq, fg_color="transparent")
        f_eq_linha.pack(fill=tk.X, padx=12, pady=(0, 6))

        f_eq_col = ctk.CTkFrame(f_eq_linha, fg_color="transparent")
        f_eq_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ctk.CTkLabel(f_eq_col, text="Equipe *:", font=("Segoe UI", 9, "bold"), text_color="#94a3b8").pack(anchor="w")
        self.entry_equipe = ctk.CTkEntry(
            f_eq_col,
            placeholder_text="Ex: CSC300",
            font=("Consolas", 13, "bold"),
            height=30
        )
        self.entry_equipe.pack(fill=tk.X, pady=(1, 0))

        f_dt_col = ctk.CTkFrame(f_eq_linha, fg_color="transparent")
        f_dt_col.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        ctk.CTkLabel(f_dt_col, text="📅 Data Lçto. (Fechamento SAP):", font=("Segoe UI", 9, "bold"), text_color="#38bdf8").pack(anchor="w")
        self.entry_data_lancamento = ctk.CTkEntry(
            f_dt_col,
            placeholder_text="dd.mm.aaaa (opcional)",
            font=("Consolas", 11),
            height=30
        )
        self.entry_data_lancamento.pack(fill=tk.X, pady=(1, 0))

        # Linha para RE dos Eletricistas da Equipe
        f_re_linha = ctk.CTkFrame(card_eq, fg_color="transparent")
        f_re_linha.pack(fill=tk.X, padx=12, pady=(0, 4))

        f_re1 = ctk.CTkFrame(f_re_linha, fg_color="transparent")
        f_re1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ctk.CTkLabel(f_re1, text="RE Eletricista 1 *:", font=("Segoe UI", 9, "bold"), text_color="#94a3b8").pack(anchor="w")
        self.entry_re_eletricista = ctk.CTkEntry(f_re1, placeholder_text="Ex: 21475", font=("Segoe UI", 11), height=28)
        self.entry_re_eletricista.pack(fill=tk.X, pady=(1, 0))
        self.entry_re_eletricista.bind("<KeyRelease>", self._ao_digitar_re_eletricista)
        self.lbl_nome_eletricista = ctk.CTkLabel(f_re1, text="👤 Eletricista 1: não informado", font=("Segoe UI", 8), text_color="#94a3b8", anchor="w")
        self.lbl_nome_eletricista.pack(anchor="w", pady=(1, 0))

        f_re2 = ctk.CTkFrame(f_re_linha, fg_color="transparent")
        f_re2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        ctk.CTkLabel(f_re2, text="RE Eletricista 2:", font=("Segoe UI", 9, "bold"), text_color="#94a3b8").pack(anchor="w")
        self.entry_re_eletricista_2 = ctk.CTkEntry(f_re2, placeholder_text="Ex: 22739", font=("Segoe UI", 11), height=28)
        self.entry_re_eletricista_2.pack(fill=tk.X, pady=(1, 0))
        self.entry_re_eletricista_2.bind("<KeyRelease>", self._ao_digitar_re_eletricista_2)
        self.lbl_nome_eletricista_2 = ctk.CTkLabel(f_re2, text="👤 Eletricista 2: não informado", font=("Segoe UI", 8), text_color="#94a3b8", anchor="w")
        self.lbl_nome_eletricista_2.pack(anchor="w", pady=(1, 0))

        # Checkbox Salvar Automático no WorkForce
        self.chk_salvar_auto = ctk.CTkCheckBox(
            card_eq,
            text="Salvar ficha automaticamente no WorkForce",
            font=("Segoe UI", 9),
            text_color="#cbd5e1",
            checkbox_width=16,
            checkbox_height=16
        )
        self.chk_salvar_auto.pack(anchor="w", padx=12, pady=(2, 8))
        self.chk_salvar_auto.select()

        # Card 2: Medidores / Seriais
        card_seriais = ctk.CTkFrame(col_esq, corner_radius=8, border_color="#334155", border_width=1)
        card_seriais.pack(fill=tk.X, pady=(0, 10))

        f_ser_top = ctk.CTkFrame(card_seriais, fg_color="transparent")
        f_ser_top.pack(fill=tk.X, padx=12, pady=(10, 4))

        ctk.CTkLabel(
            f_ser_top,
            text="2. MEDIDORES FÍSICOS BIPADOS NO VEÍCULO:",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        ).pack(side=tk.LEFT)

        self.lbl_qtd_seriais = ctk.CTkLabel(
            f_ser_top,
            text="0 medidores",
            font=("Segoe UI", 10, "bold"),
            text_color="#94a3b8"
        )
        self.lbl_qtd_seriais.pack(side=tk.RIGHT)

        self.txt_seriais = ctk.CTkTextbox(
            card_seriais,
            font=("Consolas", 12),
            height=140
        )
        self.txt_seriais.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))
        self.txt_seriais.bind("<KeyRelease>", self._ao_alterar_seriais)

        # Botões de Auditoria Prévia
        f_btn_auditoria = ctk.CTkFrame(card_seriais, fg_color="transparent")
        f_btn_auditoria.pack(fill=tk.X, padx=12, pady=(0, 10))

        self.btn_consultar_zmcse = ctk.CTkButton(
            f_btn_auditoria,
            text="🔬 1. Auditar ZMCSE01",
            font=("Segoe UI", 11, "bold"),
            fg_color="#4f46e5",
            hover_color="#4338ca",
            height=32,
            command=self.consultar_zmcse01_acao
        )
        self.btn_consultar_zmcse.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.btn_buscar_iw39 = ctk.CTkButton(
            f_btn_auditoria,
            text="🔎 2. Buscar IW39",
            font=("Segoe UI", 11, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=32,
            command=self.buscar_ordem_iw39_acao
        )
        self.btn_buscar_iw39.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # Botão de Limpeza/Remoção de Divergentes
        self.btn_remover_divergentes = ctk.CTkButton(
            card_seriais,
            text="🧹 Remover Divergentes da Lista",
            font=("Segoe UI", 11, "bold"),
            fg_color="#be123c",
            hover_color="#9f1239",
            height=30,
            command=self.remover_divergentes_acao
        )
        self.btn_remover_divergentes.pack(fill=tk.X, padx=12, pady=(0, 10))
        self.btn_remover_divergentes.configure(state="disabled")

        # Card 3: Execução do Ciclo Completo
        card_exec = ctk.CTkFrame(col_esq, corner_radius=8, border_color="#334155", border_width=1)
        card_exec.pack(fill=tk.X)

        self.btn_executar_tudo = ctk.CTkButton(
            card_exec,
            text="🚀 Executar Ciclo Completo no SAP\n(ZMCSE01 ➡️ IW39 ➡️ Z62 ➡️ Z61)",
            font=("Segoe UI", 12, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=48,
            command=self.executar_ciclo_completo_acao
        )
        self.btn_executar_tudo.pack(fill=tk.X, padx=12, pady=(10, 6))

        self.btn_abrir_wf = ctk.CTkButton(
            card_exec,
            text="🌐 Abrir Ficha no WorkForce",
            font=("Segoe UI", 11, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=32,
            command=self.abrir_ficha_workforce_acao
        )
        self.btn_abrir_wf.pack(fill=tk.X, padx=12, pady=(0, 6))

        f_btn_sec = ctk.CTkFrame(card_exec, fg_color="transparent")
        f_btn_sec.pack(fill=tk.X, padx=12, pady=(0, 10))

        btn_limpar = ctk.CTkButton(
            f_btn_sec,
            text="🧹 Limpar",
            font=("Segoe UI", 10),
            fg_color="#334155",
            hover_color="#475569",
            height=28,
            command=self.limpar_tudo
        )
        btn_limpar.pack(side="left", fill="x", expand=True, padx=(0, 4))

        btn_copiar_seriais = ctk.CTkButton(
            f_btn_sec,
            text="📋 Copiar Seriais",
            font=("Segoe UI", 10),
            fg_color="#334155",
            hover_color="#475569",
            height=28,
            command=self.copiar_seriais_acao
        )
        btn_copiar_seriais.pack(side="right", fill="x", expand=True, padx=(4, 0))

        # ==========================================
        # COLUNA DIREITA: STATUS, GRID & LOGS
        # ==========================================
        col_dir = ctk.CTkFrame(frame_corpo, fg_color="transparent")
        col_dir.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Card de Resumo de Ordens e Docs
        card_badges = ctk.CTkFrame(col_dir, corner_radius=8, border_color="#334155", border_width=1, height=65)
        card_badges.pack(fill=tk.X, pady=(0, 10))
        card_badges.pack_propagate(False)

        f_badges_grid = ctk.CTkFrame(card_badges, fg_color="transparent")
        f_badges_grid.pack(fill="both", expand=True, padx=12, pady=6)
        f_badges_grid.columnconfigure((0, 1, 2, 3), weight=1)

        # Badge 1: Ordem IW39
        f_b1 = ctk.CTkFrame(f_badges_grid, fg_color="#1e293b", corner_radius=6)
        f_b1.grid(row=0, column=0, sticky="nsew", padx=4)
        ctk.CTkLabel(f_b1, text="ORDEM ATIVA (IW39)", font=("Segoe UI", 8, "bold"), text_color="#94a3b8").pack(pady=(4, 0))
        self.lbl_badge_iw39 = ctk.CTkLabel(f_b1, text="—", font=("Consolas", 12, "bold"), text_color="#38bdf8")
        self.lbl_badge_iw39.pack(pady=(0, 4))

        # Badge 2: Status Auditoria ZMCSE01
        f_b2 = ctk.CTkFrame(f_badges_grid, fg_color="#1e293b", corner_radius=6)
        f_b2.grid(row=0, column=1, sticky="nsew", padx=4)
        ctk.CTkLabel(f_b2, text="AUDITORIA ZMCSE01", font=("Segoe UI", 8, "bold"), text_color="#94a3b8").pack(pady=(4, 0))
        self.lbl_badge_zmcse = ctk.CTkLabel(f_b2, text="Pendente", font=("Segoe UI", 11, "bold"), text_color="#f59e0b")
        self.lbl_badge_zmcse.pack(pady=(0, 4))

        # Badge 3: Doc Z62
        f_b3 = ctk.CTkFrame(f_badges_grid, fg_color="#1e293b", corner_radius=6)
        f_b3.grid(row=0, column=2, sticky="nsew", padx=4)
        ctk.CTkLabel(f_b3, text="DOC SAP Z62 (ESTORNO)", font=("Segoe UI", 8, "bold"), text_color="#94a3b8").pack(pady=(4, 0))
        self.lbl_badge_z62 = ctk.CTkLabel(f_b3, text="—", font=("Consolas", 12, "bold"), text_color="#cbd5e1")
        self.lbl_badge_z62.pack(pady=(0, 4))

        # Badge 4: Doc Z61
        f_b4 = ctk.CTkFrame(f_badges_grid, fg_color="#1e293b", corner_radius=6)
        f_b4.grid(row=0, column=3, sticky="nsew", padx=4)
        ctk.CTkLabel(f_b4, text="DOC SAP Z61 (SAÍDA)", font=("Segoe UI", 8, "bold"), text_color="#94a3b8").pack(pady=(4, 0))
        self.lbl_badge_z61 = ctk.CTkLabel(f_b4, text="—", font=("Consolas", 12, "bold"), text_color="#cbd5e1")
        self.lbl_badge_z61.pack(pady=(0, 4))

        # Tabela Treeview dos Itens Auditados
        card_tabela = ctk.CTkFrame(col_dir, corner_radius=8, border_color="#334155", border_width=1)
        card_tabela.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        lbl_tab_tit = ctk.CTkLabel(card_tabela, text="📋 ITENS AUDITADOS & ORDENS NO SAP:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8")
        lbl_tab_tit.pack(anchor="w", padx=10, pady=(6, 2))

        colunas = ("serial", "material", "ordem_origem", "ordem_destino", "status_zmcse", "resultado")
        self.tree = ttk.Treeview(card_tabela, columns=colunas, show="headings", height=8)

        self.tree.heading("serial", text="Número de Série")
        self.tree.heading("material", text="Material")
        self.tree.heading("ordem_origem", text="Ordem Origem (OSME)")
        self.tree.heading("ordem_destino", text="Ordem Destino (IW39)")
        self.tree.heading("status_zmcse", text="Status ZMCSE01")
        self.tree.heading("resultado", text="Resultado Final")

        self.tree.column("serial", width=120, anchor="center")
        self.tree.column("material", width=90, anchor="center")
        self.tree.column("ordem_origem", width=130, anchor="center")
        self.tree.column("ordem_destino", width=130, anchor="center")
        self.tree.column("status_zmcse", width=140, anchor="center")
        self.tree.column("resultado", width=140, anchor="center")

        scroll_y = ttk.Scrollbar(card_tabela, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(0, 8))
        scroll_y.pack(side="right", fill="y", padx=(0, 8), pady=(0, 8))

        # Terminal de Logs
        card_log = ctk.CTkFrame(col_dir, corner_radius=8, border_color="#334155", border_width=1, height=130)
        card_log.pack(fill=tk.X)

        f_log_top = ctk.CTkFrame(card_log, fg_color="transparent")
        f_log_top.pack(fill=tk.X, padx=10, pady=(4, 2))

        ctk.CTkLabel(f_log_top, text="💻 TERMINAL DE PROCESSAMENTO SAP:", font=("Segoe UI", 9, "bold"), text_color="#94a3b8").pack(side="left")

        btn_limpar_log = ctk.CTkButton(
            f_log_top, text="Limpar Log", font=("Segoe UI", 9), height=20, width=70,
            fg_color="#334155", hover_color="#475569", command=self.limpar_log
        )
        btn_limpar_log.pack(side="right")

        self.txt_log = ctk.CTkTextbox(card_log, font=("Consolas", 10), height=95)
        self.txt_log.pack(fill="both", expand=True, padx=8, pady=(0, 6))

    def _ao_alterar_seriais(self, event=None):
        texto = self.txt_seriais.get("1.0", tk.END).strip()
        seriais = extrair_lista_seriais(texto)
        qtd = len(seriais)
        self.lbl_qtd_seriais.configure(text=f"{qtd} medidor(es)")

    def obter_seriais_digitados(self):
        texto = self.txt_seriais.get("1.0", tk.END).strip()
        return extrair_lista_seriais(texto)

    def obter_equipe(self):
        return self.entry_equipe.get().strip().upper()

    def obter_data_lancamento(self):
        """Retorna a data de lançamento informada na interface no formato dd.mm.aaaa, ou '' se em branco."""
        val = self.entry_data_lancamento.get().strip() if hasattr(self, "entry_data_lancamento") else ""
        if not val:
            return ""
        if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", val):
            messagebox.showwarning(
                "Data de Lançamento Inválida",
                f"A data de lançamento informada ('{val}') está em formato incorreto.\n\n"
                f"Use o formato dd.mm.aaaa (ex: 01.10.2026) ou deixe o campo vazio para utilizar a data padrão de hoje."
            )
            return None
        try:
            d, m, y = map(int, val.split("."))
            if not (1 <= d <= 31 and 1 <= m <= 12 and 2000 <= y <= 2099):
                raise ValueError()
        except ValueError:
            messagebox.showwarning(
                "Data Inválida",
                f"A data de lançamento '{val}' não corresponde a uma data válida de calendário.\n\n"
                f"Informe no formato dd.mm.aaaa (ex: 01.10.2026)."
            )
            return None
        return val

    def remover_divergentes_acao(self):
        """Remove os medidores divergentes da caixa de texto e mantém apenas os aptos."""
        if not getattr(self, "itens_aptos_auditados", None):
            messagebox.showinfo("Informação", "Nenhum medidor apto disponível. Execute a auditoria ZMCSE01 primeiro.")
            return

        qtd_aptos = len(self.itens_aptos_auditados)
        self.txt_seriais.delete("1.0", tk.END)
        self.txt_seriais.insert("1.0", "\n".join(self.itens_aptos_auditados))
        self._ao_alterar_seriais()

        # Atualiza itens em memória e a Treeview para manter apenas os aptos
        itens_filtrados = [it for it in self.itens_carregados if it.get("apto")]
        self.itens_carregados = itens_filtrados

        for item in self.tree.get_children():
            self.tree.delete(item)

        for it in itens_filtrados:
            self.tree.insert(
                "", tk.END,
                values=(
                    it["serial"],
                    it["material"],
                    it["ordem_origem"] or "—",
                    self.ordem_mais_recente_iw39 or "(Pendente IW39)",
                    it["status_zmcse"],
                    "Apto para Contagem"
                )
            )

        self.btn_remover_divergentes.configure(state="disabled", text="✅ Divergentes Removidos")
        self.lbl_badge_zmcse.configure(text=f"✅ {qtd_aptos}/{qtd_aptos} OK", text_color="#10b981")
        self._status(f"Divergentes removidos! Restam {qtd_aptos} medidores aptos.")
        self._log(f"🧹 [Limpeza] Medidores divergentes removidos da lista. Restam {qtd_aptos} medidor(es) apto(s).")
        messagebox.showinfo(
            "Divergentes Removidos",
            f"Lista de medidores atualizada com sucesso!\n\n"
            f"✅ Restam {qtd_aptos} medidor(es) aptos (LIDI/DISP) prontos para o ciclo SAP."
        )

    def limpar_tudo(self):
        self.entry_equipe.delete(0, tk.END)
        if hasattr(self, "entry_data_lancamento"):
            self.entry_data_lancamento.delete(0, tk.END)
        if hasattr(self, "btn_remover_divergentes"):
            self.btn_remover_divergentes.configure(state="disabled", text="🧹 Remover Divergentes da Lista")
        self.itens_aptos_auditados = []
        self.itens_divergentes_auditados = []
        self.txt_seriais.delete("1.0", tk.END)
        self.lbl_qtd_seriais.configure(text="0 medidores")
        self.itens_carregados.clear()
        self.ordem_mais_recente_iw39 = ""
        self.lbl_badge_iw39.configure(text="—", text_color="#38bdf8")
        self.lbl_badge_zmcse.configure(text="Pendente", text_color="#f59e0b")
        self.lbl_badge_z62.configure(text="—", text_color="#cbd5e1")
        self.lbl_badge_z61.configure(text="—", text_color="#cbd5e1")
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._status("Campos limpos.")
        self._log("🧹 Campos reiniciados.")

    def limpar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.configure(state="disabled")

    def copiar_seriais_acao(self):
        seriais = self.obter_seriais_digitados()
        if not seriais:
            messagebox.showwarning("Aviso", "Nenhum serial disponível para cópia.")
            return
        if copiar_texto_clipboard_resiliente("\n".join(seriais)):
            self._log(f"📋 {len(seriais)} seriais copiados para a área de transferência.")
            messagebox.showinfo("Copiado", f"{len(seriais)} seriais copiados com sucesso!")

    # =========================================================================
    # ROTINAS SAP: EXECUÇÃO VIA THREADS
    # =========================================================================

    def _executar_script_vbs(self, script_str, timeout=None):
        with tempfile.NamedTemporaryFile(mode="w", suffix=".vbs", delete=False, encoding="latin-1") as tf:
            tf.write(script_str)
            caminho_vbs = tf.name

        try:
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            resultado = subprocess.run(
                ["cscript.exe", "//nologo", caminho_vbs],
                capture_output=True,
                text=True,
                timeout=timeout,
                creationflags=creationflags
            )
            return resultado.returncode, (resultado.stdout or "").strip(), (resultado.stderr or "").strip()
        finally:
            try:
                os.remove(caminho_vbs)
            except Exception:
                pass

    # -------------------------------------------------------------------------
    # 1. Entrada Exclusiva de Medidores Bipados
    # -------------------------------------------------------------------------
    def puxar_medidores_osme_acao(self):
        messagebox.showinfo(
            "Contagem de Carro",
            "Na Contagem de Carro, não se puxa medidores da OSME no SAP.\n\n"
            "Bipe ou preencha no campo de texto exclusivamente os medidores contados fisicamente no veículo."
        )

    # -------------------------------------------------------------------------
    # 2. Consultar ZMCSE01 (ZMMCSE001 — Análise de Equipe)
    # -------------------------------------------------------------------------
    def consultar_zmcse01_acao(self):
        seriais = self.obter_seriais_digitados()
        if not seriais:
            messagebox.showwarning("Aviso", "Insira ao menos um serial de medidor para auditar na ZMCSE01.")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP GUI precisa estar aberto e conectado.")
            return

        if self.executando:
            messagebox.showwarning("Aviso", "Existe uma automação em andamento no SAP.")
            return

        self.executando = True
        self.btn_consultar_zmcse.configure(state="disabled", text="⏳ Auditando...")
        self._status(f"Auditando {len(seriais)} seriais na ZMCSE01 (ZMMCSE001)...")
        self._log(f"\n🔬 [ZMCSE01] Iniciando auditoria e verificação de ordens para {len(seriais)} medidores...")

        threading.Thread(target=self._thread_consultar_zmcse, args=(seriais,), daemon=True).start()

    def _thread_consultar_zmcse(self, seriais):
        set_sap_ocupado(True)
        try:
            dados_sap = self._sap_consultar_zmmcse001(seriais)
            self.parent_frame.after(0, lambda: self._concluir_consultar_zmcse(dados_sap, seriais))
        except Exception as e:
            self.parent_frame.after(0, lambda err=e: self._erro_consultar_zmcse(str(err)))
        finally:
            set_sap_ocupado(False)

    def _sap_consultar_zmmcse001(self, seriais_lista):
        """Consulta a transação ZMMCSE001 > Análise de Equipe (BT003) no SAP."""
        copiar_texto_clipboard_resiliente("\r\n".join(seriais_lista))
        primeiro_serial = str(seriais_lista[0]).strip()
        qtd = len(seriais_lista)

        linhas_vbs = [
            'On Error Resume Next',
            'Dim SapGuiAuto, app, connection, session',
            'Set SapGuiAuto = GetObject("SAPGUI")',
            'If Not IsObject(SapGuiAuto) Then',
            '    WScript.Echo "ERRO: SAP GUI não encontrado."',
            '    WScript.Quit 1',
            'End If',
            'Set app = SapGuiAuto.GetScriptingEngine',
            'Set connection = app.Children(0)',
            'Set session = connection.Children(0)',
            'session.findById("wnd[0]").maximize',
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nZMMCSE001"',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 800',
            'Dim max_zmm, loop_zmm',
            'max_zmm = 0',
            'Do While session.Children.Count > 1 And max_zmm < 8',
            '    Err.Clear',
            '    session.findById("wnd[1]/tbar[0]/btn[0]").press',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]").sendVKey 0',
            '    End If',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]").close',
            '    End If',
            '    WScript.Sleep 250',
            '    max_zmm = max_zmm + 1',
            'Loop',
            '',
            'session.findById("wnd[0]/usr/btnBT003").press',
            'WScript.Sleep 600',
            '',
            'session.findById("wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/radRD_EQUI").setFocus',
            'session.findById("wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/radRD_EQUI").select',
            'WScript.Sleep 300',
            '',
            f'session.findById("wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/ctxtS_SERNR-LOW").text = "{primeiro_serial}"',
            f'session.findById("wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/txtS_SERNR-LOW").text = "{primeiro_serial}"',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 200',
            ''
        ]

        seriais_clip = vb_escape_clipboard("\r\n".join(seriais_lista))

        if qtd > 1:
            linhas_vbs.extend([
                'session.findById("wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/btn%_S_SERNR_%_APP_%-VALU_PUSH").press',
                'Dim w_wait_zmm',
                'w_wait_zmm = 0',
                'Do While session.Children.Count < 2 And w_wait_zmm < 25',
                '    WScript.Sleep 150',
                '    w_wait_zmm = w_wait_zmm + 1',
                'Loop',
                f'CreateObject("htmlfile").parentWindow.clipboardData.setData "Text", {seriais_clip}',
                'WScript.Sleep 250',
                'If session.Children.Count > 1 Then',
                '    Err.Clear',
                '    session.findById("wnd[1]/tbar[0]/btn[24]").press',
                '    If Err.Number <> 0 Then Err.Clear',
                '    WScript.Sleep 250',
                '    session.findById("wnd[1]/tbar[0]/btn[8]").press',
                '    If Err.Number <> 0 Then Err.Clear',
                '    WScript.Sleep 350',
                'End If',
                ''
            ])

        linhas_vbs.extend([
            'session.findById("wnd[0]/tbar[1]/btn[8]").press',
            'WScript.Sleep 600',
            '',
            'Dim grid, totalLinhas, l, tentativas',
            'Dim ser_val, aufnr_val, mat_val, parnr_val, st_val, st_sis, st_usr, st_disp, st_mov',
            'Set grid = Nothing',
            'totalLinhas = 0',
            '',
            '\' Aguarda ativamente o grid carregar as linhas (até 25 segundos)',
            'For tentativas = 0 To 50',
            '    WScript.Sleep 400',
            '    On Error Resume Next',
            '    If session.Children.Count > 1 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]/tbar[0]/btn[0]").press',
            '        If Err.Number <> 0 Then',
            '            Err.Clear',
            '            session.findById("wnd[1]").sendVKey 0',
            '        End If',
            '        If Err.Number <> 0 Then',
            '            Err.Clear',
            '            session.findById("wnd[1]").close',
            '        End If',
            '        Err.Clear',
            '        WScript.Sleep 250',
            '    End If',
            '    Set grid = session.findById("wnd[0]/usr/cntlCO_ITEM/shellcont/shell")',
            '    If grid Is Nothing Or Not IsObject(grid) Then',
            '        Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")',
            '    End If',
            '    If IsObject(grid) And Not (grid Is Nothing) Then',
            '        totalLinhas = CInt(grid.RowCount)',
            '        If totalLinhas > 0 Then',
            '            WScript.Sleep 300',
            '            totalLinhas = CInt(grid.RowCount)',
            '            Exit For',
            '        End If',
            '    End If',
            '    Err.Clear',
            'Next',
            '',
            'If IsObject(grid) And Not (grid Is Nothing) Then',
            '    totalLinhas = CInt(grid.RowCount)',
            '    ',
            '    \' Mapeia dinamicamente as colunas por título ou ID técnico no ALV (compatível com Array e Collection)',
            '    Dim colsOrder, colCount, colIdx, cOrder',
            '    colsOrder = grid.ColumnOrder',
            '    colCount = 0',
            '    On Error Resume Next',
            '    colCount = UBound(colsOrder) + 1',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        colCount = colsOrder.Count',
            '    End If',
            '    Err.Clear',
            '    ',
            '    Dim col_ser, col_ord, col_mat, col_rec, col_not, col_st_sis, col_st_usr, col_st_mov',
            '    col_ser = "" : col_ord = "" : col_mat = "" : col_rec = "" : col_not = "" : col_st_sis = "" : col_st_usr = "" : col_st_mov = ""',
            '    ',
            '    If colCount > 0 Then',
            '        For colIdx = 0 To colCount - 1',
            '            cOrder = colsOrder(colIdx)',
            '            Dim t1, cUp, titlesCol, tIdx',
            '            cUp = UCase(Trim(cOrder))',
            '            t1 = ""',
            '            Err.Clear',
            '            Set titlesCol = Nothing',
            '            Set titlesCol = grid.GetColumnTitles(cOrder)',
            '            If Not titlesCol Is Nothing Then',
            '                If titlesCol.Count > 0 Then',
            '                    For tIdx = 0 To titlesCol.Count - 1',
            '                        t1 = t1 & " " & UCase(Trim(CStr(titlesCol(tIdx))))',
            '                    Next',
            '                End If',
            '            End If',
            '            Err.Clear',
            '            If col_ser = "" And (InStr(cUp, "SER") > 0 Or InStr(t1, "SÉRIE") > 0 Or InStr(t1, "SERIE") > 0 Or InStr(t1, "SERIAL") > 0) Then col_ser = cOrder',
            '            If col_ord = "" And (InStr(cUp, "AUFN") > 0 Or InStr(cUp, "ORDEM") > 0 Or InStr(cUp, "OSME") > 0 Or cUp = "ORD" Or cUp = "ORDER" Or InStr(t1, "ORDEM") > 0 Or InStr(t1, "OSME") > 0) Then',
            '                If InStr(cUp, "RECORD") = 0 And InStr(cUp, "BORDER") = 0 And InStr(cUp, "COORD") = 0 And InStr(cUp, "SORT") = 0 Then col_ord = cOrder',
            '            End If',
            '            If col_mat = "" And (InStr(cUp, "MAT") > 0 Or InStr(t1, "MAT") > 0) Then col_mat = cOrder',
            '            If col_rec = "" And (InStr(cUp, "PARNR") > 0 Or InStr(cUp, "EQUIP") > 0 Or InStr(t1, "EQUIPE") > 0 Or InStr(t1, "RECURSO") > 0) Then col_rec = cOrder',
            '            If col_not = "" And (InStr(cUp, "NOT") > 0 Or InStr(cUp, "QMNUM") > 0 Or InStr(t1, "NOTA") > 0) Then col_not = cOrder',
            '            If col_st_sis = "" And (InStr(cUp, "STATUS") > 0 Or InStr(cUp, "STS_SIS") > 0 Or InStr(cUp, "STASYS") > 0 Or InStr(cUp, "STTXT") > 0 Or InStr(cUp, "TXT30") > 0 Or InStr(t1, "SISTEMA") > 0 Or InStr(t1, "STATUS") > 0) Then col_st_sis = cOrder',
            '            If col_st_usr = "" And (InStr(cUp, "STUSER") > 0 Or InStr(cUp, "STS_USR") > 0 Or InStr(t1, "USUÁR") > 0 Or InStr(t1, "USUAR") > 0) Then col_st_usr = cOrder',
            '            If col_st_mov = "" And (InStr(t1, "MERCADORIA") > 0 Or InStr(t1, "ST. MOV") > 0 Or InStr(t1, "MOV. MERC") > 0 Or InStr(cUp, "ST_MOV") > 0 Or InStr(cUp, "BTEXT") > 0) Then',
            '                col_st_mov = cOrder',
            '            End If',
            '            If col_st_mov = "" And InStr(t1, "MOV") > 0 And InStr(t1, "ESTOQUE") = 0 Then',
            '                col_st_mov = cOrder',
            '            End If',
            '        Next',
            '    End If',
            '    If col_ser = "" Then col_ser = "SERIAL"',
            '    If col_ord = "" Then col_ord = "AUFNR"',
            '    If col_mat = "" Then col_mat = "MATNR"',
            '    If col_rec = "" Then col_rec = "PARNR"',
            '    If col_not = "" Then col_not = "QMNUM"',
            '    If col_st_sis = "" Then col_st_sis = "STATUS"',
            '    If col_st_usr = "" Then col_st_usr = "STUSER"',
            '    If col_st_mov = "" Then col_st_mov = "ST_MOV"',
            '    Err.Clear',
            '    Dim vRows',
            '    vRows = 10',
            '    On Error Resume Next',
            '    vRows = CInt(grid.VisibleRowCount)',
            '    If vRows < 1 Then vRows = 10',
            '    Err.Clear',
            '    ',
            '    Dim not_val',
            '    For l = 0 To totalLinhas - 1',
            '        If l Mod vRows = 0 Then',
            '            On Error Resume Next',
            '            grid.firstVisibleRow = l',
            '            WScript.Sleep 50',
            '            Err.Clear',
            '        End If',
            '        ',
            '        On Error Resume Next',
            '        ser_val = Trim(CStr(grid.getCellValue(l, col_ser)))',
            '        If ser_val = "" Then ser_val = GetCellSafe(grid, l, "SERIAL|SERNR|EQUNR|GERAET|EQUIPMENT")',
            '        ',
            '        aufnr_val = Trim(CStr(grid.getCellValue(l, col_ord)))',
            '        If aufnr_val = "" Then aufnr_val = GetCellSafe(grid, l, "AUFNR|ORDEM|NUM_ORDEM|AUFNR_ACT|ORDERID|OSME|ORDEM_SERVICO|NRO_ORDEM|ORD_NUM|AUFNR_V|AUFNR_TXT|NUMERO_ORDEM|ORDEM_ORIGEM")',
            '        ',
            '        mat_val = Trim(CStr(grid.getCellValue(l, col_mat)))',
            '        If mat_val = "" Then mat_val = GetCellSafe(grid, l, "MATNR|MATERIAL|ID_MATERIAL|COD_MATERIAL")',
            '        ',
            '        parnr_val = Trim(CStr(grid.getCellValue(l, col_rec)))',
            '        If parnr_val = "" Then parnr_val = GetCellSafe(grid, l, "PARNR|RECURSO|EQUIPE|PERNR")',
            '        ',
            '        not_val = Trim(CStr(grid.getCellValue(l, col_not)))',
            '        If not_val = "" Then not_val = GetCellSafe(grid, l, "QMNUM|NOTA|NUM_NOTA|NOTA_SERVICO")',
            '        ',
            '        st_sis = Trim(CStr(grid.getCellValue(l, col_st_sis)))',
            '        If st_sis = "" Then st_sis = GetCellSafe(grid, l, "STATUS|STS_SISTEMA|ST_SIS|STASYS|DESC_STATUS|STATUS_TXT|EST_SIST|TXT30|TXT04|STTXT|ESTADO|SITUACAO")',
            '        ',
            '        st_usr = Trim(CStr(grid.getCellValue(l, col_st_usr)))',
            '        If st_usr = "" Then st_usr = GetCellSafe(grid, l, "STUSER|STS_USUARIO|ST_USR|STAUSR|DESC_STUSER|STATUS_USR|EST_USUA|STATUS_U")',
            '        ',
            '        st_mov = Trim(CStr(grid.getCellValue(l, col_st_mov)))',
            '        If st_mov = "" Then st_mov = GetCellSafe(grid, l, "ST_MOV|MOVIMENTO|ST_DISP|ESTOQUE|STATUS_MOV|BWART|TIPO_MOV")',
            '        ',
            '        st_disp = st_mov',
            '        st_val = Trim(st_sis & " " & st_usr & " " & st_disp & " " & st_mov)',
            '        ',
            '        If ser_val <> "" Then',
            '            WScript.Echo "ZMMCSE001_DADO:" & ser_val & "|" & aufnr_val & "|" & mat_val & "|" & parnr_val & "|" & not_val & "|" & st_val & "|" & st_sis & "|" & st_usr & "|" & st_disp & "|" & st_mov',
            '        End If',
            '    Next',
            'End If',
            '',
            'On Error Resume Next',
            'session.findById("wnd[0]/tbar[0]/btn[3]").press',
            'WScript.Sleep 300',
            'session.findById("wnd[0]/tbar[0]/btn[3]").press',
            'WScript.Sleep 300',
            'Dim max_zud, loop_zud',
            'max_zud = 0',
            'Do While session.Children.Count > 1 And max_zud < 8',
            '    Err.Clear',
            '    session.findById("wnd[1]/tbar[0]/btn[0]").press',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]").sendVKey 0',
            '    End If',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]").close',
            '    End If',
            '    WScript.Sleep 200',
            '    max_zud = max_zud + 1',
            'Loop',
            '',
            'Function GetCellSafe(g, r, colList)',
            '    On Error Resume Next',
            '    Dim cols, c, v',
            '    v = ""',
            '    cols = Split(colList, "|")',
            '    For Each c In cols',
            '        If c <> "" Then',
            '            Err.Clear',
            '            v = Trim(CStr(g.getCellValue(r, c)))',
            '            If v <> "" Then Exit For',
            '        End If',
            '    Next',
            '    GetCellSafe = v',
            'End Function'
        ])

        rc, out, err = self._executar_script_vbs("\r\n".join(linhas_vbs), timeout=None)
        dados_sap = {}
        for linha in out.splitlines():
            if "ZMMCSE001_DADO:" in linha:
                partes = linha.split("ZMMCSE001_DADO:")[1].strip().split("|")
                ser_v = partes[0].strip()
                ord_v = partes[1].strip() if len(partes) > 1 else ""
                if ord_v.isdigit():
                    ord_v = ord_v.lstrip("0") or "0"
                mat_v = partes[2].strip() if len(partes) > 2 else ""
                if mat_v.isdigit():
                    mat_v = mat_v.lstrip("0") or "0"
                eq_v = partes[3].strip() if len(partes) > 3 else ""
                not_v = partes[4].strip() if len(partes) > 4 else ""
                st_v = partes[5].strip() if len(partes) > 5 else ""
                st_sis = partes[6].strip() if len(partes) > 6 else ""
                st_usr = partes[7].strip() if len(partes) > 7 else ""
                st_disp = partes[8].strip() if len(partes) > 8 else ""
                st_mov = partes[9].strip() if len(partes) > 9 else ""

                item_info = {
                    "serial": ser_v,
                    "ordem": ord_v,
                    "material": mat_v,
                    "equipe": eq_v,
                    "nota": not_v,
                    "status": st_v,
                    "status_sis": st_sis,
                    "status_usr": st_usr,
                    "status_disp": st_disp,
                    "status_mov": st_mov
                }

                # Indexação múltipla para garantir match 100% com ou sem zeros à esquerda
                dados_sap[ser_v] = item_info
                dados_sap[ser_v.strip().upper()] = item_info
                dados_sap[normalizar_texto(ser_v)] = item_info
                dados_sap[normalizar_serial_chave(ser_v)] = item_info
                if ser_v.strip().isdigit():
                    dados_sap[ser_v.strip().lstrip("0") or "0"] = item_info
                    dados_sap[ser_v.strip().zfill(18)] = item_info

        return dados_sap

    def _concluir_consultar_zmcse(self, dados_sap, seriais_consultados):
        self.executando = False
        self.btn_consultar_zmcse.configure(state="normal", text="🔬 1. Auditar ZMCSE01")

        self.itens_carregados.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        itens_aptos = []
        itens_divergentes = []
        equipe_informada = self.obter_equipe()

        for s in seriais_consultados:
            k_busca = normalizar_serial_chave(s)
            k_norm = normalizar_texto(s)
            sap_info = (
                dados_sap.get(k_busca)
                or dados_sap.get(k_norm)
                or dados_sap.get(s.strip().upper())
                or (dados_sap.get(s.strip().lstrip("0") or "0") if s.strip().isdigit() else None)
                or (dados_sap.get(s.strip().zfill(18)) if s.strip().isdigit() else None)
                or {}
            )
            mat = sap_info.get("material") or "510388"
            if str(mat).strip().isdigit():
                mat = str(mat).strip().lstrip("0") or "0"
            ordem = sap_info.get("ordem", "")
            eq_sap = sap_info.get("equipe", equipe_informada)

            apto, motivo = validar_status_lidi_disp(sap_info)

            if apto:
                if not ordem:
                    status_tree = f"✅ OK ({motivo} — Sem Estorno Z62)"
                else:
                    status_tree = f"✅ OK ({motivo})"
                resultado_tree = "Apto para Contagem"
                itens_aptos.append(s)
            else:
                status_tree = f"🚫 Ignorado ({motivo})"
                resultado_tree = f"Ignorado ({motivo})"
                itens_divergentes.append((s, motivo))

            item_obj = {
                "serial": s,
                "material": mat,
                "ordem_origem": ordem,
                "ordem_destino": self.ordem_mais_recente_iw39 or "",
                "status_zmcse": status_tree,
                "apto": apto,
                "motivo": motivo,
                "equipe_sap": eq_sap
            }
            self.itens_carregados.append(item_obj)

            self.tree.insert(
                "", tk.END,
                values=(
                    s,
                    mat,
                    ordem or "—",
                    self.ordem_mais_recente_iw39 or "(Pendente IW39)",
                    status_tree,
                    resultado_tree
                )
            )

        total = len(seriais_consultados)
        self.itens_aptos_auditados = itens_aptos
        self.itens_divergentes_auditados = itens_divergentes
        if hasattr(self, "btn_remover_divergentes"):
            if itens_divergentes and itens_aptos:
                self.btn_remover_divergentes.configure(
                    state="normal",
                    text=f"🧹 Remover {len(itens_divergentes)} Divergente(s)"
                )
            else:
                self.btn_remover_divergentes.configure(
                    state="disabled",
                    text="🧹 Remover Divergentes da Lista" if not itens_divergentes else "🚫 Nenhum Apto para Manter"
                )
        if itens_aptos:
            if itens_divergentes:
                self.lbl_badge_zmcse.configure(text=f"⚠️ {len(itens_aptos)}/{total} Aptos", text_color="#f59e0b")
                self._status(f"ZMCSE01: {len(itens_aptos)}/{total} aptos. {len(itens_divergentes)} ignorados.")
                self._log(f"⚠️ [ZMCSE01] {len(itens_aptos)} medidor(es) aptos (LIDI/DISP) e {len(itens_divergentes)} ignorados por divergência:")
                for s_div, m_div in itens_divergentes:
                    self._log(f"   🚫 Serial {s_div}: {m_div}")

                detalhes_div = "\n".join([f"• Serial {s_d}: {m_d}" for s_d, m_d in itens_divergentes[:8]])
                if len(itens_divergentes) > 8:
                    detalhes_div += f"\n• ... e mais {len(itens_divergentes) - 8} medidor(es)."

                messagebox.showwarning(
                    "Divergência na Auditoria ZMCSE01",
                    f"Atenção: Foram encontrados medidores com divergência de status!\n\n"
                    f"✅ Aptos (LIDI/DISP): {len(itens_aptos)} un\n"
                    f"🚫 Ignorados: {len(itens_divergentes)} un\n\n"
                    f"Detalhes das divergências:\n{detalhes_div}\n\n"
                    f"ℹ️ Os medidores não aptos (montados, sucatas ou sem LIDI/DISP) serão ignorados na execução do ciclo, procedendo normalmente com os medidores válidos."
                )
            else:
                self.lbl_badge_zmcse.configure(text=f"✅ {len(itens_aptos)}/{total} OK", text_color="#10b981")
                self._status(f"ZMCSE01: Todos os {len(itens_aptos)} medidores validados com sucesso.")
                self._log(f"✅ Auditoria ZMCSE01 concluída: Todos os {len(itens_aptos)} medidores com status válido LIDI / DISP!")
        else:
            self.lbl_badge_zmcse.configure(text="🚫 Nenhum Apto", text_color="#ef4444")
            self._status("Nenhum medidor LIDI/DISP encontrado na ZMCSE01.")
            self._log("🚫 [ZMCSE01] Nenhum dos medidores informados está LIDI / DISP no SAP.")
            messagebox.showwarning(
                "Nenhum Medidor Apto",
                "Nenhum dos medidores informados atende ao critério obrigatório LIDI / DISP na ZMCSE01. Todos foram ignorados."
            )

    def _erro_consultar_zmcse(self, erro_msg):
        self.executando = False
        self.btn_consultar_zmcse.configure(state="normal", text="🔬 1. Auditar ZMCSE01")
        self.lbl_badge_zmcse.configure(text="Erro", text_color="#ef4444")
        self._status("Erro na ZMCSE01.")
        self._log(f"❌ Erro ZMCSE01: {erro_msg}")
        messagebox.showerror("Erro ZMCSE01", f"Falha na consulta ZMCSE01:\n\n{erro_msg}")

    # -------------------------------------------------------------------------
    # 3. Buscar Ordem Mais Recente na IW39
    # -------------------------------------------------------------------------
    def buscar_ordem_iw39_acao(self):
        equipe = self.obter_equipe()
        if not equipe:
            messagebox.showwarning("Aviso", "Informe o código da equipe (ex: CSC300) para buscar a ordem mais recente na IW39.")
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP GUI precisa estar aberto e conectado.")
            return

        if self.executando:
            messagebox.showwarning("Aviso", "Existe uma automação em andamento no SAP.")
            return

        self.executando = True
        self.btn_buscar_iw39.configure(state="disabled", text="⏳ Buscando IW39...")
        self._status(f"Buscando ordem ativa mais recente para {equipe} na IW39...")
        self._log(f"\n🔎 [IW39] Consultando ordens ativas para o recurso '{equipe}'...")

        threading.Thread(target=self._thread_buscar_iw39, args=(equipe,), daemon=True).start()

    def _thread_buscar_iw39(self, equipe):
        set_sap_ocupado(True)
        try:
            ordem = self._sap_buscar_ordem_iw39(equipe)
            self.parent_frame.after(0, lambda: self._concluir_buscar_iw39(ordem, equipe))
        except Exception as e:
            self.parent_frame.after(0, lambda err=e: self._erro_buscar_iw39(str(err)))
        finally:
            set_sap_ocupado(False)

    def _sap_buscar_ordem_iw39(self, equipe):
        vbs = f'''
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
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = "{equipe.lower()}"
session.findById("wnd[0]/usr/ctxtDY_PARNR").setFocus
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 200

session.findById("wnd[0]/usr/chkDY_OFN").selected = True
session.findById("wnd[0]/usr/chkDY_IAR").selected = True
session.findById("wnd[0]/usr/chkDY_MAB").selected = False
session.findById("wnd[0]/usr/chkDY_HIS").selected = False

session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 800

Dim max_iw, loop_iw
max_iw = 0
Do While session.Children.Count > 1 And max_iw < 8
    Err.Clear
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[1]").sendVKey 0
    End If
    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[1]").close
    End If
    WScript.Sleep 200
    max_iw = max_iw + 1
Loop

' Função auxiliar para busca recursiva de controle por padrão de ID (ignora rótulos)
Function FindControlByPattern(parentObj, patternStr)
    Set FindControlByPattern = Nothing
    If parentObj Is Nothing Then Exit Function
    On Error Resume Next
    Dim idx, c, res
    For idx = 0 To parentObj.Children.Count - 1
        Set c = parentObj.Children(CInt(idx))
        If c.Type <> "GuiLabel" And c.Type <> "GuiTitlebar" And InStr(UCase(c.Id), "/LBL") = 0 Then
            If InStr(UCase(c.Id), UCase(patternStr)) > 0 Then
                Set FindControlByPattern = c
                Exit Function
            End If
        End If
        If c.Children.Count > 0 Then
            Set res = FindControlByPattern(c, patternStr)
            If Not res Is Nothing Then
                Set FindControlByPattern = res
                Exit Function
            End If
        End If
    Next
    Err.Clear
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
    Err.Clear
End Function

Dim ordemEncontrada
ordemEncontrada = ""

' CENÁRIO 1: Cabeçalho individual da ordem (ex: CDC400)
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
Err.Clear

' 1.2 Tenta achar pelo controle CAUFVD-AUFNR
If ordemEncontrada = "" Then
    On Error Resume Next
    Dim txtOrd
    Set txtOrd = FindControlByPattern(session.findById("wnd[0]/usr"), "CAUFVD-AUFNR")
    If Not txtOrd Is Nothing Then
        Dim rawOrd
        rawOrd = Trim(txtOrd.text)
        If IsNumeric(rawOrd) And Len(rawOrd) >= 8 Then
            ordemEncontrada = rawOrd
        End If
    End If
    Err.Clear
End If

If ordemEncontrada <> "" Then
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    WScript.Sleep 200
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    Err.Clear
End If

' CENÁRIO 2: Grade ALV (múltiplas ordens, ex: CDC300)
If ordemEncontrada = "" Then
    Dim grid
    Set grid = FindFirstGridView(session.findById("wnd[0]"))
    If Not grid Is Nothing Then
        Dim totalLinhas, r, cOrd, cVenc, bestOrd, bestVenc
        totalLinhas = 0
        On Error Resume Next
        totalLinhas = grid.RowCount
        Err.Clear
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
            Err.Clear
            
            If bestOrd = "" Then
                bestOrd = cOrd
                bestVenc = cVenc
            Else
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
        ordemEncontrada = bestOrd

        On Error Resume Next
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        WScript.Sleep 200
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        Err.Clear
    End If
End If

If ordemEncontrada <> "" Then
    WScript.Echo "ORDEM_IW39:" & ordemEncontrada
Else
    WScript.Echo "NENHUMA_ORDEM"
End If
'''
        rc, out, err = self._executar_script_vbs(vbs, timeout=None)
        for linha in out.splitlines():
            if "ORDEM_IW39:" in linha:
                return linha.split("ORDEM_IW39:")[1].strip()
        return ""

    def _concluir_buscar_iw39(self, ordem, equipe):
        self.executando = False
        self.btn_buscar_iw39.configure(state="normal", text="🔎 2. Buscar IW39")

        if ordem:
            self.ordem_mais_recente_iw39 = ordem
            self.lbl_badge_iw39.configure(text=ordem, text_color="#38bdf8")
            self._status(f"IW39: Ordem mais recente localizada: {ordem}")
            self._log(f"🎯 [IW39] Ordem mais recente da equipe {equipe} localizada com sucesso: {ordem}")

            # Atualiza coluna de ordem_destino na Treeview
            for item_id in self.tree.get_children():
                vals = list(self.tree.item(item_id, "values"))
                if len(vals) >= 4:
                    vals[3] = ordem
                    self.tree.item(item_id, values=vals)

            for it in self.itens_carregados:
                it["ordem_destino"] = ordem
        else:
            self.lbl_badge_iw39.configure(text="Não localizada", text_color="#ef4444")
            self._status(f"IW39: Nenhuma ordem ativa encontrada para {equipe}.")
            self._log(f"⚠️ [IW39] Nenhuma ordem ativa localizada para a equipe {equipe}.")
            messagebox.showwarning("IW39", f"Nenhuma ordem ativa localizada na IW39 para a equipe {equipe}.")

    def _erro_buscar_iw39(self, erro_msg):
        self.executando = False
        self.btn_buscar_iw39.configure(state="normal", text="🔎 2. Buscar IW39")
        self.lbl_badge_iw39.configure(text="Erro", text_color="#ef4444")
        self._status("Erro na IW39.")
        self._log(f"❌ Erro IW39: {erro_msg}")
        messagebox.showerror("Erro IW39", f"Falha ao consultar IW39:\n\n{erro_msg}")

    # -------------------------------------------------------------------------
    # 4. Executar Ciclo Completo (ZMCSE01 ➡️ IW39 ➡️ MIGO Z62 ➡️ MIGO Z61)
    # -------------------------------------------------------------------------
    def executar_ciclo_completo_acao(self):
        equipe = self.obter_equipe()
        seriais = self.obter_seriais_digitados()

        if not equipe:
            messagebox.showwarning("Aviso", "Informe a equipe (ex: CSC300).")
            return

        if not seriais:
            messagebox.showwarning("Aviso", "Insira os números de série dos medidores a processar na contagem.")
            return

        data_lcto = self.obter_data_lancamento()
        if data_lcto is None:
            return

        if not verificar_sap_aberto():
            messagebox.showerror("SAP Fechado", "O SAP GUI precisa estar aberto e conectado.")
            return

        if self.executando:
            messagebox.showwarning("Aviso", "Existe uma automação em andamento no SAP.")
            return

        txt_dt = f"• Data de Lançamento (MIGO): {data_lcto} (Fechamento SAP)\n" if data_lcto else "• Data de Lançamento (MIGO): Padrão (Hoje)\n"

        msg_confirma = (
            f"Deseja iniciar a CONTAGEM DE CARRO completa no SAP?\n\n"
            f"• Equipe: {equipe}\n"
            f"• Total de Medidores Bipados no Veículo: {len(seriais)} un\n"
            f"{txt_dt}\n"
            f"Fluxo executado:\n"
            f"1. Auditoria ZMCSE01 (apenas dos {len(seriais)} medidores bipados no veículo)\n"
            f"2. Filtro de Status LIDI / DISP (ignora divergências e restrições)\n"
            f"3. Localização da Ordem Ativa na IW39\n"
            f"4. Estorno MIGO Z62 (sair das ordens onde estavam alocados)\n"
            f"5. Saída MIGO Z61 (voltar para a OSME ativa da equipe: {equipe})\n\n"
            f"⚠️ Não utilize o mouse ou teclado no SAP durante o processo."
        )
        if not messagebox.askyesno("Confirmar Contagem de Carro SAP", msg_confirma, icon="question"):
            return

        self.executando = True
        self.btn_executar_tudo.configure(state="disabled", text="⏳ Executando Ciclo SAP...")
        self._status(f"Executando Contagem de Carro completa para {equipe}...")
        self._log(f"\n🚀 ========================================================")
        self._log(f"🚀 INICIANDO CICLO COMPLETO DE CONTAGEM DE CARRO — {equipe}")
        if data_lcto:
            self._log(f"📅 Data de Lançamento configurada: {data_lcto}")
        self._log(f"🚀 ========================================================")

        threading.Thread(target=self._thread_ciclo_completo, args=(equipe, seriais, data_lcto), daemon=True).start()

    def _thread_ciclo_completo(self, equipe, seriais, data_lancamento=None):
        set_sap_ocupado(True)
        try:
            # ETAPA 1: Auditoria ZMCSE01
            self.parent_frame.after(0, lambda: self._status("Etapa 1/4: Auditando medidores na ZMCSE01..."))
            self._log_threadsafe(f"🔬 [1/4] Auditando exclusivamente os {len(seriais)} medidores bipados na ZMCSE01...")
            dados_zmcse = self._sap_consultar_zmmcse001(seriais)
            self._log_threadsafe(f"🔍 [ZMCSE01] {len(dados_zmcse)} registro(s) indexado(s) do SAP.")

            # Separação: Itens Aptos (LIDI/DISP) vs Itens Ignorados
            itens_aptos = []
            itens_ignorados = []
            sem_ordem = []

            for s in seriais:
                k_busca = normalizar_serial_chave(s)
                k_norm = normalizar_texto(s)
                sap_info = (
                    dados_zmcse.get(k_busca)
                    or dados_zmcse.get(k_norm)
                    or dados_zmcse.get(s.strip().upper())
                    or (dados_zmcse.get(s.strip().lstrip("0") or "0") if s.strip().isdigit() else None)
                    or (dados_zmcse.get(s.strip().zfill(18)) if s.strip().isdigit() else None)
                    or {}
                )
                mat = sap_info.get("material") or "510388"
                if str(mat).strip().isdigit():
                    mat = str(mat).strip().lstrip("0") or "0"
                ordem_origem = sap_info.get("ordem", "")

                if not ordem_origem:
                    # Se não veio na ZMMCSE001, verifica se já tínhamos em itens_carregados
                    for it in self.itens_carregados:
                        it_s = it.get("serial", "")
                        if (normalizar_serial_chave(it_s) == k_busca or normalizar_texto(it_s) == k_norm) and it.get("ordem_origem"):
                            ordem_origem = it.get("ordem_origem")
                            break

                apto, motivo = validar_status_lidi_disp(sap_info)

                item_info = {
                    "serial": s,
                    "material": mat,
                    "ordem_origem": ordem_origem,
                    "apto": apto,
                    "motivo": motivo
                }

                if apto:
                    itens_aptos.append(item_info)
                else:
                    itens_ignorados.append(item_info)

            # Resumo da Auditoria no Terminal de Logs
            self._log_threadsafe(f"🔬 Auditoria ZMCSE01: {len(itens_aptos)} aptos (LIDI/DISP), {len(itens_ignorados)} ignorados.")

            if itens_ignorados:
                self._log_threadsafe(f"⚠️ [Divergência Filtrada Automaticamente] {len(itens_ignorados)} medidor(es) divergente(s) ignorado(s) do ciclo:")
                for it_ign in itens_ignorados:
                    self._log_threadsafe(f"   🚫 Serial {it_ign['serial']}: {it_ign['motivo']}")

            # Se a consulta na ZMCSE01 não retornou registros no tempo hábil
            if not dados_zmcse:
                raise RuntimeError(
                    f"A consulta na ZMCSE01 (ZMMCSE001) não retornou registros ou demorou para carregar.\n\n"
                    f"Verifique se o SAP exibiu a grade de equipamentos e tente novamente."
                )

            # Se todos os medidores forem inválidos/não LIDI/DISP, interrompe com segurança
            if not itens_aptos:
                self.parent_frame.after(0, lambda: self._atualizar_tree_todos_ignorados(itens_ignorados))
                detalhes = "\n".join([f"• Serial {it['serial']}: {it.get('motivo', 'Não LIDI/DISP')}" for it in itens_ignorados[:5]])
                if len(itens_ignorados) > 5:
                    detalhes += f"\n... e mais {len(itens_ignorados) - 5} medidor(es)."
                raise RuntimeError(
                    f"Todos os {len(seriais)} medidores foram ignorados na ZMCSE01:\n\n{detalhes}\n\n"
                    f"A Contagem de Carro exige medidores novos com status obrigatório LIDI / DISP e ordem de origem."
                )

            # ETAPA 2: Localização da Ordem Ativa mais recente na IW39
            self.parent_frame.after(0, lambda: self._status(f"Etapa 2/4: Localizando ordem IW39 ({len(itens_aptos)} aptos)..."))
            self._log_threadsafe("🔎 [2/4] Buscando ordem ativa mais recente na IW39...")
            ordem_recente = self._sap_buscar_ordem_iw39(equipe)

            if not ordem_recente:
                raise RuntimeError(f"Nenhuma ordem ativa localizada na IW39 para a equipe {equipe}. O processo não pode imputar a Z61.")

            self._log_threadsafe(f"🎯 Ordem mais recente IW39 localizada: {ordem_recente}")
            self.ordem_mais_recente_iw39 = ordem_recente
            self.parent_frame.after(0, lambda o=ordem_recente: self.lbl_badge_iw39.configure(text=o, text_color="#38bdf8"))

            # Log detalhado por medidor
            for it in itens_aptos:
                ord_orig_show = it.get("ordem_origem") or "(Sem ordem origem / Depósito)"
                self._log_threadsafe(
                    f"📋 Medidor {it['serial']}: Material {it['material']} | "
                    f"Ordem Origem (Z62 Estorno): {ord_orig_show} | "
                    f"Ordem Destino (Z61 Saída): {ordem_recente}"
                )

            # Atualiza Treeview exibindo aptos em processamento e ignorados destacados
            self.parent_frame.after(0, lambda: self._atualizar_tree_pos_auditoria(itens_aptos, itens_ignorados, ordem_recente))

            # ETAPA 3 & 4: MIGO Z62 (Estorno) e MIGO Z61 (Saída) — APENAS com medidores aptos!
            self.parent_frame.after(0, lambda: self._status(f"Etapa 3/4: Executando MIGO Z62 para {len(itens_aptos)} medidores..."))
            self._log_threadsafe(f"📦 [3/4] Executando MIGO Z62 (Estorno ao Depósito) para {len(itens_aptos)} medidor(es)...")

            doc_z62, doc_z61 = self._sap_executar_migo_z62_e_z61(itens_aptos, ordem_recente, equipe, data_lancamento=data_lancamento)

            # Conclusão
            self.parent_frame.after(0, lambda: self._concluir_ciclo_completo(doc_z62, doc_z61, equipe, ordem_recente, itens_aptos, itens_ignorados))

        except Exception as e:
            self.parent_frame.after(0, lambda err=e: self._erro_ciclo_completo(str(err)))
        finally:
            set_sap_ocupado(False)

    def _atualizar_tree_todos_ignorados(self, itens_ignorados):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for it in itens_ignorados:
            self.tree.insert(
                "", tk.END,
                values=(
                    it["serial"],
                    it.get("material", "—"),
                    it.get("ordem_origem", "—") or "—",
                    "—",
                    f"🚫 Ignorado ({it.get('motivo', 'Não LIDI/DISP')})",
                    "Cancelado (Não LIDI/DISP)"
                )
            )

    def _atualizar_tree_pos_auditoria(self, itens_aptos, itens_ignorados, ordem_destino):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for it in itens_aptos:
            self.tree.insert(
                "", tk.END,
                values=(
                    it["serial"],
                    it["material"],
                    it["ordem_origem"],
                    ordem_destino,
                    f"✅ OK ({it.get('motivo', 'LIDI/DISP')})",
                    "Processando SAP..."
                )
            )

        for it in itens_ignorados:
            self.tree.insert(
                "", tk.END,
                values=(
                    it["serial"],
                    it.get("material", "—"),
                    it.get("ordem_origem", "—") or "—",
                    "—",
                    f"🚫 Ignorado ({it.get('motivo', 'Não LIDI/DISP')})",
                    "Ignorado (Desconsiderado)"
                )
            )

    def _sap_executar_migo_z62_e_z61(self, itens, ordem_destino, equipe, data_lancamento=None):
        """Gera e executa o VBScript para o estorno Z62 seguido da saída Z61 na MIGO."""
        grupos_z62 = OrderedDict()
        for it in itens:
            ord_orig = (it.get("ordem_origem") or "").strip()
            if not ord_orig:
                continue
            chave = (it["material"], ord_orig, "CL03")
            if chave not in grupos_z62:
                grupos_z62[chave] = []
            grupos_z62[chave].append(it["serial"])

        grupos_z61 = OrderedDict()
        for it in itens:
            mat = it["material"]
            if mat not in grupos_z61:
                grupos_z61[mat] = []
            grupos_z61[mat].append(it["serial"])

        linhas_vbs = [
            'On Error Resume Next',
            'Dim SapGuiAuto, app, connection, session',
            'Set SapGuiAuto = GetObject("SAPGUI")',
            'If Not IsObject(SapGuiAuto) Then',
            '    WScript.Echo "ERRO_SAP: SAP GUI não encontrado."',
            '    WScript.Quit 1',
            'End If',
            'Set app = SapGuiAuto.GetScriptingEngine',
            'Set connection = app.Children(0)',
            'Set session = connection.Children(0)',
            'session.findById("wnd[0]").maximize',
            '',
            'Dim carrier',
            'carrier = "0003"',
            '',
            'Sub GarantirDetalheAberto(cId)',
            '    On Error Resume Next',
            '    Dim detObj',
            '    Set detObj = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0301")',
            '    If detObj Is Nothing Or Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press',
            '        If Err.Number <> 0 Then',
            '            Err.Clear',
            '            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0300/btnBUTTON_DETAIL").press',
            '        End If',
            '        Err.Clear',
            '        WScript.Sleep 350',
            '    End If',
            '    Err.Clear',
            'End Sub',
            '',
            'Sub SelecionarAba(cId, tabId)',
            '    On Error Resume Next',
            '    GarantirDetalheAberto cId',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/" & tabId).select',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        WScript.Sleep 200',
            '        GarantirDetalheAberto cId',
            '        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/" & tabId).select',
            '    End If',
            '    Err.Clear',
            '    WScript.Sleep 250',
            'End Sub',
            '',
            'Sub FecharModais(maxCount)',
            '    On Error Resume Next',
            '    Dim mCount',
            '    mCount = 0',
            '    Do While session.Children.Count > 1 And mCount < maxCount',
            '        Err.Clear',
            '        session.findById("wnd[1]/tbar[0]/btn[0]").press',
            '        If Err.Number <> 0 Then',
            '            Err.Clear',
            '            session.findById("wnd[1]").sendVKey 0',
            '        End If',
            '        If Err.Number <> 0 Then',
            '            Err.Clear',
            '            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press',
            '        End If',
            '        If Err.Number <> 0 Then',
            '            Err.Clear',
            '            session.findById("wnd[1]").close',
            '        End If',
            '        Err.Clear',
            '        WScript.Sleep 300',
            '        mCount = mCount + 1',
            '    Loop',
            '    Err.Clear',
            'End Sub',
            '',
            'Sub PreencherDataLancamento(cId, dtLcto)',
            '    If Trim(dtLcto) = "" Then Exit Sub',
            '    On Error Resume Next',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL").select',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_HEADER:SAPLMIGO:0102/btnBUTTON_HEADER").press',
            '        WScript.Sleep 250',
            '        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL").select',
            '    End If',
            '    Err.Clear',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/ctxtGOHEAD-BUDAT").text = dtLcto',
            '    session.findById("wnd[0]").sendVKey 0',
            '    WScript.Sleep 350',
            '    FecharModais 4',
            '    Err.Clear',
            'End Sub',
            '',
            'Sub PreencherConta(cId, ordVal)',
            '    On Error Resume Next',
            '    SelecionarAba cId, "tabpOK_GOITEM_ACCOUNT"',
            '    Dim fOk, subScrs, blks, sSc, bK',
            '    fOk = False',
            '    subScrs = Array("0345", "0340", "0346", "0341", "0342")',
            '    blks = Array("1003", "1002", "1001", "1004")',
            '    For Each sSc In subScrs',
            '        For Each bK In blks',
            '            Err.Clear',
            '            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:" & sSc & "/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:" & bK & "/ctxtCOBL-AUFNR").text = ordVal',
            '            If Err.Number = 0 Then',
            '                session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & cId & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:" & sSc & "/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:" & bK & "/ctxtCOBL-VORNR_AUF").text = "0010"',
            '                fOk = True',
            '                Exit For',
            '            End If',
            '        Next',
            '        If fOk Then Exit For',
            '    Next',
            '    session.findById("wnd[0]").sendVKey 0',
            '    WScript.Sleep 350',
            '    FecharModais 4',
            'End Sub',
            '',
        ]

        if len(grupos_z62) > 0:
            linhas_vbs.extend([
                '\' =========================================================================',
                '\' ETAPA 1: MIGO Z62 (ESTORNO AO DEPÓSITO DAS ORDENS DE ORIGEM)',
                '\' =========================================================================',
                'session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 1000',
            '',
            'FecharModais 8',
            '',
            '\' Detecção dinâmica do carrier (0003, 0002 ou 0004) no MIGO',
            'On Error Resume Next',
            'carrier = "0003"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            'If Err.Number <> 0 Then',
            '    Err.Clear',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            '    If Err.Number = 0 Then',
            '        carrier = "0002"',
            '    Else',
            '        Err.Clear',
            '        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            '        If Err.Number = 0 Then carrier = "0004"',
            '    End If',
            'End If',
            '',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A01"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").setFocus',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R10"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").text = "Z62"',
            'If Err.Number <> 0 Then',
            '    Err.Clear',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODYNPRO-BWART").text = "Z62"',
            'End If',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").setFocus',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 500',
        ])

            dt_lcto_limpa = str(data_lancamento or "").strip()
            if dt_lcto_limpa:
                linhas_vbs.extend([
                    f'PreencherDataLancamento carrier, "{dt_lcto_limpa}"',
                    'WScript.Sleep 300',
                ])

            linhas_vbs.extend([
                '',
                'GarantirDetalheAberto carrier',
                ''
            ])

            # Itens Z62
            idx_z62 = 0
            total_grupos_z62 = len(grupos_z62)
            for (mat_cod, osme_origem, dep_origem), seriais_list in grupos_z62.items():
                qtd = str(len(seriais_list))
                dep_clean = str(dep_origem or "CL03").strip()
                recebedor = f"EMPR 1{dep_clean}"
                osme_origem_clean = str(osme_origem or "").strip()
                if osme_origem_clean.isdigit():
                    osme_origem_clean = osme_origem_clean.lstrip("0") or "0"

                seriais_clip = vb_escape_clipboard("\r\n".join(seriais_list))

                linhas_vbs.extend([
                    'On Error Resume Next',
                    'SelecionarAba carrier, "tabpOK_GOITEM_MATERIAL"',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").text = "{mat_cod}"',
                    'If Err.Number <> 0 Then',
                    '    Err.Clear',
                    f'    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MATNR").text = "{mat_cod}"',
                    'End If',
                    'session.findById("wnd[0]").sendVKey 0',
                    'WScript.Sleep 300',
                    'SelecionarAba carrier, "tabpOK_GOITEM_QUANTITIES"',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = "{qtd}"',
                    'session.findById("wnd[0]").sendVKey 0',
                    'WScript.Sleep 300',
                    'SelecionarAba carrier, "tabpOK_GOITEM_DESTINAT."',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-NAME1").text = "EMPR"',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").text = "{dep_clean}"',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-WEMPF").text = "{recebedor}"',
                    'session.findById("wnd[0]").sendVKey 0',
                    'WScript.Sleep 300',
                    f'PreencherConta carrier, "{osme_origem_clean}"',
                    'SelecionarAba carrier, "tabpOK_GOITEM_SERIAL"',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_SEARCH").press',
                    'WScript.Sleep 350',
                    'session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press',
                    'w_wait_z62 = 0',
                    'Do While session.Children.Count < 2 And w_wait_z62 < 20',
                    '    WScript.Sleep 150',
                    '    w_wait_z62 = w_wait_z62 + 1',
                    'Loop',
                    f'CreateObject("htmlfile").parentWindow.clipboardData.setData "Text", {seriais_clip}',
                    'WScript.Sleep 200',
                    'If session.Children.Count > 1 Then',
                    '    Err.Clear',
                    '    session.findById("wnd[1]/tbar[0]/btn[24]").press',
                    '    If Err.Number <> 0 Then Err.Clear',
                    '    WScript.Sleep 200',
                    '    session.findById("wnd[1]/tbar[0]/btn[0]").press',
                    '    If Err.Number <> 0 Then Err.Clear',
                    '    WScript.Sleep 200',
                    '    session.findById("wnd[1]/tbar[0]/btn[8]").press',
                    '    If Err.Number <> 0 Then Err.Clear',
                    '    WScript.Sleep 350',
                    'End If',
                    'session.findById("wnd[0]/tbar[1]/btn[8]").press',
                    'WScript.Sleep 500',
                    'On Error Resume Next',
                    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""',
                    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll',
                    'WScript.Sleep 200',
                    'session.findById("wnd[0]/tbar[1]/btn[47]").press',
                    'If Err.Number <> 0 Then Err.Clear',
                    'WScript.Sleep 500',
                    '',
                    '\' Marcar Item OK (confirma o item atual no SAP e evita sobreposição)',
                    'On Error Resume Next',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = True',
                    'session.findById("wnd[0]").sendVKey 0',
                    'WScript.Sleep 400',
                ])

                if idx_z62 + 1 < total_grupos_z62:
                    linhas_vbs.extend([
                        '\' Avançar para o próximo item no MIGO Z62 sem sobrepor o anterior',
                        'On Error Resume Next',
                        f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM").press',
                        'If Err.Number <> 0 Then',
                        '    Err.Clear',
                        '    session.findById("wnd[0]/tbar[1]/btn[13]").press',
                        'End If',
                        'Err.Clear',
                        'WScript.Sleep 500',
                        'GarantirDetalheAberto carrier',
                        'WScript.Sleep 250',
                    ])

                idx_z62 += 1

            linhas_vbs.extend([
                'session.findById("wnd[0]/tbar[1]/btn[7]").press',
                'WScript.Sleep 600',
                'session.findById("wnd[0]/tbar[1]/btn[23]").press',
                '',
                'Dim doc_z62, re_z62, matches_z62, msg_sbar_z62, txtErr_z62, w_migo_z62',
                'doc_z62 = ""',
                'msg_sbar_z62 = ""',
                'txtErr_z62 = ""',
                'Set re_z62 = New RegExp',
                're_z62.Pattern = "\\b(49\\d{8}|50\\d{8}|\\d{10})\\b"',
                '',
                'For w_migo_z62 = 0 To 30',
                '    WScript.Sleep 400',
                '    On Error Resume Next',
                '    msg_sbar_z62 = Trim(session.findById("wnd[0]/sbar").text)',
                '    If msg_sbar_z62 <> "" Then',
                '        Set matches_z62 = re_z62.Execute(msg_sbar_z62)',
                '        If matches_z62.Count > 0 Then',
                '            doc_z62 = matches_z62(0).Value',
                '            Exit For',
                '        End If',
                '    End If',
                '    If session.Children.Count > 1 Then',
                '        Err.Clear',
                '        Dim gridLog_z62',
                '        Set gridLog_z62 = Nothing',
                '        Set gridLog_z62 = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")',
                '        If gridLog_z62 Is Nothing Then Set gridLog_z62 = session.findById("wnd[1]/usr/cntlLOG/shellcont/shell")',
                '        If Not gridLog_z62 Is Nothing Then',
                '            Dim rL62, mL62',
                '            For rL62 = 0 To gridLog_z62.RowCount - 1',
                '                mL62 = Trim(CStr(gridLog_z62.getCellValue(rL62, "MESSAGE")))',
                '                If mL62 = "" Then mL62 = Trim(CStr(gridLog_z62.getCellValue(rL62, "TEXT")))',
                '                If mL62 = "" Then mL62 = Trim(CStr(gridLog_z62.getCellValue(rL62, "MSGTXT")))',
                '                If mL62 <> "" Then',
                '                    If txtErr_z62 <> "" Then txtErr_z62 = txtErr_z62 & " | "',
                '                    txtErr_z62 = txtErr_z62 & "Item " & CStr(gridLog_z62.getCellValue(rL62, "ZEILE")) & ": " & mL62',
                '                End If',
                '            Next',
                '        End If',
                '        If txtErr_z62 = "" Then txtErr_z62 = session.findById("wnd[1]/usr/txtMESSTXT1").text',
                '        If txtErr_z62 = "" Then txtErr_z62 = session.findById("wnd[1]").text',
                '        If InStr(UCase(txtErr_z62), "ERRO") > 0 Or InStr(UCase(txtErr_z62), "NÃO") > 0 Or InStr(UCase(txtErr_z62), "BLOQ") > 0 Then',
                '            Exit For',
                '        Else',
                '            session.findById("wnd[1]").sendVKey 0',
                '        End If',
                '    End If',
                '    Dim sbarType_z62',
                '    sbarType_z62 = session.findById("wnd[0]/sbar").messageType',
                '    If UCase(sbarType_z62) = "E" Then',
                '        txtErr_z62 = msg_sbar_z62',
                '        Exit For',
                '    End If',
                'Next',
                '',
                'If doc_z62 <> "" Then',
                '    WScript.Echo "DOC_Z62: " & doc_z62',
                '    FecharModais 8',
                'Else',
                '    If txtErr_z62 = "" Then txtErr_z62 = msg_sbar_z62',
                '    If txtErr_z62 = "" Then txtErr_z62 = "MIGO não gerou documento Z62 (verifique dados no SAP)"',
                '    WScript.Echo "ERRO_SAP: Falha ao capturar Doc Z62: " & txtErr_z62',
                '    WScript.Quit 1',
                'End If',
                '',
            ])
        else:
            linhas_vbs.extend([
                '\' =========================================================================',
                '\' ETAPA 1: MIGO Z62 NÃO REQUERIDA (Nenhum medidor com ordem de origem)',
                '\' =========================================================================',
                'WScript.Echo "DOC_Z62: NÃO REQUERIDO (SEM ESTORNO)"',
                '',
            ])

        linhas_vbs.extend([
            '\' =========================================================================',
            '\' ETAPA 2: MIGO Z61 (SAÍDA PARA A ORDEM MAIS RECENTE DA IW39)',
            '\' =========================================================================',
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 1000',
            '',
            'FecharModais 8',
            '',
            'carrier = "0003"',
            'On Error Resume Next',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            'If Err.Number <> 0 Then',
            '    Err.Clear',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            '    If Err.Number = 0 Then',
            '        carrier = "0002"',
            '    Else',
            '        Err.Clear',
            '        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            '        If Err.Number = 0 Then carrier = "0004"',
            '    End If',
            'End If',
            '',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A07"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").setFocus',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R10"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").text = "Z61"',
            'If Err.Number <> 0 Then',
            '    Err.Clear',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODYNPRO-BWART").text = "Z61"',
            'End If',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").setFocus',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 500',
        ])

        if dt_lcto_limpa:
            linhas_vbs.extend([
                f'PreencherDataLancamento carrier, "{dt_lcto_limpa}"',
                'WScript.Sleep 300',
            ])

        linhas_vbs.extend([
            '',
            'GarantirDetalheAberto carrier',
            ''
        ])

        # Itens Z61
        idx_z61 = 0
        total_grupos_z61 = len(grupos_z61)
        ordem_destino_clean = str(ordem_destino or "").strip()
        if ordem_destino_clean.isdigit():
            ordem_destino_clean = ordem_destino_clean.lstrip("0") or "0"
        for mat_cod, seriais_list in grupos_z61.items():
            qtd = str(len(seriais_list))

            seriais_clip_z61 = vb_escape_clipboard("\r\n".join(seriais_list))

            linhas_vbs.extend([
                'On Error Resume Next',
                'SelecionarAba carrier, "tabpOK_GOITEM_MATERIAL"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").text = "{mat_cod}"',
                'If Err.Number <> 0 Then',
                '    Err.Clear',
                f'    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MATNR").text = "{mat_cod}"',
                'End If',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                'SelecionarAba carrier, "tabpOK_GOITEM_QUANTITIES"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = "{qtd}"',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                'SelecionarAba carrier, "tabpOK_GOITEM_DESTINAT."',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-NAME1").text = "EMPR"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").text = "CL03"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-WEMPF").text = "EMPR 1CL03"',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                f'PreencherConta carrier, "{ordem_destino_clean}"',
                'SelecionarAba carrier, "tabpOK_GOITEM_SERIAL"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_SEARCH").press',
                'WScript.Sleep 350',
                'session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press',
                'w_wait_z61 = 0',
                'Do While session.Children.Count < 2 And w_wait_z61 < 20',
                '    WScript.Sleep 150',
                '    w_wait_z61 = w_wait_z61 + 1',
                'Loop',
                f'CreateObject("htmlfile").parentWindow.clipboardData.setData "Text", {seriais_clip_z61}',
                'WScript.Sleep 200',
                'If session.Children.Count > 1 Then',
                '    Err.Clear',
                '    session.findById("wnd[1]/tbar[0]/btn[24]").press',
                '    If Err.Number <> 0 Then Err.Clear',
                '    WScript.Sleep 200',
                '    session.findById("wnd[1]/tbar[0]/btn[0]").press',
                '    If Err.Number <> 0 Then Err.Clear',
                '    WScript.Sleep 200',
                '    session.findById("wnd[1]/tbar[0]/btn[8]").press',
                '    If Err.Number <> 0 Then Err.Clear',
                '    WScript.Sleep 350',
                'End If',
                'session.findById("wnd[0]/tbar[1]/btn[8]").press',
                'WScript.Sleep 500',
                'On Error Resume Next',
                'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""',
                'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll',
                'WScript.Sleep 200',
                'session.findById("wnd[0]/tbar[1]/btn[47]").press',
                'If Err.Number <> 0 Then Err.Clear',
                'WScript.Sleep 500',
                '',
                '\' Marcar Item OK (confirma o item atual no SAP e evita sobreposição)',
                'On Error Resume Next',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = True',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 400',
            ])

            if idx_z61 + 1 < total_grupos_z61:
                linhas_vbs.extend([
                    '\' Avançar para o próximo item no MIGO Z61 sem sobrepor o anterior',
                    'On Error Resume Next',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM").press',
                    'If Err.Number <> 0 Then',
                    '    Err.Clear',
                    '    session.findById("wnd[0]/tbar[1]/btn[13]").press',
                    'End If',
                    'Err.Clear',
                    'WScript.Sleep 500',
                    'GarantirDetalheAberto carrier',
                    'WScript.Sleep 250',
                ])

            idx_z61 += 1

        linhas_vbs.extend([
            'session.findById("wnd[0]/tbar[1]/btn[7]").press',
            'WScript.Sleep 600',
            'session.findById("wnd[0]/tbar[1]/btn[23]").press',
            '',
            'Dim doc_z61, re_z61, matches_z61, msg_sbar_z61, txtErr_z61, w_migo_z61',
            'doc_z61 = ""',
            'msg_sbar_z61 = ""',
            'txtErr_z61 = ""',
            'Set re_z61 = New RegExp',
            're_z61.Pattern = "\\b(49\\d{8}|50\\d{8}|\\d{10})\\b"',
            '',
            'For w_migo_z61 = 0 To 30',
            '    WScript.Sleep 400',
            '    On Error Resume Next',
            '    msg_sbar_z61 = Trim(session.findById("wnd[0]/sbar").text)',
            '    If msg_sbar_z61 <> "" Then',
            '        Set matches_z61 = re_z61.Execute(msg_sbar_z61)',
            '        If matches_z61.Count > 0 Then',
            '            doc_z61 = matches_z61(0).Value',
            '            Exit For',
            '        End If',
            '    End If',
            '    If session.Children.Count > 1 Then',
            '        Err.Clear',
            '        Dim gridLog_z61',
            '        Set gridLog_z61 = Nothing',
            '        Set gridLog_z61 = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")',
            '        If gridLog_z61 Is Nothing Then Set gridLog_z61 = session.findById("wnd[1]/usr/cntlLOG/shellcont/shell")',
            '        If Not gridLog_z61 Is Nothing Then',
            '            Dim rL61, mL61',
            '            For rL61 = 0 To gridLog_z61.RowCount - 1',
            '                mL61 = Trim(CStr(gridLog_z61.getCellValue(rL61, "MESSAGE")))',
            '                If mL61 = "" Then mL61 = Trim(CStr(gridLog_z61.getCellValue(rL61, "TEXT")))',
            '                If mL61 = "" Then mL61 = Trim(CStr(gridLog_z61.getCellValue(rL61, "MSGTXT")))',
            '                If mL61 <> "" Then',
            '                    If txtErr_z61 <> "" Then txtErr_z61 = txtErr_z61 & " | "',
            '                    txtErr_z61 = txtErr_z61 & "Item " & CStr(gridLog_z61.getCellValue(rL61, "ZEILE")) & ": " & mL61',
            '                End If',
            '            Next',
            '        End If',
            '        If txtErr_z61 = "" Then txtErr_z61 = session.findById("wnd[1]/usr/txtMESSTXT1").text',
            '        If txtErr_z61 = "" Then txtErr_z61 = session.findById("wnd[1]").text',
            '        If InStr(UCase(txtErr_z61), "ERRO") > 0 Or InStr(UCase(txtErr_z61), "NÃO") > 0 Or InStr(UCase(txtErr_z61), "BLOQ") > 0 Then',
            '            Exit For',
            '        Else',
            '            session.findById("wnd[1]").sendVKey 0',
            '        End If',
            '    End If',
            '    Dim sbarType_z61',
            '    sbarType_z61 = session.findById("wnd[0]/sbar").messageType',
            '    If UCase(sbarType_z61) = "E" Then',
            '        txtErr_z61 = msg_sbar_z61',
            '        Exit For',
            '    End If',
            'Next',
            '',
            'If doc_z61 <> "" Then',
            '    WScript.Echo "DOC_Z61: " & doc_z61',
            '    FecharModais 8',
            'Else',
            '    If txtErr_z61 = "" Then txtErr_z61 = msg_sbar_z61',
            '    If txtErr_z61 = "" Then txtErr_z61 = "MIGO não gerou documento Z61 (verifique dados no SAP)"',
            '    WScript.Echo "ERRO_SAP: Falha ao capturar Doc Z61: " & txtErr_z61',
            '    WScript.Quit 1',
            'End If',
            ''
        ])

        rc, out, err = self._executar_script_vbs("\r\n".join(linhas_vbs), timeout=None)
        doc_z62 = ""
        doc_z61 = ""

        for line in out.splitlines():
            if "DOC_Z62:" in line:
                doc_z62 = line.split("DOC_Z62:")[1].strip()
            elif "DOC_Z61:" in line:
                doc_z61 = line.split("DOC_Z61:")[1].strip()
            elif "ERRO_SAP:" in line:
                erro_txt = line.split("ERRO_SAP:")[1].strip()
                raise RuntimeError(erro_txt)

        if not doc_z62 or not doc_z61:
            matches = re.findall(r"\b(49\d{8}|50\d{8}|\d{10})\b", out)
            if len(matches) >= 2:
                doc_z62, doc_z61 = matches[0], matches[1]
            elif len(matches) == 1 and not doc_z62:
                doc_z62 = matches[0]

        if not doc_z62 or not doc_z61:
            raise RuntimeError(f"O SAP não confirmou os documentos Z62 e Z61. Retorno: {out or err}")

        return doc_z62, doc_z61

    def _concluir_ciclo_completo(self, doc_z62, doc_z61, equipe, ordem_recente, itens_aptos, itens_ignorados=None):
        self.executando = False
        self.btn_executar_tudo.configure(
            state="normal",
            text="🚀 Executar Ciclo Completo no SAP\n(ZMCSE01 ➡️ IW39 ➡️ Z62 ➡️ Z61)"
        )

        if doc_z62 == "NÃO REQUERIDO (SEM ESTORNO)":
            self.lbl_badge_z62.configure(text="N/A", text_color="#10b981")
            doc_z62_log = "Não requerido (medidores já no veículo/depósito)"
            doc_z62_status = "N/A"
        else:
            self.lbl_badge_z62.configure(text=doc_z62, text_color="#38bdf8")
            doc_z62_log = doc_z62
            doc_z62_status = doc_z62
        self.lbl_badge_z61.configure(text=doc_z61, text_color="#10b981")

        total_aptos = len(itens_aptos)
        total_ign = len(itens_ignorados) if itens_ignorados else 0

        self._status(f"🎉 Contagem Concluída! Doc Z62: {doc_z62_status} | Doc Z61: {doc_z61} ({total_aptos} un)")
        self._log(f"\n🎉 SUCESSO NA CONTAGEM DE CARRO — {equipe}!")
        self._log(f"   📋 Documento Z62 (Estorno): {doc_z62_log}")
        self._log(f"   📋 Documento Z61 (Saída para Ordem {ordem_recente}): {doc_z61}")
        self._log(f"   ✅ Total de Medidores Processados com Sucesso: {total_aptos} un")
        if total_ign > 0:
            self._log(f"   ⚠️ Total de Medidores Ignorados (Não LIDI/DISP): {total_ign} un")

        # Atualiza resultado na Treeview para os medidores que foram processados com sucesso
        seriais_aptos_set = {normalizar_texto(it["serial"]) for it in itens_aptos}
        for item_id in self.tree.get_children():
            vals = list(self.tree.item(item_id, "values"))
            if len(vals) >= 6:
                if normalizar_texto(vals[0]) in seriais_aptos_set:
                    vals[5] = f"✅ Sucesso (Z61: {doc_z61})"
                    self.tree.item(item_id, values=vals)

        # Gravação no histórico CSV com métricas completas
        self._salvar_historico_contagem(equipe, doc_z62, doc_z61, ordem_recente, itens_aptos, itens_ignorados)
        incrementar_metrica("contagens_carro_executadas", 1)

        # Monta lista de materiais consolidados
        mats_map = OrderedDict()
        for it in itens_aptos:
            m = str(it.get("material", "")).strip()
            if m and m != "—":
                if m.isdigit():
                    m = m.lstrip("0") or "0"
                mats_map[m] = mats_map.get(m, 0) + 1
        materiais_lista = [(m, str(qtd)) for m, qtd in mats_map.items()]

        # Exibe o Pop-up moderno com os documentos e alerta de itens ignorados (se houver)
        ordem_origem_resumo = itens_aptos[0]["ordem_origem"] if itens_aptos else ""
        ModalDocumentosContagem(
            self.parent_frame,
            doc_z62=doc_z62,
            doc_z61=doc_z61,
            equipe=equipe,
            ordem_origem=ordem_origem_resumo,
            ordem_destino=ordem_recente,
            total_medidores=total_aptos,
            itens_ignorados=itens_ignorados,
            materiais_lista=materiais_lista,
            callback_abrir_workforce=lambda: self.abrir_ficha_workforce_acao(equipe, doc_z62, doc_z61, materiais_lista),
            paletas=self.paletas,
            obter_tema_atual=self.obter_tema_atual
        )

        # Dispara abertura automática da ficha no WorkForce em background
        if materiais_lista:
            self._log("\n🌐 Abrindo Ficha no WorkForce com os materiais contados...")
            self.abrir_ficha_workforce_acao(equipe, doc_z62, doc_z61, materiais_lista)

    def _erro_ciclo_completo(self, erro_msg):
        self.executando = False
        self.btn_executar_tudo.configure(
            state="normal",
            text="🚀 Executar Ciclo Completo no SAP\n(ZMCSE01 ➡️ IW39 ➡️ Z62 ➡️ Z61)"
        )
        self._status("Erro durante execução da Contagem.")
        self._log(f"\n❌ ERRO NO CICLO DE CONTAGEM: {erro_msg}")
        messagebox.showerror("Erro na Contagem de Carro", f"Falha na execução do processo SAP:\n\n{erro_msg}")

    def _salvar_historico_contagem(self, equipe, doc_z62, doc_z61, ordem_destino, itens_aptos, itens_ignorados=None):
        try:
            pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
            os.makedirs(pasta, exist_ok=True)
            caminho_csv = os.path.join(pasta, "Historico_Contagem_Carro.csv")

            existe = os.path.exists(caminho_csv)
            with open(caminho_csv, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                if not existe:
                    writer.writerow([
                        "DataHora", "Equipe", "Doc_Z62", "Doc_Z61", "Ordem_Destino",
                        "Qtd_Aptos", "Seriais_Aptos", "Qtd_Ignorados", "Seriais_Ignorados"
                    ])

                seriais_aptos_str = ", ".join([it["serial"] for it in itens_aptos])
                ign_lista = itens_ignorados or []
                seriais_ign_str = ", ".join([f"{it['serial']}({it.get('motivo','')})" for it in ign_lista])

                doc_z62_salvar = "N/A" if doc_z62 == "NÃO REQUERIDO (SEM ESTORNO)" else doc_z62
                writer.writerow([
                    datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    equipe,
                    doc_z62_salvar,
                    doc_z61,
                    ordem_destino,
                    len(itens_aptos),
                    seriais_aptos_str,
                    len(ign_lista),
                    seriais_ign_str
                ])
        except Exception as e:
            self._log(f"⚠️ Aviso ao gravar histórico CSV: {e}")

    def abrir_ficha_workforce_acao(self, equipe=None, doc_z62=None, doc_z61=None, materiais_lista=None):
        """Abre nova ficha no WorkForce preenchendo os materiais da contagem."""
        eq = equipe or (self.entry_equipe.get().strip().upper() if hasattr(self, "entry_equipe") else "")
        re_elet1 = self.entry_re_eletricista.get().strip() if hasattr(self, "entry_re_eletricista") else ""
        re_elet2 = self.entry_re_eletricista_2.get().strip() if hasattr(self, "entry_re_eletricista_2") else ""
        salvar_auto = bool(self.chk_salvar_auto.get()) if hasattr(self, "chk_salvar_auto") else True
        doc62 = doc_z62 or (self.lbl_badge_z62.cget("text").strip() if hasattr(self, "lbl_badge_z62") else "")
        if doc62 == "—":
            doc62 = ""
        doc61 = doc_z61 or (self.lbl_badge_z61.cget("text").strip() if hasattr(self, "lbl_badge_z61") else "")
        if doc61 == "—":
            doc61 = ""

        # Se materiais_lista não foi fornecido, monta a partir da Treeview ou dos seriais digitados
        if not materiais_lista:
            mats_map = OrderedDict()
            for item_id in self.tree.get_children():
                vals = self.tree.item(item_id, "values")
                if len(vals) >= 6:
                    mat_val = str(vals[1]).strip()
                    status_val = str(vals[4]).strip()
                    res_val = str(vals[5]).strip()
                    if "Ignorado" not in status_val and "Ignorado" not in res_val:
                        if mat_val and mat_val != "—":
                            if mat_val.isdigit():
                                mat_val = mat_val.lstrip("0") or "0"
                            mats_map[mat_val] = mats_map.get(mat_val, 0) + 1
            materiais_lista = [(m, str(qtd)) for m, qtd in mats_map.items()]

        if not materiais_lista:
            # Se ainda estiver vazio mas houver seriais digitados, tenta extrair
            seriais_txt = self.txt_seriais.get("1.0", tk.END).strip() if hasattr(self, "txt_seriais") else ""
            if seriais_txt:
                linhas = [s.strip() for s in seriais_txt.splitlines() if s.strip()]
                if linhas:
                    materiais_lista = [("510388", str(len(linhas)))]

        # Normaliza materiais_lista para remover zeros à esquerda se numérico
        if materiais_lista:
            mats_norm = OrderedDict()
            for m, qtd in materiais_lista:
                m_str = str(m).strip()
                if m_str.isdigit():
                    m_str = m_str.lstrip("0") or "0"
                mats_norm[m_str] = mats_norm.get(m_str, 0) + int(qtd)
            materiais_lista = [(m, str(q)) for m, q in mats_norm.items()]

        if not materiais_lista:
            messagebox.showwarning(
                "WorkForce — Contagem de Carro",
                "Nenhum material apto encontrado para preencher na ficha WorkForce.\n"
                "Realize a contagem ou informe os medidores primeiro.",
                parent=self.parent_frame
            )
            return

        creds = carregar_credenciais_workforce() if carregar_credenciais_workforce else {"usuario": "19595", "senha": "979648"}
        usuario_wf = creds.get("usuario", "19595") or "19595"
        senha_wf = creds.get("senha", "979648") or "979648"

        obs_texto = "Contagem de Carro"
        if eq:
            obs_texto += f" — {eq}"
        if doc61:
            obs_texto += f" (Doc Z61: {doc61})"
        elif doc62:
            obs_texto += f" (Doc Z62: {doc62})"

        dados_wf = {
            "data": datetime.now().strftime("%d-%m-%Y"),
            "base": "SANTO ANDRE",
            "re_almox": "19595",
            "re_almox_entregue": "19595",
            "re_eletricista_1": re_elet1,
            "re_eletricista_2": re_elet2,
            "recurso": eq,
            "observacao": obs_texto,
            "materiais": materiais_lista,
            "salvar_auto": salvar_auto
        }

        # Opcionalmente sincroniza com a aba Fichas do N.E.X.U.S se estiver carregada
        try:
            import conecta_hub2
            if getattr(conecta_hub2, "fichas_view", None):
                conecta_hub2.fichas_view.carregar_dados_ficha(dados_wf)
        except Exception:
            pass

        self._status("🌐 Abrindo ficha no WorkForce...")
        self._log(f"\n🌐 Iniciando preenchimento da Ficha no WorkForce...")
        self._log(f"   📅 Data: {dados_wf['data']} | Base: SANTO ANDRE")
        self._log(f"   👤 Resp. Almox.: 19595 | Entregue: 19595")
        self._log(f"   ⚡ Eletricistas: RE1={re_elet1 or 'em branco'}, RE2={re_elet2 or 'em branco'} | Recurso: {eq or 'em branco'}")
        self._log(f"   📦 Materiais: {', '.join([f'{m} ({q} un)' for m, q in materiais_lista])}")
        self._log(f"   💾 Gravação Automática: {'Sim (Salvar Automático)' if salvar_auto else 'Não (Conferência Manual)'}")

        def _cb_status(msg):
            self._log_threadsafe(f"   [WorkForce] {msg}")
            self.parent_frame.after(0, lambda: self._status(msg))

        def _worker():
            try:
                if executar_fluxo_workforce is None:
                    _cb_status("⚠️ Módulo WorkForce não carregado.")
                    return
                mapa_nomes = _carregar_banco_nomes() if _carregar_banco_nomes else {}
                sucesso, res = executar_fluxo_workforce(
                    dados=dados_wf,
                    callback_status=_cb_status,
                    login_usuario=usuario_wf,
                    login_senha=senha_wf,
                    mapa_nomes=mapa_nomes
                )
                if sucesso:
                    msg_sucesso = "Ficha salva e finalizada no WorkForce com sucesso!" if salvar_auto else "Ficha preenchida no WorkForce! Confira e clique em Salvar."
                    self.parent_frame.after(0, lambda: self._status(msg_sucesso))
                    self._log_threadsafe(f"✅ {msg_sucesso}")
                else:
                    self.parent_frame.after(0, lambda: self._status("Atenção ao abrir WorkForce."))
                    self._log_threadsafe(f"⚠️ WorkForce: {res}")
            except Exception as ex:
                self.parent_frame.after(0, lambda: self._status("Erro ao abrir WorkForce."))
                self._log_threadsafe(f"❌ Erro WorkForce: {ex}")

        threading.Thread(target=_worker, daemon=True).start()


def construir_aba_contagem_carro(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    """Factory exportada para criação da aba no Notebook do N.E.X.U.S / Conecta Hub."""
    return AbaContagemCarro(parent_frame, label_status_global, paletas, obter_tema_atual)
