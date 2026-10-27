import os
import re
import json
import time
import subprocess
import tempfile
import threading
from collections import OrderedDict
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from openpyxl import load_workbook
import win32com.client

from utils_os import obter_caminho_config, obter_diretorio_dados, obter_diretorio_executavel
from modulo_extrator import extrair_lista_seriais
from modulo_Sap import verificar_sap_aberto
from utils_pendencias_planilha import (
    adicionar_pendencia,
    obter_quantidade_pendencias,
    processar_todas_pendencias,
    ModalVisualizarPendencias,
    is_file_writable,
    criar_backup_preventivo
)
from utils_metrics import incrementar_metrica
from utils_backup_diario import (
    gravar_saida_backup_diario,
    abrir_backup_diario,
    comparar_backup_com_mestre,
    obter_caminho_backup_diario
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


PADRAO_MODELO_MEDIDOR = re.compile(
    r"^(?:D12|D13|D23|D34|S23|S34|A34|C0006)[A-Z0-9_\-]+$",
    re.IGNORECASE
)
PADRAO_MATERIAL_SAP = re.compile(
    r"^510\d{3}$"
)
PALAVRAS_IGNORAR_SERIAIS = {
    "MATERIAL", "MATERIAIS", "MAT", "COD", "CODIGO", "CÓDIGO",
    "MODELO", "MEDIDOR", "MEDIDORES", "SERIAIS", "SERIAL", "SERIE", "SÉRIE",
    "QTD", "QUANTIDADE", "UN", "UND", "UNID", "UNIDADE", "UNIDADES",
    "CL03", "CL04", "DEPOSITO", "DEPÓSITO", "EQUIPE", "RECURSO",
    "MONOFASICO", "MONOFÁSICO", "BIFASICO", "BIFÁSICO", "TRIFASICO", "TRIFÁSICO",
    "NANSEN", "LANDIS", "ELETRA", "ELSTER", "ITRON", "SCHLUMBERGER", "SMART"
}

_CACHE_MATERIAIS_BANCO = None


def obter_materiais_banco_local():
    """Carrega todos os códigos de materiais e modelos cadastrados em banco_de_dados.xlsx."""
    global _CACHE_MATERIAIS_BANCO
    if _CACHE_MATERIAIS_BANCO is not None:
        return _CACHE_MATERIAIS_BANCO
    mats = set()
    try:
        import pandas as pd
        caminho_banco = os.path.join(obter_diretorio_executavel(), "data", "banco_de_dados.xlsx")
        if not os.path.exists(caminho_banco):
            from utils_os import obter_caminho_recurso
            caminho_banco = obter_caminho_recurso("data", "banco_de_dados.xlsx")
        if os.path.exists(caminho_banco):
            df = pd.read_excel(caminho_banco)
            df.columns = df.columns.str.strip()
            if "Material" in df.columns:
                for val in df["Material"].dropna():
                    s = str(val).strip().upper()
                    if s:
                        mats.add(s)
                        norm = re.sub(r"[^A-Z0-9]", "", s)
                        if norm:
                            mats.add(norm)
    except Exception:
        pass
    _CACHE_MATERIAIS_BANCO = mats
    return _CACHE_MATERIAIS_BANCO


def normalizar_texto(valor):
    if valor is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(valor).strip().upper())


def formatar_data_excel(valor):
    if valor is None:
        return "-"
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    try:
        num = int(valor)
        if 40000 < num < 60000:
            return (datetime(1899, 12, 30) + timedelta(days=num)).strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(valor).strip()


def copiar_texto_clipboard_resiliente(texto):
    """Copia texto para o Clipboard do Windows de forma resiliente."""
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
        import subprocess
        subprocess.run(["clip.exe"], input=str(texto), text=True, check=True, timeout=2)
        return True
    except Exception:
        pass
    return False


def vb_escape_clipboard(texto):
    """Escapa quebras de linha para inserção segura no VBScript clipboard."""
    partes = [p.replace('"', '""') for p in str(texto).split("\r\n")]
    if len(partes) == 1:
        return f'"{partes[0]}"'
    vbs_partes = [f'"{p}"' for p in partes]
    return " & vbCrLf & ".join(vbs_partes)


class ModalAlertaFornecidos(ctk.CTkToplevel):
    """Pop-up modal em destaque avisando que os seriais já constam como fornecidos."""
    def __init__(self, parent, itens_bloqueados, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.itens = itens_bloqueados
        self.title("⛔ ALERTA: Seriais Já Fornecidos Detectados!")
        self.geometry("580x440")
        self.minsize(500, 360)
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.configure(fg_color="#18181b")

        # Cabeçalho de Alerta
        frame_topo = ctk.CTkFrame(self, fg_color="#7f1d1d", corner_radius=0, height=55)
        frame_topo.pack(side="top", fill="x")
        frame_topo.pack_propagate(False)

        lbl_tit = ctk.CTkLabel(
            frame_topo,
            text="⛔ BLOQUEIO: MATERIAIS JÁ FORNECIDOS NA PLANILHA",
            font=("Segoe UI", 13, "bold"),
            text_color="#fecaca"
        )
        lbl_tit.pack(side="left", padx=16, pady=14)

        lbl_msg = ctk.CTkLabel(
            self,
            text=f"Foram identificados {len(self.itens)} seriais que já constam na aba FORNECIDOS da planilha.\n"
                 "A saída desses itens foi bloqueada para evitar duplicidade de baixa no SAP:",
            font=("Segoe UI", 11),
            text_color="#e4e4e7",
            justify="left"
        )
        lbl_msg.pack(anchor="w", padx=16, pady=(12, 8))

        # Lista com rolagem
        scroll = ctk.CTkScrollableFrame(self, fg_color="#27272a", corner_radius=8)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        for it in self.itens:
            card = ctk.CTkFrame(scroll, fg_color="#3f3f46", corner_radius=6)
            card.pack(fill="x", pady=3, padx=2)

            f_lin1 = ctk.CTkFrame(card, fg_color="transparent")
            f_lin1.pack(fill="x", padx=10, pady=(6, 2))

            ctk.CTkLabel(
                f_lin1,
                text=f"Serial: {it.get('serial')}",
                font=("Consolas", 12, "bold"),
                text_color="#f87171"
            ).pack(side="left")

            ctk.CTkLabel(
                f_lin1,
                text=f"Material: {it.get('material', '-')}",
                font=("Segoe UI", 10),
                text_color="#fecaca"
            ).pack(side="right")

            f_lin2 = ctk.CTkFrame(card, fg_color="transparent")
            f_lin2.pack(fill="x", padx=10, pady=(0, 6))

            detalhes = []
            if it.get("doc_sap"):
                detalhes.append(f"Doc SAP: {it.get('doc_sap')}")
            if it.get("data"):
                detalhes.append(f"Data: {it.get('data')}")
            if it.get("equipe"):
                detalhes.append(f"Equipe: {it.get('equipe')}")

            ctk.CTkLabel(
                f_lin2,
                text=" • ".join(detalhes) if detalhes else "Já registrado na aba FORNECIDOS",
                font=("Segoe UI", 10),
                text_color="#a1a1aa"
            ).pack(side="left")

        # Botão Fechar
        btn_fechar = ctk.CTkButton(
            self,
            text="Entendido (Remover ou Corrigir Seriais)",
            font=("Segoe UI", 11, "bold"),
            fg_color="#dc2626",
            hover_color="#b91c1c",
            height=36,
            command=self.destroy
        )
        btn_fechar.pack(fill="x", padx=16, pady=(0, 14))


class ModalCompararBackup(ctk.CTkToplevel):
    """Pop-up modal moderno exibindo o resultado da auditoria entre o backup diário local e a planilha mestre."""
    def __init__(self, parent, resultado, caminho_mestre, callback_recarregar=None, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.resultado = resultado
        self.caminho_mestre = caminho_mestre
        self.callback_recarregar = callback_recarregar
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")

        self.title("🔍 Auditoria de Saídas: Backup Local vs Planilha Mestre")
        self.geometry("640x520")
        self.minsize(560, 420)
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.configure(fg_color="#18181b")

        self._construir_ui()

    def _construir_ui(self):
        sucesso = self.resultado.get("sucesso", False)
        if not sucesso:
            f_topo = ctk.CTkFrame(self, fg_color="#7f1d1d", corner_radius=0, height=55)
            f_topo.pack(side="top", fill="x")
            f_topo.pack_propagate(False)
            ctk.CTkLabel(
                f_topo,
                text="❌ ERRO AO COMPARAR BACKUP COM PLANILHA MESTRE",
                font=("Segoe UI", 12, "bold"),
                text_color="#fecaca"
            ).pack(side="left", padx=16, pady=14)

            lbl_err = ctk.CTkLabel(
                self,
                text=self.resultado.get("erro", "Erro desconhecido ao comparar."),
                font=("Segoe UI", 11),
                text_color="#f87171",
                justify="left",
                wraplength=580
            )
            lbl_err.pack(padx=20, pady=25)

            ctk.CTkButton(
                self,
                text="Fechar",
                font=("Segoe UI", 11, "bold"),
                height=36,
                command=self.destroy
            ).pack(side="bottom", padx=20, pady=16, fill="x")
            return

        total_bk = self.resultado.get("total_backup", 0)
        confirmados = self.resultado.get("confirmados", 0)
        faltantes = self.resultado.get("faltantes", [])

        if len(faltantes) == 0:
            f_topo = ctk.CTkFrame(self, fg_color="#14532d", corner_radius=0, height=55)
            f_topo.pack(side="top", fill="x")
            f_topo.pack_propagate(False)
            ctk.CTkLabel(
                f_topo,
                text="✅ 100% AUDITADO E SINCRONIZADO COM A REDE!",
                font=("Segoe UI", 13, "bold"),
                text_color="#bbf7d0"
            ).pack(side="left", padx=16, pady=14)

            f_corpo = ctk.CTkFrame(self, fg_color="transparent")
            f_corpo.pack(fill="both", expand=True, padx=20, pady=20)

            lbl_grande = ctk.CTkLabel(
                f_corpo,
                text=f"🎉 Todas as {total_bk} saídas de seriais registradas no seu backup diário de hoje constam com sucesso na aba FORNECIDOS da planilha mestre!",
                font=("Segoe UI", 12),
                text_color="#f8fafc",
                justify="left",
                wraplength=580
            )
            lbl_grande.pack(pady=(15, 20), anchor="w")

            card_info = ctk.CTkFrame(f_corpo, fg_color="#27272a", corner_radius=8)
            card_info.pack(fill="x", pady=10)

            ctk.CTkLabel(
                card_info,
                text=f"📁 Arquivo de Backup Local: {os.path.basename(self.resultado.get('caminho_backup', ''))}\n"
                     f"🌐 Planilha Mestre (Rede): {os.path.basename(self.resultado.get('caminho_mestre', ''))}\n"
                     f"📊 Total Conferido: {confirmados} itens validados com êxito.",
                font=("Consolas", 10),
                text_color="#94a3b8",
                justify="left"
            ).pack(padx=14, pady=12, anchor="w")

            f_botoes = ctk.CTkFrame(self, fg_color="transparent")
            f_botoes.pack(side="bottom", fill="x", padx=20, pady=16)

            ctk.CTkButton(
                f_botoes,
                text="📁 Abrir Backup Diário no Excel",
                font=("Segoe UI", 11, "bold"),
                fg_color="#059669",
                hover_color="#047857",
                height=36,
                command=lambda: [abrir_backup_diario(), self.destroy()]
            ).pack(side="left", fill="x", expand=True, padx=(0, 8))

            ctk.CTkButton(
                f_botoes,
                text="Fechar",
                font=("Segoe UI", 11),
                height=36,
                command=self.destroy
            ).pack(side="right", fill="x", expand=True)

        else:
            f_topo = ctk.CTkFrame(self, fg_color="#78350f", corner_radius=0, height=55)
            f_topo.pack(side="top", fill="x")
            f_topo.pack_propagate(False)
            ctk.CTkLabel(
                f_topo,
                text=f"⚠️ DIVERGÊNCIA: {len(faltantes)} SAÍDA(S) NÃO GRAVADA(S) NA REDE",
                font=("Segoe UI", 13, "bold"),
                text_color="#fef3c7"
            ).pack(side="left", padx=16, pady=14)

            lbl_sub = ctk.CTkLabel(
                self,
                text=f"Foram identificados {len(faltantes)} seriais no seu backup local que ainda NÃO constam na aba FORNECIDOS da planilha mestre.\n"
                     "Você pode sincronizá-los agora mesmo clicando no botão verde abaixo:",
                font=("Segoe UI", 11),
                text_color="#f8fafc",
                justify="left"
            )
            lbl_sub.pack(anchor="w", padx=20, pady=(12, 8))

            scroll = ctk.CTkScrollableFrame(self, fg_color="#27272a", corner_radius=8)
            scroll.pack(fill="both", expand=True, padx=20, pady=(0, 12))

            for it in faltantes:
                card = ctk.CTkFrame(scroll, fg_color="#3f3f46", corner_radius=6)
                card.pack(fill="x", pady=3, padx=2)

                f_l1 = ctk.CTkFrame(card, fg_color="transparent")
                f_l1.pack(fill="x", padx=10, pady=(6, 2))
                ctk.CTkLabel(f_l1, text=f"Serial: {it['serial']}", font=("Consolas", 12, "bold"), text_color="#fbbf24").pack(side="left")
                ctk.CTkLabel(f_l1, text=f"Equipe: {it.get('equipe', '-')}", font=("Segoe UI", 10, "bold"), text_color="#38bdf8").pack(side="right")

                f_l2 = ctk.CTkFrame(card, fg_color="transparent")
                f_l2.pack(fill="x", padx=10, pady=(0, 6))
                ctk.CTkLabel(
                    f_l2,
                    text=f"Material: {it.get('material', '-')}  |  Doc SAP: {it.get('doc_sap', '-')}  |  Dep: {it.get('deposito', 'CL03')}  |  Hora: {it.get('hora_nexus', '-')}",
                    font=("Segoe UI", 10),
                    text_color="#cbd5e1"
                ).pack(side="left")

            f_botoes = ctk.CTkFrame(self, fg_color="transparent")
            f_botoes.pack(side="bottom", fill="x", padx=20, pady=(0, 14))

            ctk.CTkButton(
                f_botoes,
                text="🚀 Gravar Faltantes na Rede Agora",
                font=("Segoe UI", 11, "bold"),
                fg_color="#10b981",
                hover_color="#059669",
                height=38,
                command=self._gravar_faltantes
            ).pack(side="left", fill="x", expand=True, padx=(0, 8))

            ctk.CTkButton(
                f_botoes,
                text="📁 Abrir Backup",
                font=("Segoe UI", 11),
                height=38,
                command=abrir_backup_diario
            ).pack(side="left", padx=(0, 8))

            ctk.CTkButton(
                f_botoes,
                text="Fechar",
                font=("Segoe UI", 11),
                height=38,
                command=self.destroy
            ).pack(side="right")

    def _gravar_faltantes(self):
        faltantes = self.resultado.get("faltantes", [])
        if not faltantes:
            return

        from utils_pendencias_planilha import tentar_gravar_saida_osme, adicionar_pendencia
        grupos = {}
        for it in faltantes:
            chave = (it.get("equipe") or "-", it.get("doc_sap") or "-", it.get("data_lcto") or "-")
            if chave not in grupos:
                grupos[chave] = []
            grupos[chave].append(it)

        sucessos = 0
        pendencias = 0
        for (eq, doc, dt), lista_itens in grupos.items():
            itens_fmt = [{"serial": x["serial"], "material": x["material"], "deposito": x.get("deposito", "CL03")} for x in lista_itens]
            ok, metodo = tentar_gravar_saida_osme(self.caminho_mestre, eq, doc, dt, itens_fmt)
            if ok:
                sucessos += len(lista_itens)
            else:
                adicionar_pendencia(
                    tipo="saida_osme",
                    caminho_planilha=self.caminho_mestre,
                    dados_operacao={
                        "equipe": eq,
                        "doc_sap": doc,
                        "data": dt,
                        "deposito": lista_itens[0].get("deposito", "CL03"),
                        "itens": itens_fmt
                    },
                    motivo="Sincronização manual via Auditoria de Backup"
                )
                pendencias += len(lista_itens)

        msg = []
        if sucessos > 0:
            msg.append(f"🎉 {sucessos} item(ns) gravados diretamente na planilha mestre!")
        if pendencias > 0:
            msg.append(f"⏳ {pendencias} item(ns) adicionados à Fila de Pendências (planilha em uso na rede).")

        messagebox.showinfo("Sincronização de Backup", "\n\n".join(msg), parent=self)
        if self.callback_recarregar:
            self.callback_recarregar()
        self.destroy()


class ModalVisualizarWorkforce(ctk.CTkToplevel):
    """Pop-up modal exibindo os seriais agrupados por material no formato WorkForce com ';'."""
    def __init__(self, parent, texto_workforce, totais_materiais, paletas=None, obter_tema_atual=None):
        super().__init__(parent)
        self.texto_workforce = texto_workforce
        self.totais_materiais = totais_materiais
        self.title("📋 Lista de Seriais para WorkForce")
        self.geometry("500x580")
        self.minsize(440, 480)
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.configure(fg_color="#18181b")

        # Topo
        frame_topo = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0, height=55)
        frame_topo.pack(side="top", fill="x")
        frame_topo.pack_propagate(False)

        lbl_tit = ctk.CTkLabel(
            frame_topo,
            text="📋 LISTA FORMATADA — WORKFORCE",
            font=("Segoe UI", 13, "bold"),
            text_color="#38bdf8"
        )
        lbl_tit.pack(side="left", padx=16, pady=14)

        total_seriais = sum(len(itens) for itens in totais_materiais.values()) if totais_materiais else 0
        lbl_sub = ctk.CTkLabel(
            self,
            text=f"Total: {total_seriais} seriais separados por material (1 por linha terminando com ';'):",
            font=("Segoe UI", 11),
            text_color="#cbd5e1"
        )
        lbl_sub.pack(anchor="w", padx=16, pady=(10, 6))

        # Textbox com a lista
        self.txt = ctk.CTkTextbox(self, font=("Consolas", 12), fg_color="#27272a", corner_radius=8)
        self.txt.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        self.txt.insert("1.0", self.texto_workforce)
        self.txt.configure(state="disabled")

        # Botões
        f_botoes = ctk.CTkFrame(self, fg_color="transparent")
        f_botoes.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_copiar = ctk.CTkButton(
            f_botoes,
            text="📋 Copiar Tudo para a Área de Transferência",
            font=("Segoe UI", 12, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=38,
            command=self._copiar
        )
        self.btn_copiar.pack(side="left", fill="x", expand=True, padx=(0, 8))

        btn_fechar = ctk.CTkButton(
            f_botoes,
            text="Fechar",
            width=80,
            height=38,
            fg_color="transparent",
            border_width=1,
            command=self.destroy
        )
        btn_fechar.pack(side="left")

    def _copiar(self):
        self.clipboard_clear()
        self.clipboard_append(self.texto_workforce)
        self.btn_copiar.configure(text="✓ Copiado com Sucesso!", fg_color="#10b981")
        self.after(2000, lambda: self.btn_copiar.configure(text="📋 Copiar Tudo para a Área de Transferência", fg_color="#0284c7"))


class ModalSucessoSaidaOSME(ctk.CTkToplevel):
    """Pop-up modal moderno com o Documento SAP em destaque, botão de cópia rápida e integração com o WorkForce."""
    def __init__(self, parent, doc_sap, equipe, osme, todos_itens, sucesso_planilha, modo_planilha, paletas=None, obter_tema_atual=None, totais_materiais=None):
        super().__init__(parent)
        self.doc_sap = str(doc_sap).strip()
        self.equipe = str(equipe).strip()
        self.osme = str(osme).strip()
        self.todos_itens = todos_itens or []
        self.sucesso_planilha = sucesso_planilha
        self.modo_planilha = modo_planilha

        # Agrupa materiais caso não tenham sido passados diretamente
        if not totais_materiais and self.todos_itens:
            grupos = {}
            for it in self.todos_itens:
                m = str(it.get("material", "—")).strip()
                grupos.setdefault(m, []).append(it)
            self.totais_materiais = grupos
        else:
            self.totais_materiais = totais_materiais or {}

        self.title("🎉 Saída SAP Concluída com Sucesso!")
        self.geometry("540x630")
        self.minsize(500, 560)
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.configure(fg_color="#18181b")

        # Copia automaticamente para o clipboard na inicialização para economia de tempo!
        try:
            self.clipboard_clear()
            self.clipboard_append(self.doc_sap)
        except Exception:
            pass

        # Topo de Sucesso
        frame_topo = ctk.CTkFrame(self, fg_color="#064e3b", corner_radius=0, height=55)
        frame_topo.pack(side="top", fill="x")
        frame_topo.pack_propagate(False)

        lbl_tit = ctk.CTkLabel(
            frame_topo,
            text="🎉 SAÍDA NO SAP REALIZADA COM SUCESSO!",
            font=("Segoe UI", 13, "bold"),
            text_color="#a7f3d0"
        )
        lbl_tit.pack(side="left", padx=16, pady=14)

        # Card Principal com o Documento SAP em Destaque
        card_doc = ctk.CTkFrame(self, fg_color="#1e293b", corner_radius=10, border_color="#10b981", border_width=2)
        card_doc.pack(fill="x", padx=20, pady=(16, 10))

        ctk.CTkLabel(
            card_doc,
            text="DOCUMENTO SAP MIGO GERADO:",
            font=("Segoe UI", 11, "bold"),
            text_color="#94a3b8"
        ).pack(pady=(10, 2))

        self.lbl_numero = ctk.CTkLabel(
            card_doc,
            text=self.doc_sap,
            font=("Consolas", 28, "bold"),
            text_color="#34d399"
        )
        self.lbl_numero.pack(pady=(0, 8))

        self.btn_copiar = ctk.CTkButton(
            card_doc,
            text="📋 Copiar Documento SAP",
            font=("Segoe UI", 12, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=34,
            command=self._copiar_doc
        )
        self.btn_copiar.pack(padx=20, pady=(0, 12), fill="x")

        # Card de Resumo da Operação
        card_detalhes = ctk.CTkFrame(self, fg_color="#27272a", corner_radius=8)
        card_detalhes.pack(fill="x", padx=20, pady=(0, 10))

        f_linha1 = ctk.CTkFrame(card_detalhes, fg_color="transparent")
        f_linha1.pack(fill="x", padx=14, pady=(8, 2))
        ctk.CTkLabel(f_linha1, text=f"• Equipe / Recurso: {self.equipe}", font=("Segoe UI", 11, "bold"), text_color="#f1f5f9").pack(side="left")
        ctk.CTkLabel(f_linha1, text=f"OSME: {self.osme}", font=("Consolas", 11, "bold"), text_color="#38bdf8").pack(side="right")

        f_linha2 = ctk.CTkFrame(card_detalhes, fg_color="transparent")
        f_linha2.pack(fill="x", padx=14, pady=(2, 6))
        ctk.CTkLabel(f_linha2, text=f"• Total de Itens Baixados: {len(self.todos_itens)} item(ns)", font=("Segoe UI", 11), text_color="#cbd5e1").pack(side="left")

        # Status Planilha
        f_plan = ctk.CTkFrame(card_detalhes, fg_color="#18181b", corner_radius=6)
        f_plan.pack(fill="x", padx=14, pady=(2, 8))

        if self.sucesso_planilha:
            icone_plan = "✓"
            txt_plan = f"Planilha atualizada com sucesso ({self.modo_planilha})\nOs seriais foram transferidos de ESTOQUE para FORNECIDOS."
            cor_plan = "#34d399"
        elif "Fila" in str(self.modo_planilha):
            icone_plan = "⏳"
            txt_plan = "Planilha em uso por outro usuário no Excel da rede.\nSalvo na Fila de Sincronização: o N.E.X.U.S gravará automaticamente assim que for fechada!"
            cor_plan = "#f59e0b"
        else:
            icone_plan = "⚠️"
            txt_plan = "Planilha em uso na rede.\nSalvo na contingência local para sincronização posterior."
            cor_plan = "#f59e0b"

        ctk.CTkLabel(
            f_plan,
            text=f"{icone_plan} {txt_plan}",
            font=("Segoe UI", 10),
            text_color=cor_plan,
            justify="left"
        ).pack(anchor="w", padx=10, pady=6)

        # ================= CARD INTEGRAÇÃO WORKFORCE =================
        card_wf = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=10, border_color="#0284c7", border_width=1)
        card_wf.pack(fill="x", padx=20, pady=(0, 10))

        f_wf_top = ctk.CTkFrame(card_wf, fg_color="transparent")
        f_wf_top.pack(fill="x", padx=14, pady=(8, 2))

        ctk.CTkLabel(
            f_wf_top,
            text="🌐 INTEGRAÇÃO WORKFORCE",
            font=("Segoe UI", 11, "bold"),
            text_color="#38bdf8"
        ).pack(side="left")

        self.chk_salvar_wf = ctk.CTkCheckBox(
            f_wf_top,
            text="Salvar automaticamente",
            font=("Segoe UI", 10),
            height=18
        )
        self.chk_salvar_wf.pack(side="right")
        self.chk_salvar_wf.select()

        ctk.CTkLabel(
            card_wf,
            text=f"Abre o Chrome, localiza a Ficha de '{self.equipe}', preenche Doc '{self.doc_sap}' e os seriais.",
            font=("Segoe UI", 10),
            text_color="#94a3b8"
        ).pack(anchor="w", padx=14, pady=(0, 6))

        self.btn_lancar_wf = ctk.CTkButton(
            card_wf,
            text="🌐 Lançar Seriais e Doc no WorkForce",
            font=("Segoe UI", 12, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=36,
            command=self._disparar_workforce
        )
        self.btn_lancar_wf.pack(fill="x", padx=14, pady=(0, 6))

        self.lbl_status_wf = ctk.CTkLabel(
            card_wf,
            text="Pronto para lançar no WorkForce.",
            font=("Segoe UI", 10, "italic"),
            text_color="#cbd5e1"
        )
        self.lbl_status_wf.pack(anchor="w", padx=14, pady=(0, 8))

        # Botão Fechar
        btn_ok = ctk.CTkButton(
            self,
            text="Concluir e Voltar",
            font=("Segoe UI", 11, "bold"),
            fg_color="#3f3f46",
            hover_color="#52525b",
            height=34,
            command=self.destroy
        )
        btn_ok.pack(fill="x", padx=20, pady=(0, 14))

    def _copiar_doc(self):
        self.clipboard_clear()
        self.clipboard_append(self.doc_sap)
        self.btn_copiar.configure(text="✓ Copiado para a Área de Transferência!", fg_color="#059669")
        self.after(2000, lambda: self.btn_copiar.configure(text="📋 Copiar Documento SAP", fg_color="#10b981"))

    def _disparar_workforce(self):
        """Inicia a automação do WorkForce em segundo plano sem congelar a interface."""
        self.btn_lancar_wf.configure(text="⏳ Conectando ao WorkForce...", state="disabled", fg_color="#475569")
        self.lbl_status_wf.configure(text="Iniciando navegador Chrome...", text_color="#38bdf8")
        threading.Thread(target=self._fluxo_workforce_thread, daemon=True).start()

    def _atualizar_status_wf(self, msg):
        self.after(0, lambda: self.lbl_status_wf.configure(text=msg))

    def _fluxo_workforce_thread(self):
        try:
            from utils_workforce import preencher_ficha_saida_workforce
            sucesso, ret = preencher_ficha_saida_workforce(
                equipe=self.equipe,
                doc_sap=self.doc_sap,
                materiais_agrupados=self.totais_materiais,
                callback_status=self._atualizar_status_wf,
                salvar_auto=bool(self.chk_salvar_wf.get())
            )
            if sucesso:
                self.after(0, lambda: self.btn_lancar_wf.configure(
                    text="✓ Lançado no WorkForce com Sucesso!",
                    fg_color="#10b981",
                    hover_color="#059669",
                    state="normal"
                ))
            else:
                self.after(0, lambda: self.btn_lancar_wf.configure(
                    text="🔄 Tentar Novamente",
                    fg_color="#ef4444",
                    hover_color="#dc2626",
                    state="normal"
                ))
        except Exception as e:
            self.after(0, lambda err=str(e): [
                self.lbl_status_wf.configure(text=f"⚠️ Erro: {err}", text_color="#ef4444"),
                self.btn_lancar_wf.configure(text="🔄 Tentar Novamente", fg_color="#ef4444", state="normal")
            ])


class AbaSaidaOSME:
    """Módulo de Saída para OSME com validação de planilha e controle de estoque/fornecidos."""

    ALIASES_CADASTRO_EQUIPE = ("EQUIPES", "EQUIPE", "RECURSO", "RECURSOS")
    ALIASES_CADASTRO_ORDEM = ("ORDEM", "OSME", "NUMERO_OSME", "NUMORDEM")
    ALIASES_CADASTRO_DEP = ("DEPOSITO", "DEPÓSITO", "DEP")
    ALIASES_CADASTRO_SEG = ("SEGMENTO", "SEG", "SETOR")
    ALIASES_CADASTRO_VENC = ("VENCIMENTO", "VENCTO", "DATA_VENC")

    ALIASES_CADASTRO_MAT = ("MATERIAIS", "MATERIAL", "CODIGO", "CODIGOMATERIAL", "MATNR")
    ALIASES_CADASTRO_DESC = ("DESCRICAO", "DESCRIÇÃO", "DENOMINACAO", "DENOMINAÇÃO", "NOME")
    ALIASES_CADASTRO_UND = ("UND", "UNIDADE", "UN", "UM", "MEINS")

    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or DEFAULT_PALETAS
        self.obter_tema_atual = obter_tema_atual or (lambda: "Nexus")

        self.caminho_planilha = None
        self.dados_cadastro = {}           # { EQUIPE_NORM: {"equipe_orig": str, "ordem": str, "deposito": str, "segmento": str, "vencimento": str} }
        self.dados_fornecidos = {}         # { SERIAL_NORM: {"serial": str, "material": str, "equipe": str, "data": str, "doc_sap": str} }
        self.dados_estoque = {}            # { SERIAL_NORM: {"serial": str, "material": str, "deposito": str} }
        self.dados_materiais_cadastro = {} # { MAT_NORM: {"codigo": str, "descricao": str, "unidade": str} }
        self.itens_nao_serializados = []   # [ {"material": str, "descricao": str, "unidade": str, "quantidade": float, "deposito": str} ]
        self.seriais_validados_atuais = [] # [ {"serial": str, ...} ]
        self.carregando = False
        self.processando_sap = False
        self.modal_fornecidos_ativo = None

        self.texto_workforce = ""
        self.totais_materiais = OrderedDict()

        self._carregar_caminho_salvo()
        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())

        # Se já tiver planilha salva, inicia carregamento automático
        if self.caminho_planilha and os.path.isfile(self.caminho_planilha):
            self.carregar_planilha_async()

        # Inicia o verificador automático de pendências de sincronização
        self._iniciar_timer_pendencias()

    def _cores(self):
        tema = self.obter_tema_atual()
        return self.paletas.get(tema, self.paletas.get("Nexus", DEFAULT_PALETAS["Nexus"]))

    def _status(self, msg):
        if self.label_status_global:
            try:
                self.label_status_global.configure(text=f" {msg}")
            except Exception:
                pass

    def _status_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._status(msg))

    def _log(self, msg):
        try:
            self.txt_log.configure(state="normal")
            self.txt_log.insert(tk.END, msg)
            if not msg.endswith("\n"):
                self.txt_log.insert(tk.END, "\n")
            self.txt_log.see(tk.END)
            self.txt_log.configure(state="disabled")
        except Exception:
            pass

    def _log_threadsafe(self, msg):
        self.parent_frame.after(0, lambda: self._log(msg))

    def _carregar_caminho_salvo(self):
        if os.path.exists(CAMINHO_CONFIG):
            try:
                with open(CAMINHO_CONFIG, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                caminho = dados.get("caminho") or dados.get("caminho_planilha")
                if caminho and os.path.isfile(caminho):
                    self.caminho_planilha = caminho
            except Exception:
                pass

    def _salvar_caminho(self):
        try:
            with open(CAMINHO_CONFIG, "w", encoding="utf-8") as f:
                json.dump({
                    "caminho": self.caminho_planilha,
                    "caminho_planilha": self.caminho_planilha
                }, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def _definir_estado_botoes_processar(self, executando=False):
        for btn in (getattr(self, "btn_processar_sap", None), getattr(self, "btn_processar_topo", None)):
            if btn and hasattr(btn, "configure"):
                if executando:
                    btn.configure(state="disabled", text="⏳ Processando SAP...", fg_color="#64748b")
                else:
                    btn.configure(state="normal", text="🚀 Processar Saída no SAP", fg_color="#10b981")

    def _construir_interface(self):
        # 1. BARRA DE AÇÕES FIXA NO RODAPÉ (SEMPRE VISÍVEL - Não oculta na rolagem)
        self.f_barra_fixa = ctk.CTkFrame(self.parent_frame, corner_radius=10, fg_color="#172a45")
        self.f_barra_fixa.pack(side="bottom", fill="x", padx=10, pady=(4, 8))

        f_acoes = ctk.CTkFrame(self.f_barra_fixa, fg_color="transparent")
        f_acoes.pack(fill="x", padx=12, pady=8)

        self.btn_copiar_workforce = ctk.CTkButton(
            f_acoes,
            text="📋 Copiar p/ WorkForce (;)",
            font=("Segoe UI", 11, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=38,
            command=self._copiar_workforce
        )
        self.btn_copiar_workforce.pack(side="left", padx=(0, 8))

        self.btn_ver_workforce = ctk.CTkButton(
            f_acoes,
            text="👁️ Ver Lista WorkForce",
            font=("Segoe UI", 11, "bold"),
            fg_color="transparent",
            border_width=1,
            height=38,
            command=self._abrir_modal_workforce
        )
        self.btn_ver_workforce.pack(side="left", padx=(0, 8))

        self.btn_copiar_tabela = ctk.CTkButton(
            f_acoes,
            text="📋 Copiar Tabela (Excel)",
            font=("Segoe UI", 11, "bold"),
            fg_color="transparent",
            border_width=1,
            height=38,
            command=self._copiar_tabela
        )
        self.btn_copiar_tabela.pack(side="left")

        self.btn_processar_sap = ctk.CTkButton(
            f_acoes,
            text="🚀 Processar Saída no SAP",
            font=("Segoe UI", 12, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=38,
            width=220,
            command=self._processar_saida_sap
        )
        self.btn_processar_sap.pack(side="right")
        self.frame_inferior = self.f_barra_fixa

        # 2. ÁREA ROLÁVEL PARA TODO O CONTEÚDO (Evita encolhimento e permite visualização confortável)
        self.frame_conteudo = ctk.CTkScrollableFrame(self.parent_frame, fg_color="transparent")
        self.frame_conteudo.pack(side="top", fill="both", expand=True, padx=10, pady=(6, 0))

        # 1. CABEÇALHO SUPERIOR
        self.frame_topo = ctk.CTkFrame(self.frame_conteudo, corner_radius=8)
        self.frame_topo.pack(fill="x", pady=(0, 8))

        lbl_tit = ctk.CTkLabel(
            self.frame_topo,
            text="📤 SAÍDA PARA OSME — CONTROLE DE MATERIAIS & ESTOQUE",
            font=("Segoe UI", 15, "bold")
        )
        lbl_tit.pack(anchor="w", padx=15, pady=(8, 2))

        lbl_sub = ctk.CTkLabel(
            self.frame_topo,
            text="Validação inteligente de equipes, OSME atual, conferência de estoque e agrupamento por material para o WorkForce.",
            font=("Segoe UI", 11),
            text_color="#94a3b8"
        )
        lbl_sub.pack(anchor="w", padx=15, pady=(0, 8))

        # 2. SELEÇÃO DA PLANILHA DE CONTROLE
        self.frame_planilha = ctk.CTkFrame(self.frame_conteudo, corner_radius=8)
        self.frame_planilha.pack(fill="x", pady=(0, 8))

        f_plan_botoes = ctk.CTkFrame(self.frame_planilha, fg_color="transparent")
        f_plan_botoes.pack(fill="x", padx=12, pady=(8, 4))

        self.btn_selecionar = ctk.CTkButton(
            f_plan_botoes,
            text="📂 Selecionar Planilha de Controle...",
            font=("Segoe UI", 11, "bold"),
            command=self._selecionar_planilha,
            height=32,
            width=220
        )
        self.btn_selecionar.pack(side="left", padx=(0, 8))

        self.btn_recarregar = ctk.CTkButton(
            f_plan_botoes,
            text="🔄 Recarregar Dados",
            font=("Segoe UI", 11, "bold"),
            command=self.carregar_planilha_async,
            height=32,
            width=140
        )
        self.btn_recarregar.pack(side="left", padx=(0, 12))

        self.lbl_nome_planilha = ctk.CTkLabel(
            f_plan_botoes,
            text=f"Planilha: {os.path.basename(self.caminho_planilha) if self.caminho_planilha else 'Nenhuma planilha selecionada'}",
            font=("Segoe UI", 11, "italic"),
            text_color="#38bdf8" if self.caminho_planilha else "#f87171"
        )
        self.lbl_nome_planilha.pack(side="left")

        self.btn_comparar_backup = ctk.CTkButton(
            f_plan_botoes,
            text="🔍 Comparar vs Mestre",
            font=("Segoe UI", 10, "bold"),
            fg_color="#0284c7",
            hover_color="#0369a1",
            height=30,
            width=150,
            command=self._comparar_backup_com_mestre_acao
        )
        self.btn_comparar_backup.pack(side="right", padx=(6, 0))

        self.btn_abrir_backup = ctk.CTkButton(
            f_plan_botoes,
            text="📁 Backup Diário (.xlsx)",
            font=("Segoe UI", 10, "bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=30,
            width=160,
            command=self._abrir_backup_diario_acao
        )
        self.btn_abrir_backup.pack(side="right", padx=(6, 0))

        # Badges informativos de leitura
        f_badges = ctk.CTkFrame(self.frame_planilha, fg_color="transparent")
        f_badges.pack(fill="x", padx=12, pady=(0, 8))

        self.badge_equipes = ctk.CTkLabel(f_badges, text="📋 Equipes: 0", font=("Segoe UI", 10, "bold"), text_color="#a1a1aa")
        self.badge_equipes.pack(side="left", padx=(0, 15))

        self.badge_estoque = ctk.CTkLabel(f_badges, text="📦 Itens Estoque: 0", font=("Segoe UI", 10, "bold"), text_color="#a1a1aa")
        self.badge_estoque.pack(side="left", padx=(0, 15))

        self.badge_fornecidos = ctk.CTkLabel(f_badges, text="🏷️ Já Fornecidos: 0", font=("Segoe UI", 10, "bold"), text_color="#a1a1aa")
        self.badge_fornecidos.pack(side="left", padx=(0, 15))

        self.badge_backup = ctk.CTkLabel(f_badges, text="🛡️ Backup Diário: Ativo", font=("Segoe UI", 10, "bold"), text_color="#10b981")
        self.badge_backup.pack(side="right")

        # Banner de Pendências de Sincronização (Auto-Retry quando a planilha estiver em uso)
        self.f_banner_pendencias = ctk.CTkFrame(self.frame_planilha, fg_color="#78350f", corner_radius=6)
        
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

        # 3. ÁREA DE ENTRADA: EQUIPE (ESQUERDA) & SERIAIS (DIREITA)
        self.frame_entradas = ctk.CTkFrame(self.frame_conteudo, corner_radius=8)
        self.frame_entradas.pack(fill="x", pady=(0, 8))
        self.frame_entradas.grid_columnconfigure(0, weight=1)
        self.frame_entradas.grid_columnconfigure(1, weight=1)

        # Coluna Esquerda: Equipe & Dados da OSME
        card_equipe = ctk.CTkFrame(self.frame_entradas, fg_color="transparent")
        card_equipe.grid(row=0, column=0, padx=12, pady=10, sticky="nsew")

        ctk.CTkLabel(card_equipe, text="1. EQUIPE / RECURSO:", font=("Segoe UI", 12, "bold")).pack(anchor="w")

        f_eq_input = ctk.CTkFrame(card_equipe, fg_color="transparent")
        f_eq_input.pack(fill="x", pady=(4, 8))

        self.entry_equipe = ctk.CTkEntry(
            f_eq_input,
            placeholder_text="Ex: CDC407, CSC300, CCL500...",
            font=("Segoe UI", 12, "bold"),
            height=34
        )
        self.entry_equipe.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry_equipe.bind("<KeyRelease>", self._ao_digitar_equipe)

        btn_checar_eq = ctk.CTkButton(
            f_eq_input,
            text="Buscar OSME",
            width=100,
            height=34,
            command=self._consultar_equipe_atual
        )
        btn_checar_eq.pack(side="left")

        # Data de Lançamento (Opcional - Fechamento de Mês no SAP)
        f_dt_lcto = ctk.CTkFrame(card_equipe, fg_color="transparent")
        f_dt_lcto.pack(fill="x", pady=(0, 6))

        lbl_dt_lcto = ctk.CTkLabel(
            f_dt_lcto,
            text="📅 Data Lçto. (Fechamento SAP):",
            font=("Segoe UI", 11, "bold"),
            text_color="#f59e0b"
        )
        lbl_dt_lcto.pack(side="left", padx=(0, 8))

        self.entry_data_lancamento = ctk.CTkEntry(
            f_dt_lcto,
            placeholder_text="dd.mm.aaaa (opcional)",
            font=("Consolas", 11),
            width=150,
            height=28
        )
        self.entry_data_lancamento.pack(side="left")

        lbl_dt_help = ctk.CTkLabel(
            f_dt_lcto,
            text="* Deixe vazio para usar a data de hoje",
            font=("Segoe UI", 9, "italic"),
            text_color="#94a3b8"
        )
        lbl_dt_help.pack(side="left", padx=(8, 0))

        # Card de Exibição da OSME Encontrada
        self.card_osme_info = ctk.CTkFrame(card_equipe, corner_radius=8, fg_color="#18273d")
        self.card_osme_info.pack(fill="x", pady=4)

        f_osme_header = ctk.CTkFrame(self.card_osme_info, fg_color="transparent")
        f_osme_header.pack(fill="x", padx=12, pady=(8, 2))

        ctk.CTkLabel(f_osme_header, text="ORDEM / OSME ATUAL:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8").pack(side="left")
        self.lbl_osme_status = ctk.CTkLabel(f_osme_header, text="Aguardando Equipe", font=("Segoe UI", 10, "bold"), text_color="#cbd5e1")
        self.lbl_osme_status.pack(side="right")

        self.lbl_osme_numero = ctk.CTkLabel(
            self.card_osme_info,
            text="—",
            font=("Consolas", 20, "bold"),
            text_color="#f57c00"
        )
        self.lbl_osme_numero.pack(anchor="w", padx=12, pady=(0, 6))

        f_osme_sub = ctk.CTkFrame(self.card_osme_info, fg_color="transparent")
        f_osme_sub.pack(fill="x", padx=12, pady=(0, 8))

        self.lbl_deposito = ctk.CTkLabel(f_osme_sub, text="Depósito: —", font=("Segoe UI", 11), text_color="#e2e8f0")
        self.lbl_deposito.pack(side="left", padx=(0, 15))

        self.lbl_segmento = ctk.CTkLabel(f_osme_sub, text="Segmento: —", font=("Segoe UI", 11), text_color="#e2e8f0")
        self.lbl_segmento.pack(side="left", padx=(0, 15))

        self.lbl_vencimento = ctk.CTkLabel(f_osme_sub, text="Vencimento: —", font=("Segoe UI", 11), text_color="#e2e8f0")
        self.lbl_vencimento.pack(side="left")

        # Sub-Card: Materiais Não Serializados
        self.card_nao_serializados = ctk.CTkFrame(card_equipe, corner_radius=8, fg_color="#18273d")
        self.card_nao_serializados.pack(fill="x", pady=(6, 0))

        f_ns_header = ctk.CTkFrame(self.card_nao_serializados, fg_color="transparent")
        f_ns_header.pack(fill="x", padx=10, pady=(6, 2))
        ctk.CTkLabel(
            f_ns_header,
            text="3. MATERIAIS NÃO SERIALIZADOS (OPCIONAL):",
            font=("Segoe UI", 10, "bold"),
            text_color="#38bdf8"
        ).pack(side="left")

        f_ns_inputs = ctk.CTkFrame(self.card_nao_serializados, fg_color="transparent")
        f_ns_inputs.pack(fill="x", padx=10, pady=(2, 4))

        self.entry_mat_ns = ctk.CTkEntry(
            f_ns_inputs,
            placeholder_text="Cód (ex: 330702)",
            font=("Segoe UI", 11),
            width=130,
            height=30
        )
        self.entry_mat_ns.pack(side="left", padx=(0, 6))
        self.entry_mat_ns.bind("<KeyRelease>", self._ao_digitar_material_ns)

        self.entry_qtd_ns = ctk.CTkEntry(
            f_ns_inputs,
            placeholder_text="Qtd",
            font=("Segoe UI", 11),
            width=65,
            height=30
        )
        self.entry_qtd_ns.pack(side="left", padx=(0, 6))

        self.btn_add_ns = ctk.CTkButton(
            f_ns_inputs,
            text="➕ Adicionar",
            font=("Segoe UI", 11, "bold"),
            width=90,
            height=30,
            fg_color="#0284c7",
            hover_color="#0369a1",
            command=self._adicionar_material_nao_serializado
        )
        self.btn_add_ns.pack(side="left")

        # Label da Descrição e Unidade instantânea
        self.lbl_desc_mat_ns = ctk.CTkLabel(
            self.card_nao_serializados,
            text="Digite o código para autocompletar descrição e unidade da aba CADASTRO...",
            font=("Segoe UI", 10, "italic"),
            text_color="#94a3b8",
            anchor="w"
        )
        self.lbl_desc_mat_ns.pack(fill="x", padx=10, pady=(0, 4))

        # Frame com badges dos materiais não serializados já incluídos
        self.frame_chips_ns = ctk.CTkScrollableFrame(self.card_nao_serializados, height=36, orientation="horizontal", fg_color="transparent")
        self.frame_chips_ns.pack(fill="x", padx=6, pady=(0, 6))

        # Coluna Direita: Seriais para Saída
        card_seriais = ctk.CTkFrame(self.frame_entradas, fg_color="transparent")
        card_seriais.grid(row=0, column=1, padx=12, pady=10, sticky="nsew")

        ctk.CTkLabel(card_seriais, text="2. SERIAIS (COLE A LISTA OU INTERVALOS):", font=("Segoe UI", 12, "bold")).pack(anchor="w")

        self.txt_seriais = ctk.CTkTextbox(
            card_seriais,
            height=120,
            font=("Consolas", 11),
            corner_radius=6
        )
        self.txt_seriais.pack(fill="x", pady=(4, 6))

        # Eventos para autolimpeza de códigos de material/modelo ao colar
        self.txt_seriais.bind("<<Paste>>", self._ao_colar_seriais)
        self.txt_seriais.bind("<Control-v>", self._ao_colar_seriais)
        self.txt_seriais.bind("<Control-V>", self._ao_colar_seriais)
        if hasattr(self.txt_seriais, "_textbox"):
            self.txt_seriais._textbox.bind("<<Paste>>", self._ao_colar_seriais)
            self.txt_seriais._textbox.bind("<Control-v>", self._ao_colar_seriais)
            self.txt_seriais._textbox.bind("<Control-V>", self._ao_colar_seriais)

        f_seriais_botoes = ctk.CTkFrame(card_seriais, fg_color="transparent")
        f_seriais_botoes.pack(fill="x")

        self.btn_validar = ctk.CTkButton(
            f_seriais_botoes,
            text="🔍 Validar Seriais & Conferir Estoque",
            font=("Segoe UI", 11, "bold"),
            height=32,
            command=self.validar_seriais_acao
        )
        self.btn_validar.pack(side="left", padx=(0, 8))

        self.btn_limpar = ctk.CTkButton(
            f_seriais_botoes,
            text="Limpar",
            width=80,
            height=32,
            fg_color="transparent",
            border_width=1,
            command=self._limpar_campos
        )
        self.btn_limpar.pack(side="left")

        # 4. TABELA DE PRÉVIA, TOTALIDADE POR MATERIAL & VALIDAÇÃO DA SAÍDA
        self.frame_tabela = ctk.CTkFrame(self.frame_conteudo, corner_radius=8)
        self.frame_tabela.pack(fill="both", expand=True, pady=(0, 8))

        f_tab_topo = ctk.CTkFrame(self.frame_tabela, fg_color="transparent")
        f_tab_topo.pack(fill="x", padx=12, pady=(10, 4))

        ctk.CTkLabel(f_tab_topo, text="PRÉVIA DA SAÍDA / RESULTADOS DA CONFERÊNCIA:", font=("Segoe UI", 12, "bold")).pack(side="left")

        self.btn_processar_topo = ctk.CTkButton(
            f_tab_topo,
            text="🚀 Processar Saída no SAP",
            font=("Segoe UI", 11, "bold"),
            fg_color="#10b981",
            hover_color="#059669",
            height=30,
            width=190,
            command=self._processar_saida_sap
        )
        self.btn_processar_topo.pack(side="right", padx=(10, 0))

        self.lbl_contagem_tabela = ctk.CTkLabel(f_tab_topo, text="0 itens adicionados", font=("Segoe UI", 11, "italic"), text_color="#94a3b8")
        self.lbl_contagem_tabela.pack(side="right")

        # Container de Totalidade por Material (Chips com status de cada material)
        self.frame_resumo_materiais = ctk.CTkScrollableFrame(self.frame_tabela, height=44, orientation="horizontal", fg_color="transparent")
        self.frame_resumo_materiais.pack(fill="x", padx=10, pady=(2, 6))

        self.lbl_resumo_vazio = ctk.CTkLabel(
            self.frame_resumo_materiais,
            text="Aguardando validação dos seriais...",
            font=("Segoe UI", 10, "italic"),
            text_color="#a1a1aa"
        )
        self.lbl_resumo_vazio.pack(side="left", padx=5)

        colunas = ("item", "serial", "material", "deposito", "equipe", "osme", "status_validacao")
        titulos = {
            "item": "#",
            "serial": "Número de Série",
            "material": "Material",
            "deposito": "Depósito",
            "equipe": "Equipe Destino",
            "osme": "OSME Alvo",
            "status_validacao": "Status / Validação"
        }
        larguras = {
            "item": 40, "serial": 180, "material": 120, "deposito": 90,
            "equipe": 90, "osme": 130, "status_validacao": 260
        }

        frame_tv = ctk.CTkFrame(self.frame_tabela, fg_color="transparent")
        frame_tv.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.tree = ttk.Treeview(frame_tv, columns=colunas, show="headings", height=12, style="SaidaOSME.Treeview")
        for c in colunas:
            self.tree.heading(c, text=titulos[c])
            self.tree.column(c, width=larguras[c], anchor="center")

        scroll_y = ttk.Scrollbar(frame_tv, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        # Tags de cores para o Treeview: itens bloqueados/fornecidos com fundo e texto vermelho!
        self.tree.tag_configure("ok", foreground="#34d399")
        self.tree.tag_configure("bloqueado", foreground="#fca5a5", background="#450a0a")
        self.tree.tag_configure("alerta", foreground="#fcd34d")
        self.tree.tag_configure("padrao", foreground="#f8fafc")

        # 5. PAINEL DE LOG DE AUDITORIA E EVENTOS
        self.card_log = card_log = ctk.CTkFrame(self.frame_conteudo, corner_radius=8, fg_color="#172a45")
        card_log.pack(fill="x", pady=(0, 10))

        f_tit_log = ctk.CTkFrame(card_log, fg_color="transparent")
        f_tit_log.pack(fill="x", padx=12, pady=(8, 2))

        ctk.CTkLabel(f_tit_log, text="📜 LOG DE AUDITORIA E EVENTOS:", font=("Segoe UI", 10, "bold"), text_color="#94a3b8").pack(side="left")

        self.txt_log = ctk.CTkTextbox(card_log, height=65, font=("Consolas", 10))
        self.txt_log.pack(fill="x", padx=12, pady=(0, 8))
        self.txt_log.configure(state="disabled")

    def _abrir_backup_diario_acao(self):
        try:
            ok, res = abrir_backup_diario()
            if ok:
                self._log(f"📁 Planilha de backup diário aberta no Excel: {os.path.basename(res)}")
                self._status(f"Backup diário aberto: {os.path.basename(res)}")
            else:
                messagebox.showerror("Backup Diário", f"Não foi possível abrir o arquivo de backup:\n{res}")
        except Exception as e:
            messagebox.showerror("Backup Diário", f"Erro ao abrir backup:\n{e}")

    def _comparar_backup_com_mestre_acao(self):
        if not self.caminho_planilha or not os.path.exists(self.caminho_planilha):
            messagebox.showwarning("Auditoria de Saídas", "Selecione primeiro a planilha mestre de controle.")
            return

        self._status("Comparando backup diário com planilha mestre...")
        self._log("⏳ Realizando auditoria: cruzando backup diário com a aba FORNECIDOS da rede...")

        def _thread_auditoria():
            res = comparar_backup_com_mestre(self.caminho_planilha)
            self.parent_frame.after(0, lambda: self._exibir_modal_auditoria(res))

        threading.Thread(target=_thread_auditoria, daemon=True).start()

    def _exibir_modal_auditoria(self, res):
        self._status("Auditoria concluída.")
        ModalCompararBackup(
            self.parent_frame,
            resultado=res,
            caminho_mestre=self.caminho_planilha,
            callback_recarregar=self.carregar_planilha_async,
            paletas=self.paletas,
            obter_tema_atual=self.obter_tema_atual
        )

    def _selecionar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Planilha de Controle (CONTROLE_SERIAIS)",
            filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls"), ("Todos os arquivos", "*.*")]
        )
        if caminho and os.path.isfile(caminho):
            self.caminho_planilha = caminho
            self._salvar_caminho()
            self.lbl_nome_planilha.configure(
                text=f"Planilha: {os.path.basename(caminho)}",
                text_color="#38bdf8"
            )
            self._log(f"📁 Planilha selecionada: {caminho}")
            self.carregar_planilha_async()

    def carregar_planilha_async(self):
        if not self.caminho_planilha or not os.path.isfile(self.caminho_planilha):
            self._carregar_caminho_salvo()

        if not self.caminho_planilha or not os.path.isfile(self.caminho_planilha):
            messagebox.showwarning("Saída OSME", "Por favor, selecione primeiro a planilha de controle.")
            return

        if self.carregando:
            return

        self.carregando = True
        self.btn_recarregar.configure(state="disabled", text="Lendo...")
        self._status_threadsafe("Saída OSME: lendo abas CADASTRO, ESTOQUE e FORNECIDOS...")
        self._log_threadsafe(f"⏳ Carregando dados da planilha em segundo plano ({self.caminho_planilha})...")

        threading.Thread(target=self._ler_planilha_thread, daemon=True).start()

    def _ler_planilha_thread(self):
        try:
            import io
            with open(self.caminho_planilha, "rb") as f_raw:
                stream_planilha = io.BytesIO(f_raw.read())

            wb = load_workbook(
                stream_planilha,
                read_only=True,
                data_only=True,
                keep_vba=self.caminho_planilha.lower().endswith(".xlsm")
            )
            try:
                # 1. Leitura da Aba CADASTRO (Equipes/OSME e Materiais Cadastrados)
                cadastro_lido = {}
                materiais_lidos = {}
                if "CADASTRO" in wb.sheetnames:
                    ws_cad = wb["CADASTRO"]
                    cabecalho = None
                    idx_equipe, idx_ordem, idx_dep, idx_seg, idx_venc = None, None, None, None, None
                    idx_mat, idx_desc, idx_und = None, None, None

                    for r_num, row in enumerate(ws_cad.iter_rows(max_row=10, values_only=True), 1):
                        for c_idx, val in enumerate(row):
                            norm = normalizar_texto(val)
                            if norm in self.ALIASES_CADASTRO_EQUIPE and idx_equipe is None:
                                idx_equipe = c_idx
                            elif norm in self.ALIASES_CADASTRO_ORDEM and idx_ordem is None:
                                idx_ordem = c_idx
                            elif norm in self.ALIASES_CADASTRO_DEP and idx_dep is None:
                                idx_dep = c_idx
                            elif norm in self.ALIASES_CADASTRO_SEG and idx_seg is None:
                                idx_seg = c_idx
                            elif norm in self.ALIASES_CADASTRO_VENC and idx_venc is None:
                                idx_venc = c_idx
                            elif norm in self.ALIASES_CADASTRO_MAT and idx_mat is None:
                                idx_mat = c_idx
                            elif norm in self.ALIASES_CADASTRO_DESC and idx_desc is None:
                                idx_desc = c_idx
                            elif norm in self.ALIASES_CADASTRO_UND and idx_und is None:
                                idx_und = c_idx

                        if idx_equipe is not None and idx_ordem is not None:
                            cabecalho = r_num
                            break

                    if idx_equipe is None: idx_equipe = 4
                    if idx_dep is None: idx_dep = 5
                    if idx_seg is None: idx_seg = 6
                    if idx_ordem is None: idx_ordem = 7
                    if idx_venc is None: idx_venc = 8
                    if idx_mat is None: idx_mat = 10
                    if idx_desc is None: idx_desc = 11
                    if idx_und is None: idx_und = 12
                    if cabecalho is None: cabecalho = 1

                    for row in ws_cad.iter_rows(min_row=cabecalho + 1, values_only=True):
                        if not row:
                            continue
                        # Leitura de Equipes
                        if len(row) > idx_equipe and row[idx_equipe]:
                            eq_val = row[idx_equipe]
                            eq_norm = normalizar_texto(eq_val)
                            ordem_val = row[idx_ordem] if len(row) > idx_ordem else ""
                            dep_val = row[idx_dep] if len(row) > idx_dep else ""
                            seg_val = row[idx_seg] if len(row) > idx_seg else ""
                            venc_val = row[idx_venc] if len(row) > idx_venc else ""

                            if eq_norm:
                                cadastro_lido[eq_norm] = {
                                    "equipe_orig": str(eq_val).strip(),
                                    "ordem": str(ordem_val or "").strip(),
                                    "deposito": str(dep_val or "").strip(),
                                    "segmento": str(seg_val or "").strip(),
                                    "vencimento": formatar_data_excel(venc_val)
                                }

                        # Leitura de Materiais Cadastrados (colunas K, L, M)
                        if len(row) > idx_mat and row[idx_mat]:
                            mat_raw = row[idx_mat]
                            m_norm = normalizar_texto(mat_raw)
                            if m_norm:
                                d_val = str(row[idx_desc] or "").strip() if len(row) > idx_desc else ""
                                u_val = str(row[idx_und] or "").strip() if len(row) > idx_und else "UN"
                                materiais_lidos[m_norm] = {
                                    "codigo": str(mat_raw).strip(),
                                    "descricao": d_val,
                                    "unidade": u_val
                                }

                # 2. Leitura da Aba FORNECIDOS (com detecção inteligente de cabeçalhos)
                fornecidos_lidos = {}
                if "FORNECIDOS" in wb.sheetnames:
                    ws_forn = wb["FORNECIDOS"]
                    idx_f_ser, idx_f_mat, idx_f_dep, idx_f_eq, idx_f_dt, idx_f_doc = None, None, None, None, None, None
                    for c_idx, val in enumerate(next(ws_forn.iter_rows(max_row=1, values_only=True), ())):
                        norm = normalizar_texto(val)
                        if norm in ("SERIE", "SERIAL", "SERNR", "NUMEROSERIE") and idx_f_ser is None:
                            idx_f_ser = c_idx
                        elif norm in ("MATERIAL", "MAT", "MATNR", "CODIGO") and idx_f_mat is None:
                            idx_f_mat = c_idx
                        elif norm in ("DEPOSITO", "DEP", "LGORT") and idx_f_dep is None:
                            idx_f_dep = c_idx
                        elif norm in ("EQUIPE", "RECURSO") and idx_f_eq is None:
                            idx_f_eq = c_idx
                        elif norm in ("DATA", "DATADOC", "DATE") and idx_f_dt is None:
                            idx_f_dt = c_idx
                        elif norm in ("DOCSAP", "DOC", "DOCUMENTO", "NUMDOC") and idx_f_doc is None:
                            idx_f_doc = c_idx

                    if idx_f_ser is None: idx_f_ser = 0
                    if idx_f_mat is None: idx_f_mat = 1
                    if idx_f_dep is None: idx_f_dep = 2
                    if idx_f_eq is None: idx_f_eq = 3
                    if idx_f_dt is None: idx_f_dt = 4
                    if idx_f_doc is None: idx_f_doc = 5

                    for row in ws_forn.iter_rows(min_row=2, values_only=True):
                        if not row or len(row) <= idx_f_ser or not row[idx_f_ser]:
                            continue
                        serial_val = str(row[idx_f_ser]).strip()
                        s_norm = normalizar_texto(serial_val)
                        if not s_norm:
                            continue
                        mat_val = row[idx_f_mat] if len(row) > idx_f_mat else ""
                        dep_val = row[idx_f_dep] if len(row) > idx_f_dep else ""
                        eq_val = row[idx_f_eq] if len(row) > idx_f_eq else ""
                        data_val = row[idx_f_dt] if len(row) > idx_f_dt else ""
                        doc_val = row[idx_f_doc] if len(row) > idx_f_doc else ""

                        fornecidos_lidos[s_norm] = {
                            "serial": serial_val,
                            "material": str(mat_val or "").strip(),
                            "deposito": str(dep_val or "").strip(),
                            "equipe": str(eq_val or "").strip(),
                            "data": formatar_data_excel(data_val),
                            "doc_sap": str(doc_val or "").strip()
                        }

                # 3. Leitura da Aba ESTOQUE (com detecção inteligente de cabeçalhos)
                estoque_lido = {}
                if "ESTOQUE" in wb.sheetnames:
                    ws_est = wb["ESTOQUE"]
                    idx_est_ser, idx_est_mat, idx_est_dep = None, None, None
                    for c_idx, val in enumerate(next(ws_est.iter_rows(max_row=1, values_only=True), ())):
                        norm = normalizar_texto(val)
                        if norm in ("SERIE", "SERIAL", "SERNR", "NUMEROSERIE") and idx_est_ser is None:
                            idx_est_ser = c_idx
                        elif norm in ("MATERIAL", "MAT", "MATNR", "CODIGO") and idx_est_mat is None:
                            idx_est_mat = c_idx
                        elif norm in ("DEPOSITO", "DEP", "LGORT") and idx_est_dep is None:
                            idx_est_dep = c_idx

                    if idx_est_ser is None: idx_est_ser = 0
                    if idx_est_mat is None: idx_est_mat = 1
                    if idx_est_dep is None: idx_est_dep = 2

                    for row in ws_est.iter_rows(min_row=2, values_only=True):
                        if not row or len(row) <= idx_est_ser or not row[idx_est_ser]:
                            continue
                        serial_val = str(row[idx_est_ser]).strip()
                        s_norm = normalizar_texto(serial_val)
                        if not s_norm:
                            continue
                        mat_val = row[idx_est_mat] if len(row) > idx_est_mat else ""
                        dep_val = row[idx_est_dep] if len(row) > idx_est_dep else ""

                        estoque_lido[s_norm] = {
                            "serial": serial_val,
                            "material": str(mat_val or "").strip(),
                            "deposito": str(dep_val or "").strip()
                        }
            finally:
                wb.close()
                stream_planilha.close()

            import time
            for _ in range(50):
                try:
                    self.parent_frame.after(
                        0,
                        lambda: self._concluir_leitura(cadastro_lido, fornecidos_lidos, estoque_lido, materiais_lidos)
                    )
                    break
                except (RuntimeError, tk.TclError):
                    time.sleep(0.1)

        except Exception as e:
            import time
            for _ in range(30):
                try:
                    self.parent_frame.after(0, lambda e=e: self._erro_leitura(str(e)))
                    break
                except (RuntimeError, tk.TclError):
                    time.sleep(0.1)

    def _concluir_leitura(self, cad, forn, est, mats=None):
        self.carregando = False
        self.dados_cadastro = cad
        self.dados_fornecidos = forn
        self.dados_estoque = est
        self.dados_materiais_cadastro = mats or {}

        self.btn_recarregar.configure(state="normal", text="🔄 Recarregar Dados")
        self.badge_equipes.configure(text=f"📋 Equipes: {len(cad)}", text_color="#10b981")
        self.badge_estoque.configure(text=f"📦 Itens Estoque: {len(est)}", text_color="#38bdf8")
        self.badge_fornecidos.configure(text=f"🏷️ Já Fornecidos: {len(forn)}", text_color="#f59e0b")

        self._status(f"Saída OSME: Planilha carregada! {len(cad)} equipes, {len(est)} itens estoque, {len(forn)} fornecidos.")
        self._log(f"✅ Leitura concluída: {len(cad)} equipes cadastradas, {len(est)} em estoque, {len(forn)} seriais fornecidos.")

        # Revalida equipe caso já tenha algo digitado
        self._consultar_equipe_atual()

        # Revalida seriais automaticamente se houver seriais preenchidos no campo de texto
        texto_seriais = self.txt_seriais.get("1.0", tk.END).strip()
        if texto_seriais:
            self._status("Saída OSME: Planilha atualizada! Revalidando seriais em tela...")
            self.validar_seriais_acao()

    def _erro_leitura(self, erro_msg):
        self.carregando = False
        self.btn_recarregar.configure(state="normal", text="🔄 Recarregar Dados")
        self._status("Saída OSME: Erro ao carregar planilha.")
        self._log(f"❌ Erro na leitura da planilha: {erro_msg}")
        messagebox.showerror("Saída OSME", f"Falha ao ler planilha:\n{erro_msg}")

    def _ao_digitar_equipe(self, event=None):
        texto = self.entry_equipe.get().strip()
        if len(texto) >= 3:
            self._consultar_equipe_atual()

    def _consultar_equipe_atual(self):
        eq_digitada = self.entry_equipe.get().strip()
        if not eq_digitada:
            self.lbl_osme_numero.configure(text="—", text_color="#f57c00")
            self.lbl_osme_status.configure(text="Aguardando Equipe", text_color="#cbd5e1")
            self.lbl_deposito.configure(text="Depósito: —")
            self.lbl_segmento.configure(text="Segmento: —")
            self.lbl_vencimento.configure(text="Vencimento: —")
            return

        eq_norm = normalizar_texto(eq_digitada)
        if eq_norm in self.dados_cadastro:
            dados = self.dados_cadastro[eq_norm]
            ordem = dados.get("ordem") or "Sem OSME"
            dep = dados.get("deposito") or "-"
            seg = dados.get("segmento") or "-"
            venc = dados.get("vencimento") or "-"

            self.lbl_osme_numero.configure(text=ordem, text_color="#34d399")
            self.lbl_osme_status.configure(text="Equipe Localizada", text_color="#10b981")
            self.lbl_deposito.configure(text=f"Depósito: {dep}")
            self.lbl_segmento.configure(text=f"Segmento: {seg}")
            self.lbl_vencimento.configure(text=f"Vencimento: {venc}")
            self._status(f"OSME {ordem} encontrada para equipe {dados.get('equipe_orig')}")
        else:
            self.lbl_osme_numero.configure(text="Não Encontrada", text_color="#f87171")
            self.lbl_osme_status.configure(text="Não Cadastrada", text_color="#ef4444")
            self.lbl_deposito.configure(text="Depósito: —")
            self.lbl_segmento.configure(text="Segmento: —")
            self.lbl_vencimento.configure(text="Vencimento: —")

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

    def _ao_digitar_material_ns(self, event=None):
        cod_raw = self.entry_mat_ns.get().strip()
        cod_norm = normalizar_texto(cod_raw)
        if not cod_norm:
            self.lbl_desc_mat_ns.configure(
                text="Digite o código para autocompletar descrição e unidade da aba CADASTRO...",
                text_color="#94a3b8"
            )
            return

        if cod_norm in self.dados_materiais_cadastro:
            info = self.dados_materiais_cadastro[cod_norm]
            desc = info.get("descricao", "")
            und = info.get("unidade", "UN")
            self.lbl_desc_mat_ns.configure(
                text=f"✓ {desc} | UND: {und}",
                text_color="#34d399"
            )
        else:
            self.lbl_desc_mat_ns.configure(
                text=f"Material {cod_raw} não encontrado no cadastro (usará UND padrão)",
                text_color="#f59e0b"
            )

    def _adicionar_material_nao_serializado(self):
        cod_raw = self.entry_mat_ns.get().strip()
        qtd_raw = self.entry_qtd_ns.get().strip()
        if not cod_raw:
            messagebox.showwarning("Saída OSME", "Informe o código do material não serializado.")
            self.entry_mat_ns.focus()
            return
        if not qtd_raw:
            messagebox.showwarning("Saída OSME", "Informe a quantidade do material.")
            self.entry_qtd_ns.focus()
            return

        # Limpeza inteligente da quantidade (ex: "50m", "50 un", "50,5m" -> 50.5)
        m_qtd = re.search(r"(\d+(?:[.,]\d+)?)", qtd_raw)
        if m_qtd:
            qtd_limpa = m_qtd.group(1).replace(",", ".")
        else:
            qtd_limpa = qtd_raw.replace(",", ".")

        try:
            qtd_num = float(qtd_limpa)
            if qtd_num <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showwarning("Saída OSME", "A quantidade deve ser um número maior que zero.")
            self.entry_qtd_ns.focus()
            return

        cod_norm = normalizar_texto(cod_raw)
        desc = ""
        und = "UN"
        if cod_norm in self.dados_materiais_cadastro:
            info = self.dados_materiais_cadastro[cod_norm]
            desc = info.get("descricao", "")
            und = info.get("unidade", "UN")

        eq_digitada = self.entry_equipe.get().strip().upper()
        eq_norm = normalizar_texto(eq_digitada)
        dep_padrao = "CL03"
        osme_atual = "—"
        if eq_norm in self.dados_cadastro:
            osme_atual = self.dados_cadastro[eq_norm].get("ordem") or "—"
            dep_padrao = self.dados_cadastro[eq_norm].get("deposito") or "CL03"

        qtd_formatada = str(int(qtd_num)) if qtd_num.is_integer() else f"{qtd_num:.3f}".rstrip("0").rstrip(".")

        item_ns = {
            "serial": f"(Não Serializado - {qtd_formatada} {und})",
            "material": str(cod_raw).strip(),
            "descricao": desc,
            "unidade": und,
            "quantidade": qtd_formatada,
            "quantidade_num": qtd_num,
            "deposito": dep_padrao,
            "equipe": eq_digitada or "-",
            "osme": osme_atual,
            "status_val": f"✅ Liberado (Não Serializado - {qtd_formatada} {und})",
            "tag": "ok",
            "eh_bloqueado": False,
            "eh_nao_serializado": True
        }

        self.itens_nao_serializados.append(item_ns)
        self.entry_mat_ns.delete(0, tk.END)
        self.entry_qtd_ns.delete(0, tk.END)
        self.lbl_desc_mat_ns.configure(
            text="Digite o código para autocompletar descrição e unidade da aba CADASTRO...",
            text_color="#94a3b8"
        )
        self._recarregar_chips_ns()
        self._renderizar_saida()
        self._log(f"➕ Material não serializado adicionado: {cod_raw} ({qtd_formatada} {und})")

    def _remover_material_nao_serializado(self, idx):
        if 0 <= idx < len(self.itens_nao_serializados):
            removido = self.itens_nao_serializados.pop(idx)
            self._recarregar_chips_ns()
            self._renderizar_saida()
            self._log(f"Removido material não serializado: {removido.get('material')}")

    def _recarregar_chips_ns(self):
        for w in self.frame_chips_ns.winfo_children():
            w.destroy()

        if not self.itens_nao_serializados:
            lbl = ctk.CTkLabel(
                self.frame_chips_ns,
                text="Nenhum material não serializado incluído.",
                font=("Segoe UI", 9, "italic"),
                text_color="#64748b"
            )
            lbl.pack(side="left", padx=4)
            return

        for idx, it in enumerate(self.itens_nao_serializados):
            chip = ctk.CTkFrame(self.frame_chips_ns, fg_color="#1e293b", corner_radius=5)
            chip.pack(side="left", padx=3, pady=1)

            lbl_txt = f"📦 {it['material']}: {it['quantidade']} {it['unidade']}"
            ctk.CTkLabel(chip, text=lbl_txt, font=("Segoe UI", 10, "bold"), text_color="#38bdf8").pack(side="left", padx=(6, 2), pady=2)

            btn_del = ctk.CTkButton(
                chip,
                text="✕",
                width=18,
                height=18,
                fg_color="transparent",
                hover_color="#7f1d1d",
                text_color="#f87171",
                font=("Segoe UI", 10, "bold"),
                command=lambda i=idx: self._remover_material_nao_serializado(i)
            )
            btn_del.pack(side="left", padx=(0, 4), pady=2)

    def _obter_conjunto_materiais_conhecidos(self) -> set:
        """Retorna o conjunto de todos os códigos de materiais conhecidos (Cadastro, Estoque, Fornecidos e Banco)."""
        mats = set(obter_materiais_banco_local())
        if hasattr(self, "dados_materiais_cadastro") and self.dados_materiais_cadastro:
            for k, info in self.dados_materiais_cadastro.items():
                mats.add(k.upper())
                cod = str(info.get("codigo") or "").strip().upper()
                if cod:
                    mats.add(cod)
                    mats.add(normalizar_texto(cod))
        if hasattr(self, "dados_estoque") and self.dados_estoque:
            for info in self.dados_estoque.values():
                m = str(info.get("material") or "").strip().upper()
                if m and m != "—":
                    mats.add(m)
                    mats.add(normalizar_texto(m))
        if hasattr(self, "dados_fornecidos") and self.dados_fornecidos:
            for info in self.dados_fornecidos.values():
                m = str(info.get("material") or "").strip().upper()
                if m and m != "—":
                    mats.add(m)
                    mats.add(normalizar_texto(m))
        return mats

    def _eh_codigo_ou_modelo_material(self, token, mats_conhecidos=None) -> bool:
        """Verifica se um token extraído é na verdade um código de material ou modelo de medidor."""
        t = str(token or "").strip().upper()
        if not t:
            return False
        if t in PALAVRAS_IGNORAR_SERIAIS:
            return True
        if PADRAO_MATERIAL_SAP.match(t):
            return True
        if PADRAO_MODELO_MEDIDOR.match(t):
            return True
        if mats_conhecidos:
            if t in mats_conhecidos or normalizar_texto(t) in mats_conhecidos:
                return True
        return False

    def _ao_colar_seriais(self, event=None):
        """Ao colar texto no campo de seriais, higieniza automaticamente após inserção."""
        if hasattr(self, "parent_frame") and self.parent_frame:
            self.parent_frame.after(60, lambda: self._sanitizar_texto_seriais_auto(verbose=True))

    def _sanitizar_texto_seriais_auto(self, verbose=True):
        """
        Higieniza o campo txt_seriais, removendo automaticamente códigos de material (ex: 510388, 510850)
        e modelos de medidores (ex: D2315240E10SU) colados junto com os seriais.
        """
        try:
            texto = self.txt_seriais.get("1.0", tk.END).strip()
            if not texto:
                return []

            seriais_brutos = extrair_lista_seriais(texto)
            if not seriais_brutos:
                return []

            mats_conhecidos = self._obter_conjunto_materiais_conhecidos()
            seriais_limpos = []
            materiais_removidos = []

            for s in seriais_brutos:
                if self._eh_codigo_ou_modelo_material(s, mats_conhecidos):
                    materiais_removidos.append(s)
                else:
                    seriais_limpos.append(s)

            if materiais_removidos:
                self.txt_seriais.delete("1.0", tk.END)
                if seriais_limpos:
                    self.txt_seriais.insert("1.0", "\n".join(seriais_limpos))

                unicos_mats = list(dict.fromkeys(materiais_removidos))
                txt_mats = ", ".join(unicos_mats)
                if verbose:
                    self._log(f"🧹 Filtro automático: {len(materiais_removidos)} código(s) de material/modelo removido(s): {txt_mats}")
                    self._status(f"Saída OSME: {len(materiais_removidos)} código(s) de material/modelo ({txt_mats}) removido(s) do campo de seriais.")

            return seriais_limpos
        except Exception:
            return []

    def validar_seriais_acao(self):
        texto = self.txt_seriais.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Saída OSME", "Insira um ou mais números de série para validar.")
            return

        seriais_brutos = extrair_lista_seriais(texto)
        if not seriais_brutos:
            messagebox.showwarning("Saída OSME", "Nenhum serial válido foi identificado no texto inserido.")
            return

        # Filtra automaticamente códigos de material e modelos de medidores colados no campo
        mats_conhecidos = self._obter_conjunto_materiais_conhecidos()
        seriais_extraidos = []
        materiais_removidos = []

        for s in seriais_brutos:
            if self._eh_codigo_ou_modelo_material(s, mats_conhecidos):
                materiais_removidos.append(s)
            else:
                seriais_extraidos.append(s)

        if materiais_removidos:
            self.txt_seriais.delete("1.0", tk.END)
            if seriais_extraidos:
                self.txt_seriais.insert("1.0", "\n".join(seriais_extraidos))

            unicos_mats = list(dict.fromkeys(materiais_removidos))
            txt_mats = ", ".join(unicos_mats)
            self._log(f"🧹 [Validação] {len(materiais_removidos)} código(s) de material/modelo removido(s): {txt_mats}")
            self._status(f"Saída OSME: {len(materiais_removidos)} código(s) de material/modelo ({txt_mats}) removido(s) automaticamente.")

        if not seriais_extraidos:
            if materiais_removidos:
                unicos_mats = ", ".join(dict.fromkeys(materiais_removidos))
                messagebox.showinfo(
                    "Saída OSME - Apenas Materiais Detectados",
                    f"Foram identificados apenas códigos de material/modelo ({unicos_mats}) e nenhum número de série de medidor válido no campo.\n\n"
                    "Para materiais não serializados, utilize o quadro à esquerda ('1. Materiais Não Serializados')."
                )
            else:
                messagebox.showwarning("Saída OSME", "Nenhum serial válido foi identificado no texto inserido.")
            return

        eq_digitada = self.entry_equipe.get().strip().upper()
        eq_norm = normalizar_texto(eq_digitada)
        osme_atual = "—"
        dep_padrao = "CL03"
        if eq_norm in self.dados_cadastro:
            osme_atual = self.dados_cadastro[eq_norm].get("ordem") or "—"
            dep_padrao = self.dados_cadastro[eq_norm].get("deposito") or "CL03"

        itens_processados = []
        itens_bloqueados = []

        for serial in seriais_extraidos:
            s_norm = normalizar_texto(serial)
            material = "—"
            deposito = dep_padrao
            status_val = "✅ Liberado para Saída"
            tag = "ok"
            eh_bloqueado = False

            # Verifica se já consta na aba FORNECIDOS
            if s_norm in self.dados_fornecidos:
                info_forn = self.dados_fornecidos[s_norm]
                itens_bloqueados.append(info_forn)
                eh_bloqueado = True
                material = info_forn.get("material") or material
                deposito = info_forn.get("deposito") or deposito
                doc_sap = info_forn.get("doc_sap") or "-"
                dt = info_forn.get("data") or "-"
                status_val = f"⛔ JÁ FORNECIDO (Doc {doc_sap} em {dt})"
                tag = "bloqueado"

            # Se não estiver fornecido, verifica na aba ESTOQUE
            elif s_norm in self.dados_estoque:
                info_est = self.dados_estoque[s_norm]
                material = info_est.get("material") or material
                deposito = info_est.get("deposito") or deposito
                status_val = "✅ Disponível em Estoque"
                tag = "ok"

            else:
                status_val = "⚠️ Não localizado na aba Estoque (Conferir SAP)"
                tag = "alerta"

            itens_processados.append({
                "serial": serial,
                "material": material,
                "deposito": deposito,
                "equipe": eq_digitada or "-",
                "osme": osme_atual,
                "status_val": status_val,
                "tag": tag,
                "eh_bloqueado": eh_bloqueado,
                "eh_nao_serializado": False
            })

        self.seriais_validados_atuais = itens_processados
        self._renderizar_saida()

        if itens_bloqueados:
            self._log(f"⚠️ Atenção: {len(itens_bloqueados)} seriais já constam como FORNECIDOS!")
            if getattr(self, "modal_fornecidos_ativo", None):
                try:
                    self.modal_fornecidos_ativo.destroy()
                except Exception:
                    pass
            self.modal_fornecidos_ativo = ModalAlertaFornecidos(self.parent_frame, itens_bloqueados, self.paletas, self.obter_tema_atual)
        else:
            if getattr(self, "modal_fornecidos_ativo", None):
                try:
                    self.modal_fornecidos_ativo.destroy()
                except Exception:
                    pass
                self.modal_fornecidos_ativo = None
            self._log(f"✅ Validação concluída: {len(seriais_extraidos)} seriais prontos e agrupados por material!")

    def _renderizar_saida(self):
        """Une os seriais validados e os materiais não serializados, agrupando por material."""
        todos_itens = list(self.seriais_validados_atuais) + list(self.itens_nao_serializados)

        # Agrupamento e ordenação por Material para que todos os itens do mesmo material fiquem juntos
        materiais_agrupados = OrderedDict()
        for it in todos_itens:
            mat = it["material"]
            if mat not in materiais_agrupados:
                materiais_agrupados[mat] = []
            materiais_agrupados[mat].append(it)

        itens_ordenados = []
        for mat, lista_mat in materiais_agrupados.items():
            itens_ordenados.extend(lista_mat)

        self.tree.delete(*self.tree.get_children())
        total_validos = 0
        total_bloqueados = 0

        for idx, it in enumerate(itens_ordenados, 1):
            if it["eh_bloqueado"]:
                total_bloqueados += 1
                serial_display = f"⛔ {it['serial']}"
            else:
                total_validos += 1
                serial_display = it["serial"]

            self.tree.insert(
                "",
                "end",
                values=(idx, serial_display, it["material"], it["deposito"], it["equipe"], it["osme"], it["status_val"]),
                tags=(it["tag"],)
            )

        self.lbl_contagem_tabela.configure(
            text=f"Total: {len(todos_itens)} itens | ✅ Liberados: {total_validos} | ⛔ Bloqueados: {total_bloqueados}"
        )

        # Atualiza os Chips de Totalidade por Material
        self._atualizar_chips_materiais(materiais_agrupados)

        # Gera a lista formatada para o WorkForce (1 por linha terminando com ';', apenas seriais válidos)
        linhas_wf = []
        for mat, lista_mat in materiais_agrupados.items():
            for it in lista_mat:
                if not it.get("eh_nao_serializado"):
                    linhas_wf.append(f"{it['serial']};")
        self.texto_workforce = "\n".join(linhas_wf)
        self.totais_materiais = materiais_agrupados

    def _atualizar_chips_materiais(self, materiais_agrupados):
        """Renderiza os cards com a totalidade e status de cada material encontrado."""
        for widget in self.frame_resumo_materiais.winfo_children():
            widget.destroy()

        if not materiais_agrupados:
            lbl_vazio = ctk.CTkLabel(
                self.frame_resumo_materiais,
                text="Aguardando validação dos seriais...",
                font=("Segoe UI", 10, "italic"),
                text_color="#a1a1aa"
            )
            lbl_vazio.pack(side="left", padx=5)
            return

        for mat, lista_itens in materiais_agrupados.items():
            tem_ns = any(it.get("eh_nao_serializado") for it in lista_itens)
            bloqueados = sum(1 for it in lista_itens if it["eh_bloqueado"])
            todos_ok = (bloqueados == 0) and (mat != "—")

            if tem_ns:
                soma_qtd = sum(float(it.get("quantidade_num", 0)) for it in lista_itens if it.get("eh_nao_serializado"))
                und = lista_itens[0].get("unidade", "UN")
                qtd_str = f"{soma_qtd:g} {und}"
                texto_qtd = f"{qtd_str}"
            else:
                total = len(lista_itens)
                texto_qtd = f"{total} unid{'s' if total > 1 else ''}"

            if todos_ok:
                fg_cor = "#064e3b"
                borda_cor = "#10b981"
                txt_cor = "#a7f3d0"
                icone_txt = f"📦 Material {mat}: {texto_qtd}  ✓"
            elif bloqueados > 0:
                fg_cor = "#450a0a"
                borda_cor = "#ef4444"
                txt_cor = "#fecaca"
                icone_txt = f"⛔ Material {mat}: {texto_qtd}  ({bloqueados} Bloqueado{'s' if bloqueados > 1 else ''})"
            else:
                fg_cor = "#451a03"
                borda_cor = "#f59e0b"
                txt_cor = "#fde68a"
                icone_txt = f"⚠️ Material {mat}: {texto_qtd}"

            card = ctk.CTkFrame(
                self.frame_resumo_materiais,
                fg_color=fg_cor,
                border_color=borda_cor,
                border_width=1,
                corner_radius=6
            )
            card.pack(side="left", padx=4, pady=2)

            ctk.CTkLabel(
                card,
                text=icone_txt,
                font=("Segoe UI", 11, "bold"),
                text_color=txt_cor
            ).pack(padx=10, pady=4)

    def _copiar_workforce(self):
        """Copia os seriais agrupados por material terminando com ';' para o WorkForce."""
        if not hasattr(self, "texto_workforce") or not self.texto_workforce.strip():
            messagebox.showwarning("Saída OSME", "Valide primeiro os seriais antes de copiar para o WorkForce.")
            return

        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(self.texto_workforce)
        self.btn_copiar_workforce.configure(text="✓ Copiado p/ WorkForce!", fg_color="#10b981")
        self.parent_frame.after(2000, lambda: self.btn_copiar_workforce.configure(text="📋 Copiar p/ WorkForce (;)", fg_color="#0284c7"))
        self._log("📋 Lista formatada para WorkForce copiada para a área de transferência!")

    def _abrir_modal_workforce(self):
        """Abre o pop-up com a lista formatada para WorkForce."""
        if not hasattr(self, "texto_workforce") or not self.texto_workforce.strip():
            messagebox.showwarning("Saída OSME", "Valide primeiro os seriais antes de visualizar a lista do WorkForce.")
            return
        ModalVisualizarWorkforce(self.parent_frame, self.texto_workforce, self.totais_materiais, self.paletas, self.obter_tema_atual)

    def _gerar_script_migo_vbs(self, equipe, osme, deposito, materiais_itens, data_lancamento=None):
        """
        Gera o código VBScript para execução da saída MIGO Z61 fiel ao script gravado no SAP.
        """
        linhas_vbs = [
            'On Error Resume Next',
            'Dim SapGuiAuto, app, connection, session',
            'Set SapGuiAuto = GetObject("SAPGUI")',
            'If Not IsObject(SapGuiAuto) Then',
            '    WScript.Echo "ERRO: SAP GUI não encontrado."',
            '    WScript.Quit 1',
            'End If',
            'Set app = SapGuiAuto.GetScriptingEngine',
            'If Not IsObject(app) Then',
            '    WScript.Echo "ERRO: Scripting Engine do SAP não disponível."',
            '    WScript.Quit 1',
            'End If',
            'Set connection = app.Children(0)',
            'If Not IsObject(connection) Then',
            '    WScript.Echo "ERRO: Nenhuma conexão ativa no SAP."',
            '    WScript.Quit 1',
            'End If',
            'Set session = connection.Children(0)',
            'If Not IsObject(session) Then',
            '    WScript.Echo "ERRO: Nenhuma sessão ativa no SAP."',
            '    WScript.Quit 1',
            'End If',
            '',
            'Sub FecharModais(maxVezes)',
            '    Dim mCount',
            '    mCount = 0',
            '    On Error Resume Next',
            '    Do While session.Children.Count > 1 And mCount < maxVezes',
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
            'session.findById("wnd[0]").maximize',
            'session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 800',
            'FecharModais 8',
            '',
            'Dim carrier',
            'carrier = "0003"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").setFocus',
            'If Err.Number <> 0 Then',
            '    Err.Clear',
            '    carrier = "0002"',
            'End If',
            '',
            '\' Configura cabeçalho MIGO Z61 Outros e envia Enter',
            'On Error Resume Next',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").setFocus',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A07"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").setFocus',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").key = "R10"',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").text = "Z61"',
            'If Err.Number <> 0 Then',
            '    Err.Clear',
            '    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODYNPRO-BWART").text = "Z61"',
            'End If',
            'session.findById("wnd[0]").sendVKey 0',
            'WScript.Sleep 500',
        ]

        dt_lcto_limpa = str(data_lancamento or "").strip()
        if dt_lcto_limpa:
            linhas_vbs.extend([
                f'PreencherDataLancamento carrier, "{dt_lcto_limpa}"',
                'WScript.Sleep 300',
            ])

        linhas_vbs.extend([
            '',
            '\' Garante detalhe do item aberto',
            'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press',
            'WScript.Sleep 300',
            'On Error GoTo 0',
            ''
        ])

        dep_clean = str(deposito or "CL03").strip()
        recebedor = f"EMPR 1{dep_clean}"

        for idx, item in enumerate(materiais_itens):
            mat_cod = item["material"]
            qtd = str(item["quantidade"]).replace(".", ",")
            seriais = item.get("seriais", [])
            eh_serializado = bool(seriais)
            eh_ultimo_item = (idx == len(materiais_itens) - 1)

            linhas_vbs.extend([
                f"' === ITEM {idx + 1}: Material {mat_cod} | Qtd: {qtd} | Serializado: {eh_serializado} ===",
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL").select',
                'On Error Resume Next',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").text = "{mat_cod}"',
                'If Err.Number <> 0 Then',
                '    Err.Clear',
                f'    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MATNR").text = "{mat_cod}"',
                'End If',
                'On Error GoTo 0',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                '',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").select',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = "{qtd}"',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                '',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-NAME1").text = "EMPR"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").text = "{dep_clean}"',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-WEMPF").text = "{recebedor}"',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                '',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT").select',
                'On Error Resume Next',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-AUFNR").text = "{osme}"',
                'If Err.Number <> 0 Then',
                '    Err.Clear',
                f'    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1002/ctxtCOBL-AUFNR").text = "{osme}"',
                'End If',
                'On Error GoTo 0',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
                '',
                'On Error Resume Next',
                f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-VORNR_AUF").text = "0010"',
                'If Err.Number <> 0 Then',
                '    Err.Clear',
                f'    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1002/ctxtCOBL-VORNR_AUF").text = "0010"',
                'End If',
                'On Error GoTo 0',
                'session.findById("wnd[0]").sendVKey 0',
                'WScript.Sleep 300',
            ])

            if eh_serializado:
                linhas_vbs.extend([
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").select',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_SEARCH").press',
                    'WScript.Sleep 400',
                    'session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press',
                    'WScript.Sleep 400',
                    'session.findById("wnd[1]/tbar[0]/btn[24]").press',
                    'WScript.Sleep 300',
                    'session.findById("wnd[1]/tbar[0]/btn[0]").press',
                    'WScript.Sleep 300',
                    'session.findById("wnd[1]/tbar[0]/btn[8]").press',
                    'WScript.Sleep 400',
                    'session.findById("wnd[0]/tbar[1]/btn[8]").press',
                    'WScript.Sleep 500',
                    '',
                    '\' Se abriu a lista grid com múltiplos itens, seleciona todos e adota (btn[47])',
                    '\' Se encontrou apenas 1 equipamento, o SAP adota direto e não exibe o grid',
                    'On Error Resume Next',
                    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""',
                    'session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll',
                    'WScript.Sleep 200',
                    'session.findById("wnd[0]/tbar[1]/btn[47]").press',
                    'If Err.Number <> 0 Then',
                    '    Err.Clear',
                    'End If',
                    'On Error GoTo 0',
                    'WScript.Sleep 400',
                    '',
                    '\' Checagem de segurança: se ainda estiver na tela "Seleção equipamento", nenhum equipamento foi encontrado',
                    f'Dim sTitle_{idx}',
                    f'sTitle_{idx} = ""',
                    'On Error Resume Next',
                    f'sTitle_{idx} = session.findById("wnd[0]").text',
                    'On Error GoTo 0',
                    f'If InStr(LCase(sTitle_{idx}), "equipamento") > 0 Then',
                    f'    Dim msg_selecao_{idx}',
                    f'    msg_selecao_{idx} = ""',
                    '    On Error Resume Next',
                    f'    msg_selecao_{idx} = session.findById("wnd[0]/sbar").text',
                    '    session.findById("wnd[0]/tbar[0]/btn[12]").press',
                    '    On Error GoTo 0',
                    f'    WScript.Echo "ERRO_SAP: Seriais não encontrados no SAP para o material {mat_cod}. " & msg_selecao_{idx}',
                    '    WScript.Quit 1',
                    'End If',
                ])

            if not eh_ultimo_item:
                linhas_vbs.extend([
                    '',
                    '\' Próximo Item (btnOK_NEXT_ITEM)',
                    'On Error Resume Next',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM").press',
                    'If Err.Number <> 0 Then',
                    '    Err.Clear',
                    '    session.findById("wnd[0]/tbar[1]/btn[13]").press',
                    'End If',
                    'WScript.Sleep 400',
                    f'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & carrier & "/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL").press',
                    'WScript.Sleep 300',
                    'On Error GoTo 0',
                    '',
                ])

        linhas_vbs.extend([
            "' Finalização: Verificar e Gravar",
            'session.findById("wnd[0]/tbar[1]/btn[7]").press',
            'WScript.Sleep 1000',
            '',
            'session.findById("wnd[0]/tbar[1]/btn[23]").press',
            'WScript.Sleep 2000',
            'Dim max_migo2, loop_migo2',
            'max_migo2 = 0',
            'Dim txtErrModal',
            'txtErrModal = ""',
            '',
            'Do While session.Children.Count > 1 And max_migo2 < 12',
            '    On Error Resume Next',
            '    If txtErrModal = "" Then',
            '        txtErrModal = Trim(session.findById("wnd[1]/usr/txtMESSTXT1").text)',
            '        If txtErrModal = "" Then txtErrModal = Trim(session.findById("wnd[1]").text)',
            '    End If',
            '    Err.Clear',
            '    session.findById("wnd[1]/tbar[0]/btn[0]").press',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]").sendVKey 0',
            '    End If',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press',
            '    End If',
            '    If Err.Number <> 0 Then',
            '        Err.Clear',
            '        session.findById("wnd[1]").close',
            '    End If',
            '    WScript.Sleep 350',
            '    max_migo2 = max_migo2 + 1',
            'Loop',
            '',
            'Dim msg_sbar',
            'msg_sbar = ""',
            'On Error Resume Next',
            'msg_sbar = Trim(session.findById("wnd[0]/sbar").text)',
            'On Error GoTo 0',
            '',
            'If txtErrModal <> "" And InStr(LCase(msg_sbar), "documento") = 0 Then',
            '    WScript.Echo "ERRO_SAP: " & txtErrModal',
            'End If',
            '',
            'WScript.Echo "MENSAGEM: " & msg_sbar',
            'WScript.Quit 0'
        ])

        return "\n".join(linhas_vbs)

    def _processar_saida_sap(self):
        if self.processando_sap:
            messagebox.showinfo("Saída OSME", "Já existe um processamento no SAP em andamento.")
            return

        todos_itens = list(self.seriais_validados_atuais) + list(self.itens_nao_serializados)
        if not todos_itens:
            messagebox.showwarning("Saída OSME", "Valide primeiro os seriais ou adicione materiais antes de processar a saída.")
            return

        tem_bloqueado = any(it.get("eh_bloqueado") for it in todos_itens)
        if tem_bloqueado:
            messagebox.showerror(
                "Saída OSME — Bloqueio Ativo",
                "Existem materiais marcados como 'JÁ FORNECIDO' na remessa!\n"
                "Remova-os antes de processar para evitar erro no SAP e duplicidade na planilha."
            )
            return

        eq = self.entry_equipe.get().strip().upper()
        osme = self.lbl_osme_numero.cget("text").strip()
        if not eq or osme in ["—", "Não Encontrada", "Sem OSME", ""]:
            messagebox.showwarning("Saída OSME", "Informe uma equipe válida com OSME cadastrada na planilha.")
            return

        if not verificar_sap_aberto():
            messagebox.showerror(
                "Saída OSME — SAP Fechado",
                "O SAP GUI precisa estar aberto e conectado em uma sessão ativa para realizar a saída MIGO.\n\n"
                "Por favor, abra o SAP, faça login e tente novamente."
            )
            return

        dep_info = self.lbl_deposito.cget("text").replace("Depósito:", "").strip()
        dep = dep_info if dep_info and dep_info != "—" else "CL03"

        # Monta lista consolidada de materiais com suas quantidades e unidades
        resumo_mats_dict = OrderedDict()
        for it in todos_itens:
            m = it["material"]
            if m not in resumo_mats_dict:
                resumo_mats_dict[m] = {"qtd_num": 0.0, "und": "un"}
            if it.get("eh_nao_serializado"):
                resumo_mats_dict[m]["und"] = it.get("unidade", "un")
                resumo_mats_dict[m]["qtd_num"] += float(it.get("quantidade_num", 1))
            else:
                resumo_mats_dict[m]["qtd_num"] += 1
                resumo_mats_dict[m]["und"] = "un"

        linhas_mats = []
        for m, dados in resumo_mats_dict.items():
            qtd_txt = f"{dados['qtd_num']:g}"
            linhas_mats.append(f"  ▸ {m} — {qtd_txt} {dados['und']}")
        texto_materiais_bloco = "\n".join(linhas_mats)

        data_lcto = self.obter_data_lancamento()
        if data_lcto is None:
            return

        txt_dt = f"• Data de Lançamento (MIGO): {data_lcto} (Fechamento SAP)\n" if data_lcto else "• Data de Lançamento (MIGO): Padrão (Hoje)\n"

        msg_confirma = (
            f"Confirmar Saída no SAP MIGO?\n\n"
            f"• Equipe: {eq}\n"
            f"• OSME: {osme}\n"
            f"• Depósito: {dep}\n"
            f"{txt_dt}\n"
            f"• Materiais a Baixar:\n"
            f"{texto_materiais_bloco}\n\n"
            f"Total: {len(todos_itens)} item(ns)\n\n"
            "Deseja iniciar a execução automática no SAP agora?"
        )
        if not messagebox.askyesno("Confirmar Saída OSME", msg_confirma):
            return

        self.processando_sap = True
        self._definir_estado_botoes_processar(True)
        self._status(f"Iniciando Saída MIGO Z61 para {eq} (OSME {osme})...")
        self._log(f"🚀 Iniciando fluxo automático SAP MIGO para equipe {eq}...")
        if data_lcto:
            self._log(f"📅 Data de Lançamento configurada: {data_lcto}")

        threading.Thread(target=self._executar_saida_sap_thread, args=(eq, osme, dep, todos_itens, data_lcto), daemon=True).start()

    def _executar_saida_sap_thread(self, equipe, osme, deposito, todos_itens, data_lancamento=None):
        try:
            materiais_para_sap = OrderedDict()
            for it in todos_itens:
                mat = it["material"]
                if mat not in materiais_para_sap:
                    materiais_para_sap[mat] = {
                        "material": mat,
                        "deposito": it.get("deposito") or deposito or "CL03",
                        "seriais": [],
                        "quantidade": 0.0,
                        "eh_nao_serializado": False
                    }
                if it.get("eh_nao_serializado"):
                    materiais_para_sap[mat]["eh_nao_serializado"] = True
                    materiais_para_sap[mat]["quantidade"] += float(it.get("quantidade_num", 1))
                else:
                    materiais_para_sap[mat]["seriais"].append(it["serial"])
                    materiais_para_sap[mat]["quantidade"] = len(materiais_para_sap[mat]["seriais"])

            lista_itens_sap = list(materiais_para_sap.values())

            # Coleta TODOS os seriais da remessa para o Clipboard do Windows.
            # O SAP MIGO realiza a filtragem inteligente automática na tela de Seleção de Equipamento pelo código do Material atual!
            todos_seriais = [str(it["serial"]).strip() for it in todos_itens if it.get("serial")]
            if todos_seriais:
                copiar_texto_clipboard_resiliente("\r\n".join(todos_seriais))
                self._log_threadsafe(f"📋 {len(todos_seriais)} serial(is) copiados para a Área de Transferência (Filtro Inteligente SAP MIGO).")

            vbs_code = self._gerar_script_migo_vbs(equipe, osme, deposito, lista_itens_sap, data_lancamento=data_lancamento)

            fd, tmp_vbs = tempfile.mkstemp(prefix="migo_z61_", suffix=".vbs", text=True)
            with os.fdopen(fd, "w", encoding="cp1252", errors="replace") as f:
                f.write(vbs_code)

            self._log_threadsafe("⏳ Executando automação no SAP GUI via cscript...")
            creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

            proc = subprocess.run(
                ["cscript.exe", "//nologo", tmp_vbs],
                capture_output=True,
                text=True,
                encoding="cp850",
                timeout=None,
                creationflags=creationflags
            )

            try:
                os.remove(tmp_vbs)
            except OSError:
                pass

            saida_stdout = proc.stdout or ""
            saida_stderr = proc.stderr or ""
            saida_completa = f"{saida_stdout}\n{saida_stderr}".strip()

            self._log_threadsafe(f"SAP Retorno: {saida_stdout.strip()}")

            if proc.returncode != 0 and "ERRO:" in saida_completa:
                msg_err = saida_completa.split("ERRO:")[-1].strip()
                raise RuntimeError(f"Erro no SAP GUI: {msg_err}")

            erro_sap_modal = ""
            for linha in saida_completa.splitlines():
                if "ERRO_SAP:" in linha:
                    erro_sap_modal = linha.split("ERRO_SAP:", 1)[1].strip()

            doc_sap = None
            match_doc = re.search(r"\b([0-9]{10})\b", saida_completa)
            if match_doc:
                doc_sap = match_doc.group(1)

            msg_sbar = ""
            for linha in saida_completa.splitlines():
                if "MENSAGEM:" in linha:
                    msg_sbar = linha.split("MENSAGEM:", 1)[1].strip()

            if not doc_sap:
                m2 = re.search(r"(\d{8,12})", msg_sbar)
                if m2:
                    doc_sap = m2.group(1)

            if erro_sap_modal and not doc_sap:
                self._log_threadsafe(f"⚠️ Erro reportado pelo SAP: {erro_sap_modal}")
                self.parent_frame.after(0, lambda: self._finalizar_com_aviso_sap(f"Erro no SAP MIGO:\n{erro_sap_modal}"))
                return

            if not doc_sap and ("gravado" not in msg_sbar.lower() and "salvo" not in msg_sbar.lower()):
                alerta_msg = msg_sbar if msg_sbar else (erro_sap_modal or "Nenhum documento gerado pelo SAP.")
                self.parent_frame.after(0, lambda: self._finalizar_com_aviso_sap(alerta_msg))
                return

            if not doc_sap:
                doc_sap = "REGISTRADO"

            self._log_threadsafe(f"✅ Documento SAP MIGO Gerado: {doc_sap}")
            self._log_threadsafe("💾 Atualizando planilha Excel (baixa de estoque e registro em fornecidos)...")

            sucesso_planilha, modo_planilha = self._atualizar_planilha_pos_sap(equipe, doc_sap, todos_itens)

            self.parent_frame.after(
                0,
                lambda: self._concluir_processamento_sucesso(equipe, osme, doc_sap, todos_itens, sucesso_planilha, modo_planilha)
            )

        except Exception as e:
            self.parent_frame.after(0, lambda err=str(e): self._finalizar_com_erro(err))

    def _atualizar_planilha_pos_sap(self, equipe, doc_sap, itens_processados):
        """
        Remove seriais da aba ESTOQUE e insere na aba FORNECIDOS com DOC. SAP.
        Se o Excel estiver aberto no PC do usuário, manipula via win32com para salvar sem conflito.
        Se estiver fechado, utiliza openpyxl.
        Salva backup em contingência local e atualiza cache em memória imediatamente.
        """
        caminho = self.caminho_planilha
        dt_hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        data_hoje = dt_hoje.strftime("%d/%m/%Y")
        seriais_set = {
            normalizar_texto(s["serial"]): s 
            for s in itens_processados 
            if s.get("serial") and not s.get("eh_nao_serializado")
        }

        atualizado_com_sucesso = False
        metodo_usado = ""

        # 0. Backup Preventivo Automático
        bak_criado = criar_backup_preventivo(caminho, prefixo="PRE_SAIDA_OSME")
        if bak_criado:
            self._log_threadsafe(f"🛡️ Backup preventivo criado: {os.path.basename(bak_criado)}")

        # 1. Tentativa via Excel aberto no Windows (COM)
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            wb_com = None
            for wb in xl.Workbooks:
                try:
                    if os.path.abspath(wb.FullName).lower() == os.path.abspath(caminho).lower():
                        wb_com = wb
                        break
                except Exception:
                    pass

            if wb_com is not None:
                # Deleta de ESTOQUE de baixo para cima
                try:
                    ws_est = wb_com.Sheets("ESTOQUE")
                    last_row = ws_est.UsedRange.Rows.Count + 1
                    for r in range(last_row, 1, -1):
                        val = str(ws_est.Cells(r, 1).Value or "").strip()
                        if normalizar_texto(val) in seriais_set:
                            ws_est.Rows(r).Delete()
                except Exception as e_est:
                    self._log_threadsafe(f"Aviso ESTOQUE via Excel COM: {e_est}")

                # Insere em FORNECIDOS
                try:
                    ws_forn = wb_com.Sheets("FORNECIDOS")
                    next_row = ws_forn.UsedRange.Rows.Count + 1
                    while str(ws_forn.Cells(next_row, 1).Value or "").strip():
                        next_row += 1

                    for s in itens_processados:
                        if s.get("eh_nao_serializado"):
                            continue
                        ws_forn.Cells(next_row, 1).Value = s["serial"]
                        ws_forn.Cells(next_row, 2).Value = s["material"]
                        ws_forn.Cells(next_row, 3).Value = s["deposito"]
                        ws_forn.Cells(next_row, 4).Value = equipe
                        try:
                            ws_forn.Cells(next_row, 5).NumberFormat = "dd/mm/yyyy"
                        except Exception:
                            pass
                        ws_forn.Cells(next_row, 5).Value = dt_hoje
                        ws_forn.Cells(next_row, 6).Value = doc_sap
                        next_row += 1
                except Exception as e_forn:
                    self._log_threadsafe(f"Aviso FORNECIDOS via Excel COM: {e_forn}")

                wb_com.Save()
                atualizado_com_sucesso = True
                metodo_usado = "Excel Aberto (COM)"
        except Exception:
            pass

        # 2. Se não estava aberto via COM, utiliza openpyxl
        if not atualizado_com_sucesso:
            wb = None
            try:
                wb = load_workbook(caminho, keep_vba=caminho.lower().endswith(".xlsm"))

                # Remove de ESTOQUE
                if "ESTOQUE" in wb.sheetnames:
                    ws_est = wb["ESTOQUE"]
                    linhas_para_deletar = []
                    for r_idx in range(ws_est.max_row, 1, -1):
                        val = str(ws_est.cell(row=r_idx, column=1).value or "").strip()
                        if normalizar_texto(val) in seriais_set:
                            linhas_para_deletar.append(r_idx)
                    for r_del in linhas_para_deletar:
                        ws_est.delete_rows(r_del, 1)

                # Insere em FORNECIDOS
                if "FORNECIDOS" in wb.sheetnames:
                    ws_forn = wb["FORNECIDOS"]
                    for s in itens_processados:
                        if s.get("eh_nao_serializado"):
                            continue
                        ws_forn.append([
                            s["serial"],
                            s["material"],
                            s["deposito"],
                            equipe,
                            dt_hoje,
                            doc_sap
                        ])
                        try:
                            ws_forn.cell(row=ws_forn.max_row, column=5).number_format = "dd/mm/yyyy"
                        except Exception:
                            pass

                wb.save(caminho)
                atualizado_com_sucesso = True
                metodo_usado = "openpyxl (Arquivo Fechado)"
            except Exception as e_openpyxl:
                self._log_threadsafe(f"❌ Erro ao atualizar planilha diretamente: {e_openpyxl}")
            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

        # Se não foi atualizado com sucesso (planilha em uso na rede ou aberta em Somente Leitura)
        if not atualizado_com_sucesso:
            motivo = "Planilha em uso por outro usuário no Excel da rede (Arquivo Bloqueado)"
            try:
                if wb_com is not None and getattr(wb_com, "ReadOnly", False):
                    motivo = "Planilha aberta no Excel local em modo 'Somente Leitura' (bloqueada na rede)"
            except Exception:
                pass

            pid = adicionar_pendencia(
                tipo="saida_osme",
                caminho_planilha=caminho,
                dados_operacao={
                    "equipe": equipe,
                    "doc_sap": doc_sap,
                    "data": data_hoje,
                    "deposito": self.lbl_deposito.cget("text").replace("Depósito:", "").strip() or "CL03",
                    "itens": [
                        {
                            "serial": it.get("serial"),
                            "material": it.get("material"),
                            "deposito": it.get("deposito") or "CL03",
                            "eh_nao_serializado": it.get("eh_nao_serializado", False)
                        }
                        for it in itens_processados
                    ]
                },
                motivo=motivo
            )
            metodo_usado = "Fila de Sincronização Automática (Planilha Bloqueada)"
            self._log_threadsafe(f"⏳ {motivo}. Operação #{pid} salva na fila de pendências para gravação automática.")
            self.parent_frame.after(0, lambda: self._atualizar_banner_pendencias())

        # 3. Atualiza Cache em Memória imediatamente
        for s_norm, item_s in seriais_set.items():
            if s_norm in self.dados_estoque:
                del self.dados_estoque[s_norm]
            self.dados_fornecidos[s_norm] = {
                "serial": item_s["serial"],
                "material": item_s["material"],
                "deposito": item_s["deposito"],
                "equipe": equipe,
                "data": data_hoje,
                "doc_sap": doc_sap
            }

        # 4. Salva contingência local de segurança e Backup Diário em Excel
        status_backup_txt = "Gravado na Rede" if atualizado_com_sucesso else "Pendente na Fila"
        try:
            ok_bk, arq_bk = gravar_saida_backup_diario(
                equipe=equipe,
                doc_sap=doc_sap,
                data_lcto=data_hoje,
                itens=itens_processados,
                status_mestre=status_backup_txt,
                caminho_mestre=caminho
            )
            if ok_bk:
                self._log_threadsafe(f"🛡️ Backup diário salvo em Documentos: {os.path.basename(arq_bk)}")
        except Exception as e_bk:
            self._log_threadsafe(f"Aviso backup diário: {e_bk}")

        try:
            pasta_dados = obter_diretorio_dados()
            arq_contingencia = os.path.join(pasta_dados, "contingencia_saidas_osme.json")
            historico = []
            if os.path.exists(arq_contingencia):
                try:
                    with open(arq_contingencia, "r", encoding="utf-8") as f:
                        historico = json.load(f)
                except Exception:
                    historico = []
            historico.append({
                "timestamp": datetime.now().isoformat(),
                "equipe": equipe,
                "doc_sap": doc_sap,
                "data": data_hoje,
                "itens": [
                    {"serial": it["serial"], "material": it["material"], "deposito": it["deposito"], "eh_ns": it.get("eh_nao_serializado", False)}
                    for it in itens_processados
                ],
                "planilha_atualizada": atualizado_com_sucesso,
                "metodo": metodo_usado
            })
            with open(arq_contingencia, "w", encoding="utf-8") as f:
                json.dump(historico, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

        return atualizado_com_sucesso, metodo_usado

    def _iniciar_timer_pendencias(self):
        """Inicia o ciclo de verificação em segundo plano das pendências da planilha."""
        self._verificar_timer_pendencias()

    def _verificar_timer_pendencias(self):
        """Verifica a cada 20 segundos se a planilha foi liberada para gravar pendências."""
        try:
            qtd = obter_quantidade_pendencias()
            self._atualizar_banner_pendencias(qtd)
            if qtd > 0:
                sucessos, restantes, msg = processar_todas_pendencias()
                if sucessos > 0:
                    self._log_threadsafe(f"🎉 Sincronização Automática: {sucessos} operação(ões) gravada(s) na planilha com sucesso!")
                    self._status(f"🎉 {sucessos} operação(ões) gravada(s) na planilha automaticamente!")
                    self._atualizar_banner_pendencias(restantes)
                    self.carregar_planilha_async()
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
                self.f_banner_pendencias.pack(fill="x", padx=12, pady=(0, 8))
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
            self.carregar_planilha_async()
        elif sucessos > 0 and restantes > 0:
            messagebox.showinfo(
                "Sincronização Parcial",
                f"✓ {sucessos} gravados com sucesso.\nRestam {restantes} que continuam em uso na rede."
            )
            self.carregar_planilha_async()
        else:
            messagebox.showwarning(
                "Planilha em Uso",
                "A planilha continua aberta por outro usuário na rede.\n\n"
                "O N.E.X.U.S continuará tentando gravar automaticamente a cada 20 segundos em segundo plano."
            )

    def _abrir_modal_pendencias(self):
        """Abre o modal para inspecionar, copiar dados ou descartar pendências."""
        ModalVisualizarPendencias(
            self.parent_frame,
            callback_apos_sincronizacao=lambda: [self._atualizar_banner_pendencias(), self.carregar_planilha_async()],
            paletas=self.paletas,
            obter_tema_atual=self.obter_tema_atual
        )

    def _concluir_processamento_sucesso(self, equipe, osme, doc_sap, todos_itens, sucesso_planilha, modo_planilha):
        self.processando_sap = False
        self._definir_estado_botoes_processar(False)
        self._status(f"Saída Concluída! Doc SAP: {doc_sap} | Equipe: {equipe}")
        try:
            incrementar_metrica("saidas_osme", 1)
        except Exception:
            pass

        # Atualiza badges de contagem
        self.badge_estoque.configure(text=f"📦 Itens Estoque: {len(self.dados_estoque)}")
        self.badge_fornecidos.configure(text=f"🏷️ Já Fornecidos: {len(self.dados_fornecidos)}")

        # Abre modal moderno com botão de cópia rápida
        ModalSucessoSaidaOSME(
            self.parent_frame,
            doc_sap,
            equipe,
            osme,
            todos_itens,
            sucesso_planilha,
            modo_planilha,
            self.paletas,
            self.obter_tema_atual,
            totais_materiais=getattr(self, "totais_materiais", {})
        )
        self._limpar_campos()

    def preparar_saida_via_chat(self, equipe=None, seriais=None, material_ns=None, quantidade_ns=None, data_lancamento=None):
        """
        Preenche os dados na tela de Saída OSME e confere previamente sem disparar no SAP.
        Retorna dicionário com status e resumo dos itens preparados.
        """
        if equipe:
            self.entry_equipe.delete(0, tk.END)
            self.entry_equipe.insert(0, str(equipe).strip().upper())
            self._consultar_equipe_atual()

        if data_lancamento and hasattr(self, "entry_data_lancamento"):
            self.entry_data_lancamento.delete(0, tk.END)
            self.entry_data_lancamento.insert(0, str(data_lancamento).strip())

        if seriais:
            texto_seriais = "\n".join(seriais) if isinstance(seriais, list) else str(seriais)
            self.txt_seriais.delete("1.0", tk.END)
            self.txt_seriais.insert("1.0", texto_seriais)
            self.validar_seriais_acao()

        if material_ns and quantidade_ns:
            if isinstance(material_ns, list) and isinstance(quantidade_ns, list):
                for m_item, q_item in zip(material_ns, quantidade_ns):
                    self.entry_mat_ns.delete(0, tk.END)
                    self.entry_mat_ns.insert(0, str(m_item).strip())
                    self._ao_digitar_material_ns()
                    self.entry_qtd_ns.delete(0, tk.END)
                    self.entry_qtd_ns.insert(0, str(q_item).strip())
                    self._adicionar_material_nao_serializado()
            else:
                self.entry_mat_ns.delete(0, tk.END)
                self.entry_mat_ns.insert(0, str(material_ns).strip())
                self._ao_digitar_material_ns()
                self.entry_qtd_ns.delete(0, tk.END)
                self.entry_qtd_ns.insert(0, str(quantidade_ns).strip())
                self._adicionar_material_nao_serializado()

        todos_itens = list(self.seriais_validados_atuais) + list(self.itens_nao_serializados)
        eq_atual = self.entry_equipe.get().strip().upper()
        osme_atual = self.lbl_osme_numero.cget("text").strip()
        dep_atual = self.lbl_deposito.cget("text").replace("Depósito:", "").strip()

        resumo_mats = []
        mats_map = OrderedDict()
        for it in todos_itens:
            m = it["material"]
            if m not in mats_map:
                mats_map[m] = {"qtd": 0.0, "und": it.get("unidade", "un")}
            if it.get("eh_nao_serializado"):
                mats_map[m]["qtd"] += float(it.get("quantidade_num", 1))
            else:
                mats_map[m]["qtd"] += 1
        for m, d in mats_map.items():
            resumo_mats.append(f"{m} — {d['qtd']:g} {d['und']}")

        return {
            "sucesso": True,
            "equipe": eq_atual,
            "osme": osme_atual,
            "deposito": dep_atual,
            "total_itens": len(todos_itens),
            "resumo_materiais": resumo_mats,
            "tem_bloqueados": any(it.get("eh_bloqueado") for it in todos_itens)
        }

    def executar_saida_via_chat(self):
        """Executa a saída no SAP diretamente após confirmação do usuário via Chatbot."""
        todos_itens = list(self.seriais_validados_atuais) + list(self.itens_nao_serializados)
        if not todos_itens:
            return {"erro": "Nenhum material ou serial carregado para saída no SAP. Prepare os dados primeiro."}

        tem_bloqueado = any(it.get("eh_bloqueado") for it in todos_itens)
        if tem_bloqueado:
            return {"erro": "Existem materiais marcados como 'JÁ FORNECIDO' na remessa. Remova-os na aba antes de enviar ao SAP."}

        eq = self.entry_equipe.get().strip().upper()
        osme = self.lbl_osme_numero.cget("text").strip()
        if not eq or osme in ["—", "Não Encontrada", "Sem OSME", ""]:
            return {"erro": "Informe uma equipe válida com OSME cadastrada na planilha."}

        if not verificar_sap_aberto():
            return {"erro": "O SAP GUI precisa estar aberto e conectado em uma sessão ativa."}

        if self.processando_sap:
            return {"erro": "Já existe um processamento de saída no SAP em andamento."}

        dep_info = self.lbl_deposito.cget("text").replace("Depósito:", "").strip()
        dep = dep_info if dep_info and dep_info != "—" else "CL03"

        data_lcto = self.obter_data_lancamento() if hasattr(self, "entry_data_lancamento") else ""
        if data_lcto is None:
            data_lcto = ""

        self.processando_sap = True
        self._definir_estado_botoes_processar(True)
        self._status(f"Iniciando Saída MIGO Z61 para {eq} (OSME {osme})...")
        self._log(f"🚀 Iniciando fluxo automático SAP MIGO para equipe {eq} (disparado via confirmação no Chatbot)...")
        if data_lcto:
            self._log(f"📅 Data de Lançamento configurada: {data_lcto}")

        threading.Thread(target=self._executar_saida_sap_thread, args=(eq, osme, dep, todos_itens, data_lcto), daemon=True).start()
        return {
            "sucesso": True,
            "status": f"Execução de Saída MIGO Z61 iniciada no SAP para a equipe {eq} (OSME {osme}) com {len(todos_itens)} item(ns)."
        }

    def _finalizar_com_aviso_sap(self, alerta_msg):
        self.processando_sap = False
        self._definir_estado_botoes_processar(False)
        self._status("Saída OSME: Retorno do SAP GUI verificado.")
        self._log(f"⚠️ Aviso do SAP: {alerta_msg}")
        messagebox.showwarning(
            "Saída OSME — Aviso SAP",
            f"O SAP retornou a seguinte mensagem na barra inferior:\n\n{alerta_msg}\n\nNenhuma alteração foi realizada na planilha."
        )

    def _finalizar_com_erro(self, err_msg):
        self.processando_sap = False
        self._definir_estado_botoes_processar(False)
        self._status("Erro durante execução no SAP.")
        self._log(f"❌ Erro SAP: {err_msg}")
        messagebox.showerror("Saída OSME — Erro", f"Ocorreu um erro durante a execução no SAP:\n\n{err_msg}")

    def _copiar_tabela(self):
        itens = self.tree.get_children()
        if not itens:
            return
        linhas = ["Item\tSerial\tMaterial\tDepósito\tEquipe\tOSME\tStatus"]
        for it in itens:
            vals = self.tree.item(it, "values")
            linhas.append("\t".join(str(v) for v in vals))
        texto = "\n".join(linhas)
        self.parent_frame.clipboard_clear()
        self.parent_frame.clipboard_append(texto)
        messagebox.showinfo("Saída OSME", "Tabela copiada para a área de transferência em formato compatível com Excel!")

    def _limpar_campos(self):
        self.txt_seriais.delete("1.0", tk.END)
        if hasattr(self, "entry_data_lancamento"):
            self.entry_data_lancamento.delete(0, tk.END)
        self.entry_mat_ns.delete(0, tk.END)
        self.entry_qtd_ns.delete(0, tk.END)
        self.lbl_desc_mat_ns.configure(
            text="Digite o código para autocompletar descrição e unidade da aba CADASTRO...",
            text_color="#94a3b8"
        )
        self.itens_nao_serializados = []
        self.seriais_validados_atuais = []
        self._recarregar_chips_ns()
        self._renderizar_saida()

    def aplicar_tema(self, tema_nome):
        cores = self.paletas.get(tema_nome, next(iter(self.paletas.values()), DEFAULT_PALETAS["Azul Delray"]))
        bg = cores.get("bg", "#090d16")
        container = cores.get("container", "#111827")
        inp = cores.get("input", "#1a2436")
        texto = cores.get("texto", "#f8fafc")
        accent = cores.get("accent", "#00d2ff")

        try:
            self.parent_frame.configure(fg_color=bg)
            if hasattr(self, "f_barra_fixa") and self.f_barra_fixa:
                self.f_barra_fixa.configure(fg_color=container)
            if hasattr(self, "frame_topo") and self.frame_topo:
                self.frame_topo.configure(fg_color=container)
            if hasattr(self, "frame_planilha") and self.frame_planilha:
                self.frame_planilha.configure(fg_color=container)
            if hasattr(self, "frame_entradas") and self.frame_entradas:
                self.frame_entradas.configure(fg_color=container)
            if hasattr(self, "card_osme_info") and self.card_osme_info:
                self.card_osme_info.configure(fg_color=inp)
            if hasattr(self, "card_nao_serializados") and self.card_nao_serializados:
                self.card_nao_serializados.configure(fg_color=inp)
            if hasattr(self, "lbl_osme_numero") and self.lbl_osme_numero:
                self.lbl_osme_numero.configure(text_color=accent)
            if hasattr(self, "frame_tabela") and self.frame_tabela:
                self.frame_tabela.configure(fg_color=container)
            if hasattr(self, "entry_equipe") and self.entry_equipe:
                self.entry_equipe.configure(fg_color=inp, text_color=texto)
            if hasattr(self, "entry_data_lancamento") and self.entry_data_lancamento:
                self.entry_data_lancamento.configure(fg_color=inp, text_color=texto)
            if hasattr(self, "entry_mat_ns") and self.entry_mat_ns:
                self.entry_mat_ns.configure(fg_color=inp, text_color=texto)
            if hasattr(self, "entry_qtd_ns") and self.entry_qtd_ns:
                self.entry_qtd_ns.configure(fg_color=inp, text_color=texto)
            if hasattr(self, "card_log") and self.card_log:
                self.card_log.configure(fg_color=container)
            if hasattr(self, "txt_seriais") and self.txt_seriais:
                self.txt_seriais.configure(fg_color=inp, text_color=texto)
            if hasattr(self, "txt_log") and self.txt_log:
                self.txt_log.configure(fg_color=inp, text_color=texto)
        except Exception:
            pass

        try:
            estilo = ttk.Style()
            if "clam" in estilo.theme_names():
                estilo.theme_use("clam")

            estilo.configure(
                "SaidaOSME.Treeview",
                background=inp,
                foreground=texto,
                fieldbackground=inp,
                borderwidth=0,
                relief="flat",
                font=("Segoe UI", 9),
                rowheight=26
            )
            estilo.configure(
                "SaidaOSME.Treeview.Heading",
                background=container,
                foreground="#38bdf8",
                font=("Segoe UI", 9, "bold"),
                borderwidth=1,
                relief="flat"
            )
            estilo.map(
                "SaidaOSME.Treeview",
                background=[("selected", accent)],
                foreground=[("selected", "#ffffff")]
            )
            estilo.map(
                "SaidaOSME.Treeview.Heading",
                background=[("active", cores.get("hover", "#0284c7"))],
                foreground=[("active", "#ffffff")]
            )
            if hasattr(self, "tree") and self.tree:
                self.tree.tag_configure("ok", foreground="#34d399")
                self.tree.tag_configure("bloqueado", foreground="#fca5a5", background="#450a0a")
                self.tree.tag_configure("alerta", foreground="#fcd34d")
                self.tree.tag_configure("padrao", foreground=texto)
        except Exception:
            pass


def construir_aba_saida_osme(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    return AbaSaidaOSME(parent_frame, label_status_global, paletas, obter_tema_atual)
