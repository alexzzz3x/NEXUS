# -*- coding: utf-8 -*-
"""
Módulo de Controle de Material — Nexus / Conecta Hub
Integração completa do formulário fmain (frmMain) da planilha CONTROLE_SERIAIS.vBeta.xlsm.

Funcionalidades:
- Interface dinâmica, limpa e moderna (sem scrollbars externas feias nem dropdowns gigantes).
- Autocomplete compacto e elegante para Equipes (Origem, Destino e Consulta Lateral) exibindo até 7 sugestões logo abaixo do campo.
- Auto-preenchimento instantâneo da OSME (Ordem) e Depósito ao digitar ou selecionar a equipe.
- Sem textos crus de 'CTkComboBox' (placeholders elegantes e campos limpos).
- Movimentações: Saída para OSME, Entrada de Materiais, Devolução, Transferência entre Equipes,
  Transferência entre Depósitos, Atualização de OSMES, Saldo DP e Mont-Inst.
- Validações: Duplicidade de seriais, presença em estoque/fornecidos, saldo em reserva SAP.
- Automação SAP: MIGO Z61, MIGO para Ordem (R08), MIGO Devolução Z62, ZUDWM_237, ZM29, IQ09.
- Painel de Consulta: Saldos em poder da equipe x Limites por segmento (Ligação, Religa, Corte, TMA, etc.).
- Alimentação Excel: Atualização fiel das abas ESTOQUE, FORNECIDOS e CADASTRO com keep_vba=True
  e backup automático preventivo. NENHUM macro do Excel é alterado!
"""

import os
import re
import shutil
import json
import threading
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk
import openpyxl

from utils_os import obter_diretorio_executavel, obter_caminho_recurso
from utils_logger import log_error, log_info
from utils_ui import show_toast, play_sound

# ==============================================================================
# CONFIGURAÇÃO DE CAMINHOS DA PLANILHA
# ==============================================================================

CONFIG_CONTROLE_FILE = os.path.join(obter_diretorio_executavel(), "data", "config_controle_material.json")

CAMINHOS_PADRAO_EXCEL = [
    r"X:\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm",
    r"\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm",
    os.path.join(os.path.expanduser("~"), "Desktop", "CONTROLE_SERIAIS.vBeta.xlsm")
]

def obter_caminho_planilha_controle():
    """Retorna o caminho configurado ou o primeiro padrão existente."""
    if os.path.exists(CONFIG_CONTROLE_FILE):
        try:
            with open(CONFIG_CONTROLE_FILE, "r", encoding="utf-8") as f:
                dados = json.load(f)
                caminho = dados.get("caminho_planilha", "").strip()
                if caminho and os.path.exists(caminho):
                    return caminho
        except Exception as e:
            log_error(f"Erro ao ler config_controle_material.json: {e}")

    for padrao in CAMINHOS_PADRAO_EXCEL:
        if os.path.exists(padrao):
            return padrao

    return CAMINHOS_PADRAO_EXCEL[0]

def salvar_caminho_planilha_controle(caminho):
    try:
        os.makedirs(os.path.dirname(CONFIG_CONTROLE_FILE), exist_ok=True)
        with open(CONFIG_CONTROLE_FILE, "w", encoding="utf-8") as f:
            json.dump({"caminho_planilha": caminho}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_error(f"Erro ao salvar caminho da planilha: {e}")


# ==============================================================================
# TIPOS DE MOVIMENTO (Enum correspondente ao VBA)
# ==============================================================================

TIPOS_MOVIMENTO = [
    "Saída Para OSME",                # Padrão mais comum do dia a dia
    "Entrada de Materiais",           # Entrada de fornecedor/lote
    "Devolução de Materiais",         # Devolução da equipe
    "Transferência Entre Equipes",     # Transf entre equipes
    "Transferência Entre Depósitos",   # Transf CL03 <-> CL04 ou CI
    "Atualizar Osmes",                 # Rotina SAP
    "Atualizar Saldo DP",              # Rotina SAP
    "Remover Mont-Inst"               # Rotina SAP
]

COLUNAS_LIMITES = {
    "LIGAÇÃO": 16,       # Coluna P
    "RELIGA": 17,        # Coluna Q
    "CORTE": 18,         # Coluna R
    "TMA": 19,           # Coluna S
    "CLANDESTINA": 20,   # Coluna T
    "AFERIÇÃO": 21       # Coluna U
}


# ==============================================================================
# MOTOR SAP (win32com)
# ==============================================================================

def conectar_sap():
    """Conecta na sessão ativa do SAP GUI Scripting com busca resiliente e segura."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            return None
            
        application = getattr(sap_gui_auto, "GetScriptingEngine", None)
        if callable(application):
            application = application()
        if not application:
            return None

        # 1. Tenta ActiveSession direto do Application
        try:
            act = getattr(application, "ActiveSession", None)
            if act:
                return act
        except Exception:
            pass

        # 2. Varre conexões e sessões ativas
        try:
            conns_count = getattr(application.Children, "Count", 0)
            for i in range(conns_count):
                conn = application.Children(i)
                if not conn:
                    continue
                try:
                    conn_act = getattr(conn, "ActiveSession", None)
                    if conn_act:
                        return conn_act
                except Exception:
                    pass
                sess_count = getattr(conn.Children, "Count", 0)
                for j in range(sess_count):
                    sess = conn.Children(j)
                    if sess:
                        return sess
        except Exception:
            pass

        return None
    except Exception as e:
        log_error(f"Erro ao conectar com o SAP: {e}")
        return None

def sap_verificar_reserva_ordem(session, ordem, deposito, materiais_resumo):
    """Simula /NMIGO A07/R08 e checa reservas da ordem."""
    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMIGO"
        session.findById("wnd[0]").sendVKey(0)

        try:
            btn_det = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL", False)
            if btn_det:
                btn_det.press()
        except Exception:
            pass

        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R08"
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").Text = str(ordem)
        session.findById("wnd[0]").sendVKey(0)

        if session.findById("wnd[1]", False):
            msg = session.findById("wnd[1]/usr/lbl[10,3]", False)
            texto_erro = msg.Text if msg else "Erro na ordem no SAP."
            return False, texto_erro

        return True, ""
    except Exception as e:
        return False, f"Falha na verificação de reserva SAP: {e}"


def sap_migo_to_osme(session, lista_saida, ordem, deposito, data_lancamento, is_serial=True):
    """Executa MIGO Z61 para OSME (agrupado por Material e Ordem)."""
    import win32clipboard
    linhas = [l.strip() for l in lista_saida.splitlines() if l.strip()]
    if not linhas:
        return ""

    grupos_mat = {}
    grupos_seriais = {}
    grupos_qtd = {}

    for l in linhas:
        partes = l.split("\t")
        if len(partes) >= 2:
            if is_serial:
                ser = partes[0].strip()
                mat = partes[1].strip()
                cur_ord = partes[2].strip() if len(partes) >= 3 and partes[2].strip() else ordem
                key = f"{mat}|{cur_ord}"
                if key not in grupos_mat:
                    grupos_mat[key] = (mat, cur_ord)
                    grupos_seriais[key] = []
                    grupos_qtd[key] = 0
                grupos_seriais[key].append(ser)
                grupos_qtd[key] += 1
            else:
                try:
                    qtd = int(float(partes[0].strip()))
                except Exception:
                    qtd = 1
                mat = partes[1].strip()
                cur_ord = ordem
                key = f"{mat}|{cur_ord}"
                if key not in grupos_mat:
                    grupos_mat[key] = (mat, cur_ord)
                    grupos_qtd[key] = 0
                grupos_qtd[key] += qtd

    if not grupos_mat:
        return ""

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMIGO"
    session.findById("wnd[0]").sendVKey(0)

    try:
        btn_det = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL", False)
        if btn_det:
            btn_det.press()
    except Exception:
        pass

    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").Key = "A07"
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC").Key = "R10"
    
    try:
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").Text = "Z61"
    except Exception:
        pass

    data_formatada = data_lancamento.replace("/", ".")
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/ctxtGOHEAD-BUDAT").Text = data_formatada

    keys = list(grupos_mat.keys())
    total_grupos = len(keys)

    for k, key in enumerate(keys):
        mat, cur_ord = grupos_mat[key]
        qtd = grupos_qtd[key]

        # Material
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL").Select()
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX").Text = mat
        session.findById("wnd[0]").sendVKey(0)

        # Quantidade
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").Select()
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").Text = str(qtd)
        session.findById("wnd[0]").sendVKey(0)

        # Destino
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").Select()
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-NAME1").Text = "EMPR"
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-LGOBE").Text = deposito
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-WEMPF").Text = f"EMPR 1{deposito}"
        session.findById("wnd[0]").sendVKey(0)

        # Classificação Contábil
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT").Select()
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-AUFNR").Text = cur_ord
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-VORNR_AUF").Text = "0010"
        session.findById("wnd[0]").sendVKey(0)

        # Seriais
        if is_serial and key in grupos_seriais and grupos_seriais[key]:
            texto_clip = "\r\n".join(grupos_seriais[key])
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(texto_clip)
                win32clipboard.CloseClipboard()
            except Exception:
                pass

            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").Select()

            btn_ser = None
            try:
                btn_ser = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_SEARCH", False)
            except Exception:
                pass
            if not btn_ser:
                btn_ser = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_SER_SEARCH")
            btn_ser.press()

            session.findById("wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[24]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            grid = None
            try:
                grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
            except Exception:
                pass
            if not grid:
                grid = session.findById("wnd[0]/usr/cntlCO_ITEM/shellcont/shell")
            grid.selectAll()
            session.findById("wnd[0]/tbar[1]/btn[47]").press()

        try:
            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").Selected = True
        except Exception:
            pass

        if k < (total_grupos - 1):
            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM").press()

    session.findById("wnd[0]/tbar[1]/btn[7]").press()
    if session.findById("wnd[1]", False):
        lbl_msg = session.findById("wnd[1]/usr/lbl[10,3]", False)
        raise RuntimeError(f"Transação não concluída no SAP: {lbl_msg.Text if lbl_msg else 'Verifique no SAP'}")

    session.findById("wnd[0]/tbar[1]/btn[23]").press()

    if session.findById("wnd[1]", False):
        btn_close = session.findById("wnd[1]/usr/btnBUTTON_2", False)
        if btn_close:
            btn_close.press()

    status_text = session.findById("wnd[0]/sbar/pane[0]").Text
    partes_doc = status_text.split()
    doc_sap = partes_doc[1] if len(partes_doc) > 1 else status_text
    return doc_sap


# ==============================================================================
# MOTOR DA PLANILHA EXCEL (Leitura e Escrita Segura)
# ==============================================================================

class DadosPlanilhaControle:
    def __init__(self, caminho):
        self.caminho = caminho
        self.depositos = []
        self.medidores = []
        self.equipes = []
        self.equipes_dados = {}     # equipe -> {deposito, segmento, ordem, vencimento}
        self.materiais = []
        self.materiais_dados = {}   # material -> {descricao, und, cl03_qtd, cl04_qtd}
        self.limites_segmento = {}  # segmento -> {material: limite}
        self.estoque_seriais = set()
        self.estoque_map = {}       # serial -> material
        self.estoque_lotes = []
        self.fornecidos_seriais = set()
        self.fornecidos_map = {}    # serial -> material
        self.fornecidos_equipes = {} # equipe -> {material: count}

    def carregar(self):
        """Carrega dados em modo somente leitura para alta performance."""
        if not os.path.exists(self.caminho):
            raise FileNotFoundError(f"Planilha não encontrada: {self.caminho}")

        wb = openpyxl.load_workbook(self.caminho, read_only=True, data_only=True)

        if "CADASTRO" in wb.sheetnames:
            ws = wb["CADASTRO"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                dep = str(row[0]).strip() if row[0] is not None else ""
                if dep and dep not in self.depositos:
                    self.depositos.append(dep)

                med = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                if med and med not in self.medidores:
                    self.medidores.append(med)

                if len(row) > 4 and row[4] is not None:
                    eqp = str(row[4]).strip()
                    if eqp:
                        if eqp not in self.equipes:
                            self.equipes.append(eqp)
                        self.equipes_dados[eqp] = {
                            "deposito": str(row[5]).strip() if len(row) > 5 and row[5] else "",
                            "segmento": str(row[6]).strip() if len(row) > 6 and row[6] else "",
                            "ordem": str(row[7]).strip() if len(row) > 7 and row[7] else "",
                            "vencimento": str(row[8]).strip() if len(row) > 8 and row[8] else ""
                        }

                if len(row) > 10 and row[10] is not None:
                    mat = str(row[10]).strip()
                    if mat:
                        if mat not in self.materiais:
                            self.materiais.append(mat)
                        self.materiais_dados[mat] = {
                            "descricao": str(row[11]).strip() if len(row) > 11 and row[11] else "",
                            "und": str(row[12]).strip() if len(row) > 12 and row[12] else "",
                            "cl03_qtd": row[13] if len(row) > 13 and row[13] is not None else 0,
                            "cl04_qtd": row[14] if len(row) > 14 and row[14] is not None else 0
                        }

                        for seg, col_idx in COLUNAS_LIMITES.items():
                            if seg not in self.limites_segmento:
                                self.limites_segmento[seg] = {}
                            lim_val = row[col_idx - 1] if len(row) >= col_idx and row[col_idx - 1] is not None else 0
                            try:
                                self.limites_segmento[seg][mat] = float(lim_val)
                            except Exception:
                                self.limites_segmento[seg][mat] = 0.0

        if "ESTOQUE" in wb.sheetnames:
            ws = wb["ESTOQUE"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None:
                    ser = str(row[0]).strip()
                    mat = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                    dep = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                    lot = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    if ser:
                        self.estoque_seriais.add(ser)
                        self.estoque_map[ser] = mat
                        if lot:
                            self.estoque_lotes.append(lot)

        if "FORNECIDOS" in wb.sheetnames:
            ws = wb["FORNECIDOS"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] is not None:
                    ser = str(row[0]).strip()
                    mat = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                    eqp = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                    if ser:
                        self.fornecidos_seriais.add(ser)
                        self.fornecidos_map[ser] = mat
                        if eqp:
                            if eqp not in self.fornecidos_equipes:
                                self.fornecidos_equipes[eqp] = {}
                            self.fornecidos_equipes[eqp][mat] = self.fornecidos_equipes[eqp].get(mat, 0) + 1

        wb.close()

    def calcular_proximo_lote(self, deposito):
        """Calcula o próximo lote sequencial para o depósito (ex: CL03-045)."""
        dep = str(deposito).strip().upper()
        if not dep:
            return ""
        max_num = 0
        prefixo = f"{dep}-"
        for lote in self.estoque_lotes:
            lote_str = str(lote).strip().upper()
            if lote_str.startswith(prefixo):
                try:
                    num = int(lote_str[len(prefixo):])
                    if num > max_num:
                        max_num = num
                except Exception:
                    pass
        return f"{dep}-{max_num + 1:03d}"


def realizar_backup_planilha(caminho):
    """Cria cópia preventiva no subdiretório _BKP com timestamp."""
    try:
        pasta_dir = os.path.dirname(caminho)
        pasta_bkp = os.path.join(pasta_dir, "_BKP")
        os.makedirs(pasta_bkp, exist_ok=True)
        nome_base = os.path.basename(caminho)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_bkp = f"{os.path.splitext(nome_base)[0]}_BKP_{timestamp}.xlsm"
        caminho_bkp = os.path.join(pasta_bkp, nome_bkp)
        shutil.copy2(caminho, caminho_bkp)
        log_info(f"Backup preventivo criado: {caminho_bkp}")
        return caminho_bkp
    except Exception as e:
        log_error(f"Aviso: Não foi possível criar backup preventivo: {e}")
        return None


# ==============================================================================
# COMPONENTE DE AUTOCOMPLETE ELEGANTE (Janela flutuante compacta)
# ==============================================================================

class AutocompletePopup:
    """Gerencia um popup compacto e moderno de sugestões logo abaixo do campo."""
    def __init__(self, root, entry_widget, obter_paleta, on_select_callback):
        self.root = root
        self.entry = entry_widget
        self.obter_paleta = obter_paleta
        self.on_select = on_select_callback
        self.janela = None
        self.lista = None
        self.itens_disponiveis = []

        self.entry.bind("<KeyRelease>", self.ao_digitar)
        self.entry.bind("<FocusOut>", self.ao_perder_foco)
        self.entry.bind("<Down>", self.ao_teclar_abaixo)

    def atualizar_itens(self, itens):
        self.itens_disponiveis = sorted(list(itens))

    def fechar(self):
        if self.janela and self.janela.winfo_exists():
            self.janela.destroy()
        self.janela = None
        self.lista = None

    def ao_perder_foco(self, event=None):
        self.root.after(150, self._verificar_fechar)

    def _verificar_fechar(self):
        if not self.janela:
            return
        foco = self.root.focus_get()
        if foco != self.entry and foco != self.lista:
            self.fechar()

    def ao_teclar_abaixo(self, event):
        if self.lista:
            self.lista.focus_set()
            self.lista.selection_set(0)

    def ao_digitar(self, event):
        if event.keysym in ["Up", "Return", "Escape", "Tab"]:
            return

        texto = self.entry.get().strip().upper()
        if not texto:
            self.fechar()
            return

        # Filtra as opções que contenham o texto digitado (máximo 8 itens para não poluir!)
        matches = [item for item in self.itens_disponiveis if texto in item.upper()][:8]

        if not matches:
            self.fechar()
            return

        c = self.obter_paleta()
        if not self.janela or not self.janela.winfo_exists():
            self.janela = tk.Toplevel(self.root)
            self.janela.wm_overrideredirect(True)
            self.janela.wm_attributes("-topmost", True)

            self.lista = tk.Listbox(
                self.janela, bg=c["input"], fg=c["texto"],
                selectbackground=c["accent"], selectforeground="white",
                bd=0, highlightthickness=1, highlightbackground=c["accent"],
                font=("Segoe UI", 11), activestyle="none"
            )
            self.lista.pack(fill="both", expand=True)
            self.lista.bind("<ButtonRelease-1>", self._confirmar_selecao)
            self.lista.bind("<Return>", self._confirmar_selecao)
            self.lista.bind("<FocusOut>", self.ao_perder_foco)

        self.lista.delete(0, tk.END)
        for m in matches:
            self.lista.insert(tk.END, m)

        # Posiciona de forma compacta diretamente abaixo do campo
        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        w = max(self.entry.winfo_width(), 160)
        h = len(matches) * 25 + 4
        self.janela.geometry(f"{w}x{h}+{x}+{y}")

    def _confirmar_selecao(self, event=None):
        if self.lista and self.lista.curselection():
            sel = self.lista.get(self.lista.curselection())
            self.entry.delete(0, tk.END)
            self.entry.insert(0, sel)
            self.fechar()
            self.on_select(sel)


# ==============================================================================
# CLASSE PRINCIPAL DA ABA CONTROLE DE MATERIAL
# ==============================================================================

class AbaControleMaterial:
    def __init__(self, parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
        self.parent_frame = parent_frame
        self.label_status_global = label_status_global
        self.paletas = paletas or {"Conecta Blue": {"bg": "#0a192f", "container": "#172a45", "texto": "#ffffff", "accent": "#f57c00", "hover": "#e65100", "input": "#1f3a60"}}
        self.obter_tema_atual = obter_tema_atual or (lambda: "Conecta Blue")

        self.caminho_planilha = obter_caminho_planilha_controle()
        self.dados = None
        self._carregando_dados = False
        self._verificacao_concluida = False

        self._construir_interface()
        self.aplicar_tema(self.obter_tema_atual())
        self.recarregar_dados_async()

    def _cores(self, tema_nome):
        return self.paletas.get(tema_nome, list(self.paletas.values())[0])

    def aplicar_tema(self, tema_nome):
        c = self._cores(tema_nome)
        self.parent_frame.configure(fg_color=c.get("bg", "transparent"))

        self.frame_top.configure(fg_color=c["container"])
        self.painel_esquerdo.configure(fg_color=c["container"])
        self.painel_direito.configure(fg_color=c["container"])
        self.frame_botoes.configure(fg_color="transparent")

        # Labels
        for lbl in [self.lbl_titulo, self.lbl_tipo_mov, self.lbl_cons_titulo, self.lbl_cons_eqp,
                    self.lbl_cons_seg, self.lbl_txt_ser, self.lbl_txt_sai, self.lbl_txt_res]:
            lbl.configure(text_color=c["texto"])

        # Textboxes
        for tb in [self.txt_serie, self.txt_saida, self.txt_resumo, self.txt_nser_list, self.txt_resultado_consulta]:
            tb.configure(fg_color=c["input"], text_color=c["texto"])

        # Buttons
        self.btn_verificar.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_processar.configure(fg_color="#10b981", hover_color="#059669", text_color="white")
        self.btn_limpar.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_recarregar.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])
        self.btn_consultar_saldo.configure(fg_color=c["accent"], hover_color=c["hover"], text_color="white")
        self.btn_verificar_limites.configure(fg_color=c["input"], hover_color=c["hover"], text_color=c["texto"])

    def _construir_interface(self):
        # 1. TOP BAR
        self.frame_top = ctk.CTkFrame(self.parent_frame, corner_radius=8, height=45)
        self.frame_top.pack(fill="x", padx=12, pady=(8, 6))

        self.lbl_titulo = ctk.CTkLabel(self.frame_top, text="📦 Controle de Material (fmain)", font=("Segoe UI", 15, "bold"))
        self.lbl_titulo.pack(side="left", padx=12)

        self.lbl_status_planilha = ctk.CTkLabel(self.frame_top, text="Carregando planilha...", font=("Segoe UI", 11))
        self.lbl_status_planilha.pack(side="left", padx=8)

        self.btn_recarregar = ctk.CTkButton(self.frame_top, text="🔄 Recarregar Dados", width=140, height=28, font=("Segoe UI", 11, "bold"), command=self.recarregar_dados_async)
        self.btn_recarregar.pack(side="right", padx=10)

        self.btn_alterar_planilha = ctk.CTkButton(self.frame_top, text="📁 Localizar Planilha", width=130, height=28, font=("Segoe UI", 11), command=self._selecionar_arquivo_planilha)
        self.btn_alterar_planilha.pack(side="right", padx=4)

        # 2. CONTAINER CORPO
        self.container_corpo = ctk.CTkFrame(self.parent_frame, fg_color="transparent")
        self.container_corpo.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # PAINEL ESQUERDO — CTkFrame Limpo (Sem scrollbar feia externa!)
        self.painel_esquerdo = ctk.CTkFrame(self.container_corpo, corner_radius=8)
        self.painel_esquerdo.pack(side="left", fill="both", expand=True, padx=(0, 6))

        # Barra do Tipo de Movimento (CTkOptionMenu elegante!)
        f_tipo = ctk.CTkFrame(self.painel_esquerdo, fg_color="transparent")
        f_tipo.pack(fill="x", padx=12, pady=(10, 6))

        self.lbl_tipo_mov = ctk.CTkLabel(f_tipo, text="Selecione o Movimento:", font=("Segoe UI", 13, "bold"))
        self.lbl_tipo_mov.pack(side="left", padx=(0, 10))

        self.cbo_tipo_mov = ctk.CTkOptionMenu(f_tipo, values=TIPOS_MOVIMENTO, width=260, font=("Segoe UI", 12, "bold"), command=self._ao_mudar_tipo_movimento)
        self.cbo_tipo_mov.pack(side="left")
        self.cbo_tipo_mov.set(TIPOS_MOVIMENTO[0])  # Padrão: Saída Para OSME

        # CONTAINER DINÂMICO DE CAMPOS (Aqui os campos mudam com clareza)
        self.frame_dinamico = ctk.CTkFrame(self.painel_esquerdo, fg_color="transparent")
        self.frame_dinamico.pack(fill="x", padx=12, pady=2)

        self._criar_widgets_formulario()

        # Seção Não Serializados (expansível)
        self.frame_nserial = ctk.CTkFrame(self.painel_esquerdo, corner_radius=6)
        
        f_ns_inputs = ctk.CTkFrame(self.frame_nserial, fg_color="transparent")
        f_ns_inputs.pack(fill="x", padx=8, pady=4)

        lbl_ns_mat = ctk.CTkLabel(f_ns_inputs, text="Material Não-Serial:", font=("Segoe UI", 11))
        lbl_ns_mat.grid(row=0, column=0, sticky="w", padx=2)
        self.txt_mat_nser = ctk.CTkEntry(f_ns_inputs, width=160, font=("Segoe UI", 11), placeholder_text="Código...")
        self.txt_mat_nser.grid(row=1, column=0, padx=(0, 6), pady=2)

        lbl_ns_qtd = ctk.CTkLabel(f_ns_inputs, text="Qtd:", font=("Segoe UI", 11))
        lbl_ns_qtd.grid(row=0, column=1, sticky="w", padx=2)
        self.txt_nser_qtd = ctk.CTkEntry(f_ns_inputs, width=60, font=("Segoe UI", 11))
        self.txt_nser_qtd.grid(row=1, column=1, padx=(0, 6), pady=2)

        self.lbl_nser_desc = ctk.CTkLabel(f_ns_inputs, text="Descrição: -", font=("Segoe UI", 10), text_color="#94a3b8")
        self.lbl_nser_desc.grid(row=1, column=2, sticky="w", padx=(0, 6), pady=2)

        self.btn_add_nser = ctk.CTkButton(f_ns_inputs, text="➕ Adicionar", width=90, height=26, font=("Segoe UI", 10, "bold"), command=self._adicionar_nserial_item)
        self.btn_add_nser.grid(row=1, column=3, padx=4, pady=2)

        self.txt_nser_list = ctk.CTkTextbox(self.frame_nserial, height=55, font=("Consolas", 10))
        self.txt_nser_list.pack(fill="x", padx=8, pady=(0, 6))

        # Seção de Seriais (Entrada, Conferência, Resumo)
        f_textos = ctk.CTkFrame(self.painel_esquerdo, fg_color="transparent")
        f_textos.pack(fill="both", expand=True, padx=12, pady=6)

        # Coluna Seriais Entrada
        f_c1 = ctk.CTkFrame(f_textos, fg_color="transparent")
        f_c1.pack(side="left", fill="both", expand=True, padx=(0, 4))
        self.lbl_txt_ser = ctk.CTkLabel(f_c1, text="Números de Série (Entrada):", font=("Segoe UI", 11, "bold"))
        self.lbl_txt_ser.pack(anchor="w")
        self.txt_serie = ctk.CTkTextbox(f_c1, height=130, font=("Consolas", 11))
        self.txt_serie.pack(fill="both", expand=True)

        # Coluna Saída (Serial + Material)
        f_c2 = ctk.CTkFrame(f_textos, fg_color="transparent")
        f_c2.pack(side="left", fill="both", expand=True, padx=2)
        self.lbl_txt_sai = ctk.CTkLabel(f_c2, text="Conferência (Serial + Mat):", font=("Segoe UI", 11, "bold"))
        self.lbl_txt_sai.pack(anchor="w")
        self.txt_saida = ctk.CTkTextbox(f_c2, height=130, font=("Consolas", 11))
        self.txt_saida.pack(fill="both", expand=True)

        # Coluna Resumo (Material + Qtd)
        f_c3 = ctk.CTkFrame(f_textos, fg_color="transparent")
        f_c3.pack(side="left", fill="both", expand=True, padx=(4, 0))
        self.lbl_txt_res = ctk.CTkLabel(f_c3, text="Resumo (Material + Qtd):", font=("Segoe UI", 11, "bold"))
        self.lbl_txt_res.pack(anchor="w")
        self.txt_resumo = ctk.CTkTextbox(f_c3, height=130, font=("Consolas", 11))
        self.txt_resumo.pack(fill="both", expand=True)

        # Barra de Ação (Botoes de Comando)
        self.frame_botoes = ctk.CTkFrame(self.painel_esquerdo, height=45)
        self.frame_botoes.pack(fill="x", padx=12, pady=(4, 10))

        self.btn_limpar = ctk.CTkButton(self.frame_botoes, text="🗑️ Limpar Tudo", width=120, height=32, font=("Segoe UI", 11), command=self.limpar_campos)
        self.btn_limpar.pack(side="left", padx=4)

        self.btn_verificar = ctk.CTkButton(self.frame_botoes, text="🔍 Verificar Dados", width=140, height=32, font=("Segoe UI", 11, "bold"), command=self.verificar_movimentacao_async)
        self.btn_verificar.pack(side="left", padx=6)

        self.btn_processar = ctk.CTkButton(self.frame_botoes, text="🚀 Processar Movimentação", width=180, height=32, font=("Segoe UI", 11, "bold"), state="disabled", command=self.processar_movimentacao_async)
        self.btn_processar.pack(side="left", padx=6)

        self.lbl_status_msg = ctk.CTkLabel(self.frame_botoes, text="Pronto para conferência.", font=("Segoe UI", 10), text_color="#94a3b8")
        self.lbl_status_msg.pack(side="left", padx=10)


        # PAINEL DIREITO (Consulta de Saldo & Limites)
        self.painel_direito = ctk.CTkFrame(self.container_corpo, corner_radius=8, width=320)
        self.painel_direito.pack(side="right", fill="both", padx=(6, 0))

        self.lbl_cons_titulo = ctk.CTkLabel(self.painel_direito, text="📊 Consulta de Saldo por Equipe", font=("Segoe UI", 13, "bold"))
        self.lbl_cons_titulo.pack(anchor="w", padx=12, pady=(10, 4))

        f_cons_top = ctk.CTkFrame(self.painel_direito, fg_color="transparent")
        f_cons_top.pack(fill="x", padx=12, pady=4)

        self.lbl_cons_eqp = ctk.CTkLabel(f_cons_top, text="Digite a Equipe:", font=("Segoe UI", 11))
        self.lbl_cons_eqp.pack(anchor="w")

        self.txt_cons_eqp = ctk.CTkEntry(f_cons_top, width=260, font=("Segoe UI", 11), placeholder_text="Ex: CDC328 ou M01...")
        self.txt_cons_eqp.pack(fill="x", pady=2)

        # Autocomplete para consulta lateral
        self.auto_cons_eqp = AutocompletePopup(
            self.parent_frame.winfo_toplevel(), self.txt_cons_eqp,
            lambda: self._cores(self.obter_tema_atual()), self._ao_selecionar_equipe_consulta
        )
        self.txt_cons_eqp.bind("<KeyRelease>", lambda e: self._ao_selecionar_equipe_consulta(self.txt_cons_eqp.get().strip()))

        self.lbl_cons_seg = ctk.CTkLabel(f_cons_top, text="Segmento: -", font=("Segoe UI", 11, "bold"), text_color="#f57c00")
        self.lbl_cons_seg.pack(anchor="w", pady=(4, 6))

        f_botoes_cons = ctk.CTkFrame(f_cons_top, fg_color="transparent")
        f_botoes_cons.pack(fill="x", pady=2)

        self.btn_consultar_saldo = ctk.CTkButton(f_botoes_cons, text="🔎 Consultar Saldo", width=120, height=28, font=("Segoe UI", 11, "bold"), command=self.consultar_saldo_equipe)
        self.btn_consultar_saldo.pack(side="left", padx=(0, 4))

        self.btn_verificar_limites = ctk.CTkButton(f_botoes_cons, text="📋 Ver Limites", width=110, height=28, font=("Segoe UI", 11), command=self.verificar_limites_segmento)
        self.btn_verificar_limites.pack(side="left")

        self.txt_resultado_consulta = ctk.CTkTextbox(self.painel_direito, font=("Consolas", 10))
        self.txt_resultado_consulta.pack(fill="both", expand=True, padx=12, pady=(6, 12))

    def _criar_widgets_formulario(self):
        """Instancia os campos individuais com suporte a autocomplete elegante."""
        # Depósito (OptionMenu com tamanho fixo)
        self.lbl_dep = ctk.CTkLabel(self.frame_dinamico, text="Depósito:", font=("Segoe UI", 11))
        self.cbo_deposito = ctk.CTkOptionMenu(self.frame_dinamico, values=["CL03", "CL04"], width=95, font=("Segoe UI", 11, "bold"), command=self._ao_mudar_deposito)
        self.cbo_deposito.set("CL03")

        # Depósito Destino
        self.lbl_dep_dest = ctk.CTkLabel(self.frame_dinamico, text="Dep. Destino:", font=("Segoe UI", 11))
        self.cbo_dep_dest = ctk.CTkOptionMenu(self.frame_dinamico, values=["CL03", "CL04"], width=95, font=("Segoe UI", 11, "bold"))
        self.cbo_dep_dest.set("CL04")

        # Equipe Origem (Entry com Autocomplete elegante!)
        self.lbl_eqp = ctk.CTkLabel(self.frame_dinamico, text="Equipe:", font=("Segoe UI", 11, "bold"))
        self.txt_equipe = ctk.CTkEntry(self.frame_dinamico, width=150, font=("Segoe UI", 11), placeholder_text="Digite a equipe...")
        self.auto_equipe = AutocompletePopup(
            self.parent_frame.winfo_toplevel(), self.txt_equipe,
            lambda: self._cores(self.obter_tema_atual()), self._preencher_dados_equipe
        )
        self.txt_equipe.bind("<KeyRelease>", self._ao_digitar_equipe)
        self.txt_equipe.bind("<FocusOut>", self._ao_digitar_equipe)

        # OSME / Ordem Origem
        self.lbl_ordem = ctk.CTkLabel(self.frame_dinamico, text="OSME / Ordem:", font=("Segoe UI", 11, "bold"))
        self.txt_ordem = ctk.CTkEntry(self.frame_dinamico, width=125, font=("Segoe UI", 11))
        self.lbl_venc_aviso = ctk.CTkLabel(self.frame_dinamico, text="", font=("Segoe UI", 10))

        # Equipe Destino (com Autocomplete)
        self.lbl_eqp_dest = ctk.CTkLabel(self.frame_dinamico, text="Equipe Destino:", font=("Segoe UI", 11, "bold"))
        self.txt_equipe_dest = ctk.CTkEntry(self.frame_dinamico, width=150, font=("Segoe UI", 11), placeholder_text="Equipe destino...")
        self.auto_equipe_dest = AutocompletePopup(
            self.parent_frame.winfo_toplevel(), self.txt_equipe_dest,
            lambda: self._cores(self.obter_tema_atual()), self._ao_mudar_equipe_dest
        )
        self.txt_equipe_dest.bind("<KeyRelease>", self._ao_digitar_equipe_dest)
        self.txt_equipe_dest.bind("<FocusOut>", self._ao_digitar_equipe_dest)

        # Ordem Destino
        self.lbl_ordem_dest = ctk.CTkLabel(self.frame_dinamico, text="OSME Destino:", font=("Segoe UI", 11, "bold"))
        self.txt_ordem_dest = ctk.CTkEntry(self.frame_dinamico, width=125, font=("Segoe UI", 11))

        # Data Lançamento
        self.lbl_data = ctk.CTkLabel(self.frame_dinamico, text="Data:", font=("Segoe UI", 11))
        self.txt_data = ctk.CTkEntry(self.frame_dinamico, width=100, font=("Segoe UI", 11))
        self.txt_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

        # Lote (Exibido SOMENTE na Entrada de Materiais!)
        self.lbl_lote = ctk.CTkLabel(self.frame_dinamico, text="Lote Gerado:", font=("Segoe UI", 11, "bold"))
        self.txt_lote = ctk.CTkEntry(self.frame_dinamico, width=110, font=("Segoe UI", 11, "bold"))

        # Material (SOMENTE na Entrada de Materiais!)
        self.lbl_mat = ctk.CTkLabel(self.frame_dinamico, text="Código do Material:", font=("Segoe UI", 11, "bold"))
        self.txt_material = ctk.CTkEntry(self.frame_dinamico, width=160, font=("Segoe UI", 11), placeholder_text="Digite o material...")
        self.auto_material = AutocompletePopup(
            self.parent_frame.winfo_toplevel(), self.txt_material,
            lambda: self._cores(self.obter_tema_atual()), lambda m: None
        )

        # Checkboxes
        self.chk_nserial = ctk.CTkCheckBox(self.frame_dinamico, text="Não Serializados", font=("Segoe UI", 11), command=self._ao_alternar_nserial)
        self.chk_sai_ordem = ctk.CTkCheckBox(self.frame_dinamico, text="Saída via Ordem", font=("Segoe UI", 11))
        self.chk_ci = ctk.CTkCheckBox(self.frame_dinamico, text="Gerar CI (Word)", font=("Segoe UI", 11), command=self._ao_alternar_ci)

        # Campos CI
        self.lbl_pt = ctk.CTkLabel(self.frame_dinamico, text="Pedido Transf. (PT):", font=("Segoe UI", 11))
        self.txt_pt = ctk.CTkEntry(self.frame_dinamico, width=120, font=("Segoe UI", 11))
        self.lbl_docsai = ctk.CTkLabel(self.frame_dinamico, text="Doc. de Saída:", font=("Segoe UI", 11))
        self.txt_docsai = ctk.CTkEntry(self.frame_dinamico, width=120, font=("Segoe UI", 11))


    # ==========================================================================
    # GERENCIADOR DINÂMICO DE VISIBILIDADE (Sem poluição visual)
    # ==========================================================================

    def _ao_mudar_tipo_movimento(self, escolha):
        self._verificacao_concluida = False
        self.btn_processar.configure(state="disabled")

        # Fecha popups de autocomplete abertos
        self.auto_equipe.fechar()
        self.auto_equipe_dest.fechar()
        self.auto_material.fechar()

        for widget in self.frame_dinamico.winfo_children():
            widget.grid_forget()

        self.frame_nserial.pack_forget()
        self.chk_nserial.deselect()

        # 1. SAÍDA PARA OSME (Padrão, Limpo e Focado!)
        if escolha == "Saída Para OSME":
            self.lbl_eqp.grid(row=0, column=0, sticky="w", padx=4, pady=2)
            self.txt_equipe.grid(row=1, column=0, padx=(0, 10), pady=2)

            self.lbl_ordem.grid(row=0, column=1, sticky="w", padx=4, pady=2)
            self.txt_ordem.grid(row=1, column=1, padx=(0, 10), pady=2)

            self.lbl_dep.grid(row=0, column=2, sticky="w", padx=4, pady=2)
            self.cbo_deposito.grid(row=1, column=2, padx=(0, 10), pady=2)

            self.lbl_data.grid(row=0, column=3, sticky="w", padx=4, pady=2)
            self.txt_data.grid(row=1, column=3, padx=(0, 10), pady=2)

            self.chk_sai_ordem.grid(row=1, column=4, padx=(0, 10), pady=2)
            self.chk_nserial.grid(row=1, column=5, padx=4, pady=2)

            self.lbl_venc_aviso.grid(row=2, column=1, sticky="w", padx=4, pady=1)
            self.lbl_status_msg.configure(text="Digite a equipe. A OSME e o Depósito serão preenchidos automaticamente.", text_color="#94a3b8")

        # 2. ENTRADA DE MATERIAIS (Lote aparece apenas aqui!)
        elif escolha == "Entrada de Materiais":
            self.lbl_dep.grid(row=0, column=0, sticky="w", padx=4, pady=2)
            self.cbo_deposito.grid(row=1, column=0, padx=(0, 10), pady=2)

            self.lbl_mat.grid(row=0, column=1, sticky="w", padx=4, pady=2)
            self.txt_material.grid(row=1, column=1, padx=(0, 10), pady=2)

            self.lbl_lote.grid(row=0, column=2, sticky="w", padx=4, pady=2)
            self.txt_lote.grid(row=1, column=2, padx=(0, 10), pady=2)

            self.lbl_data.grid(row=0, column=3, sticky="w", padx=4, pady=2)
            self.txt_data.grid(row=1, column=3, padx=4, pady=2)

            self._atualizar_proximo_lote()
            self.lbl_status_msg.configure(text="Selecione o depósito e o material. O próximo lote sequencial foi gerado automaticamente.", text_color="#94a3b8")

        # 3. DEVOLUÇÃO DE MATERIAIS
        elif escolha == "Devolução de Materiais":
            self.lbl_dep.grid(row=0, column=0, sticky="w", padx=4, pady=2)
            self.cbo_deposito.grid(row=1, column=0, padx=(0, 10), pady=2)

            self.lbl_data.grid(row=0, column=1, sticky="w", padx=4, pady=2)
            self.txt_data.grid(row=1, column=1, padx=(0, 10), pady=2)

            self.chk_nserial.grid(row=1, column=2, padx=10, pady=2)
            self.lbl_status_msg.configure(text="Insira os seriais devolvidos. A equipe e os materiais serão identificados automaticamente.", text_color="#94a3b8")

        # 4. TRANSFERÊNCIA ENTRE EQUIPES
        elif escolha == "Transferência Entre Equipes":
            self.lbl_eqp_dest.grid(row=0, column=0, sticky="w", padx=4, pady=2)
            self.txt_equipe_dest.grid(row=1, column=0, padx=(0, 10), pady=2)

            self.lbl_ordem_dest.grid(row=0, column=1, sticky="w", padx=4, pady=2)
            self.txt_ordem_dest.grid(row=1, column=1, padx=(0, 10), pady=2)

            self.lbl_dep_dest.grid(row=0, column=2, sticky="w", padx=4, pady=2)
            self.cbo_dep_dest.grid(row=1, column=2, padx=(0, 10), pady=2)

            self.lbl_data.grid(row=0, column=3, sticky="w", padx=4, pady=2)
            self.txt_data.grid(row=1, column=3, padx=(0, 10), pady=2)

            self.chk_nserial.grid(row=1, column=4, padx=10, pady=2)
            self.lbl_status_msg.configure(text="Digite a Equipe Destino para transferência.", text_color="#94a3b8")

        # 5. TRANSFERÊNCIA ENTRE DEPÓSITOS
        elif escolha == "Transferência Entre Depósitos":
            self.lbl_dep.grid(row=0, column=0, sticky="w", padx=4, pady=2)
            self.cbo_deposito.grid(row=1, column=0, padx=(0, 10), pady=2)

            self.lbl_dep_dest.grid(row=0, column=1, sticky="w", padx=4, pady=2)
            self.cbo_dep_dest.grid(row=1, column=1, padx=(0, 10), pady=2)

            self.lbl_data.grid(row=0, column=2, sticky="w", padx=4, pady=2)
            self.txt_data.grid(row=1, column=2, padx=(0, 10), pady=2)

            self.chk_ci.grid(row=1, column=3, padx=10, pady=2)
            self.lbl_status_msg.configure(text="Transferência física entre depósitos próprios ou externos.", text_color="#94a3b8")

        # 6, 7, 8: ROTINAS SAP
        else:
            lbl_aviso_sap = ctk.CTkLabel(
                self.frame_dinamico,
                text=f"⚡ Rotina SAP: {escolha}\nClique em 'Verificar Dados' para iniciar a execução automática no SAP.",
                font=("Segoe UI", 12, "bold"), text_color="#38bdf8"
            )
            lbl_aviso_sap.grid(row=0, column=0, columnspan=4, sticky="w", padx=4, pady=8)


    # ==========================================================================
    # AUTO-PREENCHIMENTO INSTANTÂNEO DE EQUIPE -> OSME / DEPÓSITO
    # ==========================================================================

    def _ao_digitar_equipe(self, event=None):
        if not self.dados:
            return
        texto = self.txt_equipe.get().strip()
        if not texto:
            self.txt_ordem.delete(0, tk.END)
            self.lbl_venc_aviso.configure(text="")
            return

        texto_upper = texto.upper()
        equipe_encontrada = None
        if texto in self.dados.equipes_dados:
            equipe_encontrada = texto
        else:
            for eqp in self.dados.equipes_dados:
                if eqp.upper() == texto_upper:
                    equipe_encontrada = eqp
                    break

        if equipe_encontrada:
            self._preencher_dados_equipe(equipe_encontrada)

    def _preencher_dados_equipe(self, equipe):
        info = self.dados.equipes_dados.get(equipe, {})
        dep = info.get("deposito", "")
        ordem = info.get("ordem", "")
        venc = info.get("vencimento", "")

        # 1. Depósito
        if dep:
            self.cbo_deposito.set(dep)
            if dep.upper() == "CL04":
                self.chk_sai_ordem.select()
            else:
                self.chk_sai_ordem.deselect()

        # 2. Ordem / OSME
        self.txt_ordem.delete(0, tk.END)
        self.txt_ordem.insert(0, ordem)

        # 3. Vencimento
        c = self._cores(self.obter_tema_atual())
        if venc:
            try:
                dt_venc = datetime.strptime(venc.split()[0], "%d/%m/%Y")
                if dt_venc.date() < datetime.now().date():
                    self.txt_ordem.configure(text_color="#ef4444")
                    self.lbl_venc_aviso.configure(text=f"⚠️ Ordem vencida em {venc.split()[0]}!", text_color="#ef4444")
                else:
                    self.txt_ordem.configure(text_color=c["texto"])
                    self.lbl_venc_aviso.configure(text=f"Vence em: {venc.split()[0]} (OK)", text_color="#10b981")
            except Exception:
                self.txt_ordem.configure(text_color=c["texto"])
                self.lbl_venc_aviso.configure(text="")
        else:
            self.txt_ordem.configure(text_color=c["texto"])
            self.lbl_venc_aviso.configure(text="")

    def _ao_digitar_equipe_dest(self, event=None):
        if not self.dados:
            return
        texto = self.txt_equipe_dest.get().strip().upper()
        for eqp in self.dados.equipes_dados:
            if eqp.upper() == texto:
                self._ao_mudar_equipe_dest(eqp)
                break

    def _ao_mudar_equipe_dest(self, equipe):
        if not self.dados or equipe not in self.dados.equipes_dados:
            return
        info = self.dados.equipes_dados[equipe]
        dep = info.get("deposito", "")
        ordem = info.get("ordem", "")
        if dep:
            self.cbo_dep_dest.set(dep)
        self.txt_ordem_dest.delete(0, tk.END)
        self.txt_ordem_dest.insert(0, ordem)

    def _ao_mudar_deposito(self, val):
        if self.cbo_tipo_mov.get() == "Entrada de Materiais":
            self._atualizar_proximo_lote()

    def _atualizar_proximo_lote(self):
        if self.dados:
            dep = self.cbo_deposito.get()
            proximo = self.dados.calcular_proximo_lote(dep)
            if proximo:
                self.txt_lote.delete(0, tk.END)
                self.txt_lote.insert(0, proximo)

    def _ao_alternar_nserial(self):
        if self.chk_nserial.get():
            self.frame_nserial.pack(fill="x", padx=12, pady=4)
        else:
            self.frame_nserial.pack_forget()

    def _ao_alternar_ci(self):
        if self.chk_ci.get():
            self.lbl_pt.grid(row=2, column=0, sticky="w", padx=4, pady=2)
            self.txt_pt.grid(row=3, column=0, padx=(0, 10), pady=2)
            self.lbl_docsai.grid(row=2, column=1, sticky="w", padx=4, pady=2)
            self.txt_docsai.grid(row=3, column=1, padx=(0, 10), pady=2)
        else:
            self.lbl_pt.grid_forget()
            self.txt_pt.grid_forget()
            self.lbl_docsai.grid_forget()
            self.txt_docsai.grid_forget()

    def _ao_mudar_mat_nser(self, val):
        if self.dados and val in self.dados.materiais_dados:
            info = self.dados.materiais_dados[val]
            self.lbl_nser_desc.configure(text=f"UND: {info.get('und', '')} | {info.get('descricao', '')}")

    def _adicionar_nserial_item(self):
        mat = self.txt_mat_nser.get().strip()
        qtd_txt = self.txt_nser_qtd.get().strip()
        if not mat or not qtd_txt:
            messagebox.showwarning("Atenção", "Preencha o material não-serializado e a quantidade.")
            return
        try:
            qtd = int(qtd_txt)
            if qtd <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Erro", "Quantidade inválida.")
            return

        desc = self.dados.materiais_dados.get(mat, {}).get("descricao", "") if self.dados else ""
        linha = f"{qtd}\t{mat}\t{desc}\n"
        self.txt_nser_list.insert(tk.END, linha)
        self.txt_nser_qtd.delete(0, tk.END)

    def limpar_campos(self):
        self.txt_serie.delete("1.0", tk.END)
        self.txt_saida.delete("1.0", tk.END)
        self.txt_resumo.delete("1.0", tk.END)
        self.txt_nser_list.delete("1.0", tk.END)
        self.txt_nser_qtd.delete(0, tk.END)
        self.txt_equipe.delete(0, tk.END)
        self.txt_ordem.delete(0, tk.END)
        self.lbl_venc_aviso.configure(text="")
        self._verificacao_concluida = False
        self.btn_processar.configure(state="disabled")
        self.lbl_status_msg.configure(text="Campos limpos.", text_color="#94a3b8")


    # ==========================================================================
    # CARREGAMENTO ASSÍNCRONO DA PLANILHA
    # ==========================================================================

    def recarregar_dados_async(self):
        if self._carregando_dados:
            return
        self._carregando_dados = True
        self.lbl_status_planilha.configure(text="Lendo planilha Excel...", text_color="#f59e0b")
        self.btn_recarregar.configure(state="disabled")

        def _tarefa():
            try:
                dados = DadosPlanilhaControle(self.caminho_planilha)
                dados.carregar()
                self.dados = dados
                self.parent_frame.after(0, self._ao_concluir_carregamento, True, f"Planilha carregada ({len(dados.equipes)} equipes, {len(dados.estoque_seriais)} em estoque)")
            except Exception as e:
                log_error(f"Erro ao carregar planilha de controle: {e}")
                self.parent_frame.after(0, self._ao_concluir_carregamento, False, str(e))

        threading.Thread(target=_tarefa, daemon=True).start()

    def _ao_concluir_carregamento(self, sucesso, mensagem):
        self._carregando_dados = False
        self.btn_recarregar.configure(state="normal")
        if sucesso:
            self.lbl_status_planilha.configure(text="🟢 " + mensagem, text_color="#10b981")
            
            # Atualiza lista de itens no autocomplete
            self.auto_equipe.atualizar_itens(self.dados.equipes)
            self.auto_equipe_dest.atualizar_itens(self.dados.equipes)
            self.auto_cons_eqp.atualizar_itens(self.dados.equipes)
            self.auto_material.atualizar_itens(self.dados.materiais)

            self._ao_mudar_tipo_movimento(self.cbo_tipo_mov.get())
        else:
            self.lbl_status_planilha.configure(text=f"🔴 Erro: {mensagem}", text_color="#ef4444")
            messagebox.showerror("Erro na Planilha", f"Não foi possível ler a planilha:\n{self.caminho_planilha}\n\nErro: {mensagem}")

    def _selecionar_arquivo_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Localizar Planilha CONTROLE_SERIAIS.vBeta.xlsm",
            filetypes=[("Planilhas Excel Habilitadas para Macro", "*.xlsm"), ("Todos os Arquivos", "*.*")]
        )
        if caminho and os.path.exists(caminho):
            self.caminho_planilha = caminho
            salvar_caminho_planilha_controle(caminho)
            self.recarregar_dados_async()


    # ==========================================================================
    # VERIFICAÇÃO DE DADOS (Validar sem gravar)
    # ==========================================================================

    def verificar_movimentacao_async(self):
        self.lbl_status_msg.configure(text="Verificando dados...", text_color="#f59e0b")
        threading.Thread(target=self._executar_verificacao, daemon=True).start()

    def _executar_verificacao(self):
        if not self.dados:
            self._msg_erro("Planilha ainda não foi carregada. Aguarde ou recarregue.")
            return

        escolha = self.cbo_tipo_mov.get()
        texto_seriais = self.txt_serie.get("1.0", tk.END).strip()
        linhas_seriais = [s.strip() for s in texto_seriais.splitlines() if s.strip()]

        # 1. SAÍDA PARA OSME
        if escolha == "Saída Para OSME":
            dep = self.cbo_deposito.get().strip()
            eqp = self.txt_equipe.get().strip()
            ordem = self.txt_ordem.get().strip()

            if not eqp:
                self._msg_erro("Digite ou selecione a Equipe para a Saída.")
                return

            if not self.chk_nserial.get() and not linhas_seriais:
                self._msg_erro("Insira os números de série ou marque Não Serializados.")
                return

            faltantes = [s for s in linhas_seriais if s not in self.dados.estoque_seriais]
            if faltantes:
                self._msg_erro("Os seguintes seriais NÃO estão em ESTOQUE:\n" + "\n".join(faltantes[:10]))
                return

            dups_forn = [s for s in linhas_seriais if s in self.dados.fornecidos_seriais]
            if dups_forn:
                self._msg_erro("Os seguintes seriais já constam como FORNECIDOS:\n" + "\n".join(dups_forn[:10]))
                return

            saida_linhas = []
            contagem = {}
            for s in linhas_seriais:
                m = self.dados.estoque_map.get(s, "")
                saida_linhas.append(f"{s}\t{m}")
                contagem[m] = contagem.get(m, 0) + 1

            saida = "\n".join(saida_linhas)
            resumo = "\n".join([f"{m}\t{qtd}" for m, qtd in contagem.items()])

            if self.chk_sai_ordem.get() and ordem:
                session = conectar_sap()
                if session:
                    ok, erro = sap_verificar_reserva_ordem(session, ordem, dep, resumo)
                    if not ok:
                        self._msg_erro(f"Aviso SAP sobre reserva na ordem:\n{erro}")
                        return

            self.parent_frame.after(0, self._aplicar_resultado_verificacao, saida, resumo, "Verificação OK! Seriais validados e prontos para MIGO.")

        # 2. ENTRADA DE MATERIAIS
        elif escolha == "Entrada de Materiais":
            mat = self.txt_material.get().strip()
            dep = self.cbo_deposito.get().strip()
            if not mat or not dep:
                self._msg_erro("Selecione o Material e o Depósito para a Entrada.")
                return
            if not linhas_seriais:
                self._msg_erro("Insira ao menos um número de série.")
                return

            dups_est = [s for s in linhas_seriais if s in self.dados.estoque_seriais]
            if dups_est:
                self._msg_erro(f"Os seguintes seriais já estão em ESTOQUE:\n" + "\n".join(dups_est[:10]))
                return
            dups_forn = [s for s in linhas_seriais if s in self.dados.fornecidos_seriais]
            if dups_forn:
                self._msg_erro(f"Os seguintes seriais já constam como FORNECIDOS:\n" + "\n".join(dups_forn[:10]))
                return

            saida = "\n".join([f"{s}\t{mat}" for s in linhas_seriais])
            resumo = f"{mat}\t{len(linhas_seriais)}"
            self.parent_frame.after(0, self._aplicar_resultado_verificacao, saida, resumo, "Verificação OK! Pronto para registrar entrada.")

        # 3. DEVOLUÇÃO DE MATERIAIS
        elif escolha == "Devolução de Materiais":
            dep = self.cbo_deposito.get().strip()
            if not linhas_seriais:
                self._msg_erro("Insira os seriais que estão sendo devolvidos.")
                return

            nao_forn = [s for s in linhas_seriais if s not in self.dados.fornecidos_seriais]
            if nao_forn:
                self._msg_erro("Os seguintes seriais NÃO constam como FORNECIDOS:\n" + "\n".join(nao_forn[:10]))
                return

            em_est = [s for s in linhas_seriais if s in self.dados.estoque_seriais]
            if em_est:
                self._msg_erro("Os seguintes seriais JÁ estão cadastrados no ESTOQUE:\n" + "\n".join(em_est[:10]))
                return

            saida_linhas = []
            contagem = {}
            for s in linhas_seriais:
                m = self.dados.fornecidos_map.get(s, "")
                saida_linhas.append(f"{s}\t{m}")
                contagem[m] = contagem.get(m, 0) + 1

            saida = "\n".join(saida_linhas)
            resumo = "\n".join([f"{m}\t{qtd}" for m, qtd in contagem.items()])
            self.parent_frame.after(0, self._aplicar_resultado_verificacao, saida, resumo, "Verificação OK! Pronto para processar devolução.")

        # 4. TRANSFERÊNCIA ENTRE EQUIPES
        elif escolha == "Transferência Entre Equipes":
            eqp_dest = self.txt_equipe_dest.get().strip()
            if not eqp_dest:
                self._msg_erro("Digite ou selecione a Equipe de Destino.")
                return
            if not linhas_seriais:
                self._msg_erro("Insira os seriais para transferência.")
                return

            nao_forn = [s for s in linhas_seriais if s not in self.dados.fornecidos_seriais]
            if nao_forn:
                self._msg_erro("Os seguintes seriais NÃO constam como fornecidos:\n" + "\n".join(nao_forn[:10]))
                return

            saida_linhas = []
            contagem = {}
            for s in linhas_seriais:
                m = self.dados.fornecidos_map.get(s, "")
                saida_linhas.append(f"{s}\t{m}")
                contagem[m] = contagem.get(m, 0) + 1

            saida = "\n".join(saida_linhas)
            resumo = "\n".join([f"{m}\t{qtd}" for m, qtd in contagem.items()])
            self.parent_frame.after(0, self._aplicar_resultado_verificacao, saida, resumo, "Verificação OK! Pronto para transferir de equipe.")

        else:
            self.parent_frame.after(0, self._aplicar_resultado_verificacao, "", "", "Pronto para executar rotina no SAP.")

    def _aplicar_resultado_verificacao(self, saida, resumo, msg):
        self.txt_saida.delete("1.0", tk.END)
        self.txt_saida.insert("1.0", saida)
        self.txt_resumo.delete("1.0", tk.END)
        self.txt_resumo.insert("1.0", resumo)
        self._verificacao_concluida = True
        self.btn_processar.configure(state="normal")
        self.lbl_status_msg.configure(text=f"✅ {msg}", text_color="#10b981")
        show_toast(self.parent_frame, "Verificação concluída com sucesso!", type="success")

    def _msg_erro(self, texto):
        def _show():
            self.lbl_status_msg.configure(text=f"⚠️ {texto.splitlines()[0]}", text_color="#ef4444")
            messagebox.showwarning("Atenção na Verificação", texto)
        self.parent_frame.after(0, _show)


    # ==========================================================================
    # PROCESSAMENTO E ALIMENTAÇÃO DA PLANILHA (keep_vba=True)
    # ==========================================================================

    def processar_movimentacao_async(self):
        if not self._verificacao_concluida:
            messagebox.showwarning("Aviso", "Execute a verificação antes de processar.")
            return

        resposta = messagebox.askyesno("Confirmar Processamento", "Confirma a gravação e alimentação da planilha?")
        if not resposta:
            return

        self.btn_processar.configure(state="disabled")
        self.lbl_status_msg.configure(text="Processando transação e alimentando planilha...", text_color="#f59e0b")
        threading.Thread(target=self._executar_processamento, daemon=True).start()

    def _executar_processamento(self):
        try:
            escolha = self.cbo_tipo_mov.get()
            dep = self.cbo_deposito.get().strip()
            dep_dest = self.cbo_dep_dest.get().strip()
            eqp = self.txt_equipe.get().strip()
            eqp_dest = self.txt_equipe_dest.get().strip()
            ordem = self.txt_ordem.get().strip()
            ordem_dest = self.txt_ordem_dest.get().strip()
            lote = self.txt_lote.get().strip()
            data_hoje = self.txt_data.get().strip() or datetime.now().strftime("%d/%m/%Y")
            saida_texto = self.txt_saida.get("1.0", tk.END).strip()
            linhas_saida = [l.strip() for l in saida_texto.splitlines() if l.strip()]

            doc_sap = ""
            session = conectar_sap()

            if escolha == "Saída Para OSME":
                if not session:
                    raise RuntimeError("É necessário estar com o SAP GUI aberto para registrar Saída para OSME.")
                self.lbl_status_msg.configure(text="Executando MIGO no SAP...")
                doc_sap = sap_migo_to_osme(session, saida_texto, ordem, dep, data_hoje, is_serial=True)
                log_info(f"Documento SAP gerado na saída: {doc_sap}")

            elif escolha == "Transferência Entre Equipes":
                if session:
                    self.lbl_status_msg.configure(text="Executando transferência no SAP...")
                    doc_sap = sap_migo_to_osme(session, saida_texto, ordem_dest, dep_dest or dep, data_hoje, is_serial=True)

            realizar_backup_planilha(self.caminho_planilha)

            self.lbl_status_msg.configure(text="Alimentando planilha Excel com segurança...")
            wb = openpyxl.load_workbook(self.caminho_planilha, keep_vba=True)

            if escolha == "Entrada de Materiais":
                ws_est = wb["ESTOQUE"]
                for l in linhas_saida:
                    partes = l.split("\t")
                    if len(partes) >= 2:
                        ser, mat = partes[0].strip(), partes[1].strip()
                        ws_est.append([ser, mat, dep, lote, data_hoje, "", "", ""])

            elif escolha == "Saída Para OSME":
                ws_forn = wb["FORNECIDOS"]
                ws_est = wb["ESTOQUE"]
                seriais_set = set()

                for l in linhas_saida:
                    partes = l.split("\t")
                    if len(partes) >= 2:
                        ser, mat = partes[0].strip(), partes[1].strip()
                        seriais_set.add(ser)
                        ws_forn.append([ser, mat, dep, eqp, data_hoje, doc_sap, "", "", ""])

                linhas_deletar = []
                for row_idx, row in enumerate(ws_est.iter_rows(min_row=2, values_only=True), start=2):
                    if row and row[0] and str(row[0]).strip() in seriais_set:
                        linhas_deletar.append(row_idx)

                for r in reversed(linhas_deletar):
                    ws_est.delete_rows(r, 1)

            elif escolha == "Devolução de Materiais":
                ws_forn = wb["FORNECIDOS"]
                ws_est = wb["ESTOQUE"]
                seriais_set = set()

                for l in linhas_saida:
                    partes = l.split("\t")
                    if len(partes) >= 2:
                        ser, mat = partes[0].strip(), partes[1].strip()
                        seriais_set.add(ser)
                        ws_est.append([ser, mat, dep, "LT0", data_hoje, "", "", ""])

                linhas_deletar = []
                for row_idx, row in enumerate(ws_forn.iter_rows(min_row=2, values_only=True), start=2):
                    if row and row[0] and str(row[0]).strip() in seriais_set:
                        linhas_deletar.append(row_idx)

                for r in reversed(linhas_deletar):
                    ws_forn.delete_rows(r, 1)

            elif escolha == "Transferência Entre Equipes":
                ws_forn = wb["FORNECIDOS"]
                seriais_set = set()
                for l in linhas_saida:
                    partes = l.split("\t")
                    if len(partes) >= 2:
                        ser, mat = partes[0].strip(), partes[1].strip()
                        seriais_set.add(ser)
                        ws_forn.append([ser, mat, dep_dest or dep, eqp_dest, data_hoje, doc_sap, "", "", ""])

                linhas_deletar = []
                for row_idx, row in enumerate(ws_forn.iter_rows(min_row=2, values_only=True), start=2):
                    if row and row[0] and str(row[0]).strip() in seriais_set:
                        eqp_row = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                        if eqp_row != eqp_dest:
                            linhas_deletar.append(row_idx)

                for r in reversed(linhas_deletar):
                    ws_forn.delete_rows(r, 1)

            wb.save(self.caminho_planilha)
            wb.close()

            self.parent_frame.after(0, self._ao_sucesso_processamento, doc_sap)

        except Exception as e:
            log_error(f"Erro ao processar movimentação: {e}")
            self.parent_frame.after(0, self._ao_erro_processamento, str(e))

    def _ao_sucesso_processamento(self, doc_sap):
        msg = "Movimentação processada e planilha atualizada com sucesso!"
        if doc_sap:
            msg += f"\nDocumento SAP: {doc_sap}"
        messagebox.showinfo("Sucesso", msg)
        show_toast(self.parent_frame, "Planilha alimentada com sucesso!", type="success")
        play_sound("done")
        self.limpar_campos()
        self.recarregar_dados_async()

    def _ao_erro_processamento(self, erro):
        self.btn_processar.configure(state="normal")
        self.lbl_status_msg.configure(text=f"❌ Erro no processamento: {erro}", text_color="#ef4444")
        messagebox.showerror("Erro no Processamento", f"Ocorreu um erro ao processar:\n{erro}")


    # ==========================================================================
    # PAINEL DE CONSULTA DE SALDO E LIMITES POR EQUIPE
    # ==========================================================================

    def _ao_selecionar_equipe_consulta(self, eqp):
        if not self.dados:
            return
        eqp = eqp.strip()
        if not eqp:
            self.lbl_cons_seg.configure(text="Segmento: -")
            return

        texto_upper = eqp.upper()
        eqp_match = None
        if eqp in self.dados.equipes_dados:
            eqp_match = eqp
        else:
            for e in self.dados.equipes_dados:
                if e.upper() == texto_upper:
                    eqp_match = e
                    break

        if eqp_match:
            seg = self.dados.equipes_dados[eqp_match].get("segmento", "-")
            self.lbl_cons_seg.configure(text=f"Segmento: {seg.upper()}")
        else:
            self.lbl_cons_seg.configure(text="Segmento: -")

    def consultar_saldo_equipe(self):
        eqp = self.txt_cons_eqp.get().strip()
        if not eqp or not self.dados:
            messagebox.showwarning("Aviso", "Digite uma equipe para consultar o saldo.")
            return

        eqp_match = next((e for e in self.dados.equipes_dados if e.upper() == eqp.upper()), eqp)
        seg = self.dados.equipes_dados.get(eqp_match, {}).get("segmento", "").upper()
        if not seg:
            messagebox.showwarning("Aviso", f"Equipe '{eqp}' não encontrada ou sem segmento cadastrado.")
            return

        materiais_equipe = self.dados.fornecidos_equipes.get(eqp_match, {})
        limites = self.dados.limites_segmento.get(seg, {})

        linhas_rel = []
        linhas_rel.append(f"{'MATERIAL':<12} | {'EM PODER':<8} | {'LIMITE':<8} | {'DISPONÍVEL':<10}")
        linhas_rel.append("-" * 48)

        todos_mats = sorted(set(list(materiais_equipe.keys()) + [m for m, l in limites.items() if l > 0]))

        for mat in todos_mats:
            qtd_com_equipe = materiais_equipe.get(mat, 0)
            lim_val = limites.get(mat, 0.0)
            disp = lim_val - qtd_com_equipe if lim_val > 0 else 0
            disp_str = f"{int(disp)}" if disp >= 0 else f"ESTOURADO ({int(disp)})"
            linhas_rel.append(f"{mat:<12} | {qtd_com_equipe:<8} | {int(lim_val):<8} | {disp_str:<10}")

        self.txt_resultado_consulta.delete("1.0", tk.END)
        self.txt_resultado_consulta.insert("1.0", "\n".join(linhas_rel))

    def verificar_limites_segmento(self):
        eqp = self.txt_cons_eqp.get().strip()
        if not self.dados:
            return
        eqp_match = next((e for e in self.dados.equipes_dados if e.upper() == eqp.upper()), eqp) if eqp else ""
        seg = self.dados.equipes_dados.get(eqp_match, {}).get("segmento", "").upper() if eqp_match else "LIGAÇÃO"

        limites = self.dados.limites_segmento.get(seg, {})
        com_limite = [(m, l) for m, l in limites.items() if l > 0]

        linhas = [f"📋 LIMITES CADASTRADOS PARA: {seg}", "-" * 40]
        if not com_limite:
            linhas.append("Nenhum limite configurado para este segmento.")
        else:
            for m, l in com_limite:
                desc = self.dados.materiais_dados.get(m, {}).get("descricao", "")
                linhas.append(f"{m} = {int(l)} un  ({desc[:22]})")

        self.txt_resultado_consulta.delete("1.0", tk.END)
        self.txt_resultado_consulta.insert("1.0", "\n".join(linhas))


# ==============================================================================
# FÁBRICA DO MÓDULO PARA O CONECTA HUB
# ==============================================================================

def construir_aba_controle_material(parent_frame, label_status_global=None, paletas=None, obter_tema_atual=None):
    """Instancia e retorna o controlador da aba Controle de Material."""
    return AbaControleMaterial(parent_frame, label_status_global, paletas, obter_tema_atual)
