# -*- coding: utf-8 -*-
"""
Módulo de Automação SAP: Mudança de Recurso
Executa a transferência integral de materiais e seriais entre duas equipes (Recurso Antigo -> Recurso Novo):
1. ZUDWM_237: Consulta saldos positivos e seriais da equipe de origem.
2. MIGO Z62: Estorno/devolução ao depósito (EMPR / CL03 / EMPR 1CL03) imputando nas ordens de origem.
3. IW39: Localização da Ordem Ativa do novo recurso (DY_PARVW="Z2", DY_PARNR=recurso_novo).
4. MIGO Z61: Saída/imputação para a ordem destino do novo recurso com os mesmos materiais e seriais.
5. Histórico e auditoria: Captura dos documentos SAP e gravação em CSV.
"""

import os
import sys
import time
import re
import csv
import uuid
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pythoncom
import win32com.client
import win32clipboard

NOME = "Mudança de Recurso"
DESCRICAO = "Transferência integral entre recursos via ZUDWM_237 + MIGO Z62 + IW39 + MIGO Z61"
MOVIMENTA_SAP = True

CENTRO_PADRAO = "EMPR"
DEPOSITO_PADRAO = "CL03"
RECEBEDOR_PADRAO = "EMPR 1CL03"
OPERACAO_PADRAO = "10"


def _pasta_assistente():
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    os.makedirs(pasta, exist_ok=True)
    return pasta


def _conectar_sap():
    """Obtém a sessão ativa do SAP GUI via COM."""
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        raise RuntimeError("SAP GUI não está em execução ou script COM desabilitado.") from e

    engine = sap_gui.GetScriptingEngine
    if not engine or engine.Children.Count == 0:
        raise RuntimeError("Nenhuma conexão SAP aberta encontrada.")

    for i in range(engine.Children.Count):
        conn = engine.Children(i)
        for j in range(conn.Children.Count):
            sessao = conn.Children(j)
            try:
                _ = sessao.Id
                return sessao
            except Exception:
                continue

    raise RuntimeError("Nenhuma sessão SAP disponível encontrada.")


def _aguardar_controle(sessao, identificadores, timeout=15):
    """Aguarda até que um dos identificadores esteja disponível na tela."""
    if isinstance(identificadores, str):
        identificadores = [identificadores]
    limite = time.time() + timeout
    while time.time() < limite:
        for ident in identificadores:
            try:
                ctrl = sessao.findById(ident)
                if ctrl is not None:
                    return ctrl
            except Exception:
                pass
        time.sleep(0.2)
    raise RuntimeError("Controle SAP não encontrado: " + " | ".join(identificadores))


def _controle_opcional(sessao, identificador):
    try:
        return sessao.findById(identificador)
    except Exception:
        return None


def _pressionar_enter(sessao, espera=0.6):
    sessao.findById("wnd[0]").sendVKey(0)
    time.sleep(espera)


def _fechar_janelas_modais(sessao):
    """Fecha quaisquer janelas modais pendentes (wnd[1], wnd[2], etc.)."""
    tentativas = 0
    while getattr(sessao, "Children", None) and sessao.Children.Count > 1 and tentativas < 5:
        tentativas += 1
        try:
            sessao.findById("wnd[1]/tbar[0]/btn[12]").press()
        except Exception:
            try:
                sessao.findById("wnd[1]").close()
            except Exception:
                try:
                    sessao.findById("wnd[1]/tbar[0]/btn[0]").press()
                except Exception:
                    break
        time.sleep(0.3)


def _navegar_transacao(sessao, tcode):
    """Navega para uma transação SAP de forma segura a partir de qualquer tela."""
    _fechar_janelas_modais(sessao)
    tcode_cmd = tcode if tcode.startswith("/") else f"/n{tcode}"
    try:
        sessao.findById("wnd[0]").maximize()
    except Exception:
        pass
    campo_okcd = _aguardar_controle(sessao, "wnd[0]/tbar[0]/okcd", timeout=12)
    campo_okcd.text = tcode_cmd
    sessao.findById("wnd[0]").sendVKey(0)
    time.sleep(1.0)
    _fechar_janelas_modais(sessao)


def _extrair_detalhe_popup(sessao):
    """Extrai mensagens detalhadas de janelas modais como 'Exibir logs'."""
    if not getattr(sessao, "Children", None) or sessao.Children.Count <= 1:
        return ""
    try:
        wnd = sessao.findById("wnd[1]")
        titulo = str(wnd.Text).strip()
        
        for id_tabela in [
            "wnd[1]/usr/cntlGRID1/shellcont/shell",
            "wnd[1]/usr/tblSAPLSBAL_DISPLAY",
            "wnd[1]/usr/ssubSUB_MAIN:SAPLSBAL_DISPLAY:0100/tblSAPLSBAL_DISPLAY"
        ]:
            ctrl = _controle_opcional(sessao, id_tabela)
            if ctrl is not None:
                try:
                    msgs = []
                    total_r = min(int(getattr(ctrl, "RowCount", 0)), 15)
                    for r in range(total_r):
                        item_num = ""
                        for col_item in ["ITEM", "ZEILE", "LINE", "POSNR"]:
                            try:
                                v_item = str(ctrl.getCellValue(r, col_item) or "").strip()
                                if v_item:
                                    item_num = v_item
                                    break
                            except Exception:
                                pass
                        
                        txt_msg = ""
                        for c in ["TEXT", "MSGTXT", "MESSAGE", "NATXT", "ARBTXT", "MSG_TEXT", "TEXTO", "TXT", "LTEXT", "LOG_TEXT"]:
                            try:
                                v = str(ctrl.getCellValue(r, c) or "").strip()
                                if v:
                                    txt_msg = v
                                    break
                            except Exception:
                                pass

                        if not txt_msg:
                            try:
                                col_ord = getattr(ctrl, "ColumnOrder", [])
                                for col_nome in col_ord:
                                    v = str(ctrl.getCellValue(r, col_nome) or "").strip()
                                    if v and len(v) > 4 and not v.isdigit() and col_nome != "ITEM":
                                        txt_msg = v
                                        break
                            except Exception:
                                pass

                        if txt_msg:
                            prefixo_item = f"Item {item_num}: " if item_num else ""
                            msgs.append(f"{prefixo_item}{txt_msg}")

                    if msgs:
                        return f"{titulo} -> " + " | ".join(msgs)
                except Exception:
                    pass
        return titulo
    except Exception:
        return "janela modal aberta"



def _analisar_saldo(valor):
    """
    Analisa rigorosamente o saldo retornado pelo SAP (coluna QTD_DISP).
    Retorna (eh_positivo: bool, eh_negativo: bool, valor_formatado: str).
    Detecta e isola todos os formatos negativos possíveis do SAP:
    - Trailing minus: '10-', '10,000-', '2-' (padrão ALV ABAP do SAP)
    - Leading minus: '-10', '-10.000', '-2,5'
    - Contábil entre parênteses: '(10)', '(10,000)'
    - Valores zerados ou nulos: '0', '0,000', '0.000', '', None
    """
    if valor is None:
        return (False, False, "0")
    t = str(valor).strip().replace(chr(160), "").replace(" ", "")
    if not t:
        return (False, False, "0")

    # Identifica qualquer representação de número negativo
    if "-" in t or (t.startswith("(") and t.endswith(")")):
        return (False, True, t)

    # Converte para float tratando separadores de milhar e decimal
    t_clean = t.replace(".", "").replace(",", ".")
    try:
        val = float(t_clean)
        if val > 0:
            return (True, False, _decimal_str(t))
        elif val < 0:
            return (False, True, t)
        else:
            return (False, False, "0")
    except ValueError:
        return (False, False, "0")


def _saldo_positivo(valor):
    """Retorna True apenas se o saldo for estritamente maior que zero e não negativo."""
    eh_pos, eh_neg, _ = _analisar_saldo(valor)
    return eh_pos and not eh_neg


def _decimal_str(valor):
    """Formata valor decimal para string aceita pelo SAP."""
    if not valor:
        return "0"
    t = str(valor).strip().replace(" ", "")
    if "," in t and "." in t:
        t = t.replace(".", "")
    return t


# =========================================================================
# ETAPA 1: CONSULTA ZUDWM_237
# =========================================================================

def consultar_origem(sessao, equipe_origem, log_callback):
    """
    Executa ZUDWM_237 para o recurso de origem, ordena por status e
    captura materiais, quantidades disponíveis e números de série.
    """
    log_callback(f"🔎 Acessando ZUDWM_237 para o recurso de origem '{equipe_origem}'...\n")
    _navegar_transacao(sessao, "ZUDWM_237")

    campo_equipe = _aguardar_controle(sessao, "wnd[0]/usr/ctxtS_EQUIPE-LOW", timeout=12)
    campo_equipe.text = equipe_origem
    campo_equipe.caretPosition = len(equipe_origem)
    _pressionar_enter(sessao, 0.4)

    # Garante que 'Incluir ordens encerradas' esteja DESMARCADO (apenas ordens vigentes/ativas)
    chk_enc = _controle_opcional(sessao, "wnd[0]/usr/chkP_ORDENC")
    if chk_enc is not None:
        try:
            chk_enc.selected = False
        except Exception:
            pass

    # Executa F8
    sessao.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1.5)

    grid = _aguardar_controle(sessao, "wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell", timeout=15)

    # Ordenar por DESC_STATUS decrescente (conforme script gravado)
    try:
        grid.setCurrentCell(-1, "DESC_STATUS")
        grid.selectColumn("DESC_STATUS")
        grid.pressToolbarButton("&SORT_DSC")
        time.sleep(0.8)
        # Reobter grid após ordenação
        grid = _aguardar_controle(sessao, "wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell", timeout=10)
    except Exception as e:
        log_callback(f"⚠️ Aviso ao ordenar grid: {e}\n")

    total_linhas = 0
    try:
        total_linhas = int(grid.RowCount)
    except Exception:
        total_linhas = 0

    colunas_grid = []
    try:
        ordem_c = grid.ColumnOrder
        colunas_grid = [ordem_c.Item(c) for c in range(ordem_c.Count)]
    except Exception:
        pass

    log_callback(f"📋 ZUDWM_237 carregada. Total de linhas: {total_linhas} | Colunas: {', '.join(colunas_grid) if colunas_grid else 'padrão'}\n")

    itens_coletados = []
    linhas_negativas_ignoradas = 0
    ordem_corrente = ""

    for idx in range(total_linhas):
        ordem_val = ""
        for col_ord in ["AUFNR", "ORDEM", "NUM_ORDEM"]:
            try:
                v = str(grid.getCellValue(idx, col_ord)).strip()
                if v:
                    ordem_val = v
                    break
            except Exception:
                pass

        if ordem_val:
            ordem_corrente = ordem_val

        qtd_disp_raw = ""
        for col_qtd in ["QTD_DISP", "MENGE", "SALDO", "DISPONIVEL"]:
            try:
                v = grid.getCellValue(idx, col_qtd)
                if v is not None and str(v).strip() != "":
                    qtd_disp_raw = str(v).strip()
                    break
            except Exception:
                pass

        material = ""
        for col_mat in ["MATNR", "MATERIAL", "ID_MATERIAL", "COD_MATERIAL"]:
            try:
                v = str(grid.getCellValue(idx, col_mat)).strip()
                if v:
                    material = v
                    break
            except Exception:
                pass

        descricao = ""
        for col_desc in ["MAKTX", "DESC_MATERIAL", "ARKTX", "TXZ01"]:
            try:
                v = str(grid.getCellValue(idx, col_desc)).strip()
                if v:
                    descricao = v
                    break
            except Exception:
                pass

        unidade = "UN"
        for col_uni in ["MEINS", "UNIDADE", "MEINH", "EINHEIT"]:
            try:
                v = str(grid.getCellValue(idx, col_uni)).strip()
                if v:
                    unidade = v
                    break
            except Exception:
                pass

        fim_base = ""
        for col_fim in ["GLTRP", "FIM_BASE", "DATA_FIM"]:
            try:
                v = str(grid.getCellValue(idx, col_fim)).strip()
                if v:
                    fim_base = v
                    break
            except Exception:
                pass

        eh_positivo, eh_negativo, qtd_limpa = _analisar_saldo(qtd_disp_raw)

        # Regra de Auditoria ENEL: Linhas com saldo negativo DEVEM ser explicitamente ignoradas
        if eh_negativo:
            linhas_negativas_ignoradas += 1
            log_callback(
                f"  🚫 LINHA NEGATIVA IGNORADA (Linha {idx+1}): "
                f"Material {material or '(sem código)'} ({descricao}) | "
                f"Saldo negativo: '{qtd_disp_raw}' | Ordem: {ordem_corrente} -> Descartado com segurança.\n"
            )
            continue

        if not eh_positivo:
            # Saldo zerado ou vazio
            continue

        if not material:
            log_callback(f"  ⚠️ Linha {idx+1} com saldo positivo ({qtd_limpa}), mas sem código de material identificado. Ignorando.\n")
            continue

        qtd_formatada = qtd_limpa
        log_callback(f"  📦 Item com saldo positivo (Linha {idx+1}): {material} - {descricao} | Qtd: {qtd_formatada} | Ordem: {ordem_corrente}\n")

        # Clicar na célula QTD_DISP para verificar e extrair seriais da janela modal (medidores ou serializados)
        seriais_item = []
        try:
            grid.currentCellColumn = "QTD_DISP"
            grid.currentCellRow = idx
            grid.selectedRows = str(idx)
            grid.clickCurrentCell()
            time.sleep(1.2)

            # Verifica se abriu popup wnd[1]
            popup_grid = _controle_opcional(sessao, "wnd[1]/usr/cntlGRID1/shellcont/shell")
            if popup_grid is not None:
                total_s = int(popup_grid.RowCount)
                for s_idx in range(total_s):
                    s_num = ""
                    for col_s in ["SERNR", "SERIAL", "EQUIPMENT", "GERAET"]:
                        try:
                            val_s = str(popup_grid.getCellValue(s_idx, col_s)).strip()
                            if val_s:
                                s_num = val_s
                                break
                        except Exception:
                            pass
                    if s_num:
                        seriais_item.append(s_num)
                
                # Fecha popup modal
                fechou = False
                try:
                    sessao.findById("wnd[1]").close()
                    fechou = True
                except Exception:
                    pass
                if not fechou:
                    try:
                        sessao.findById("wnd[1]/tbar[0]/btn[0]").press()
                    except Exception:
                        pass
                time.sleep(0.5)
        except Exception as e:
            log_callback(f"    (Sem leitura de seriais no popup: {e})\n")

        if seriais_item:
            log_callback(f"    🔢 Material serializado/medidor ({len(seriais_item)} serial(is)): {', '.join(seriais_item)}\n")

        itens_coletados.append({
            "Ordem Origem": ordem_corrente,
            "Material": material,
            "Descricao": descricao,
            "Quantidade": qtd_formatada,
            "Unidade": unidade or "UN",
            "Fim Base": fim_base,
            "Seriais": seriais_item,
        })

    # Sai da ZUDWM_237 fechando telas pendentes
    _fechar_janelas_modais(sessao)

    log_callback(
        f"✅ Varredura ZUDWM_237 concluída: {len(itens_coletados)} item(ns) com saldo positivo identificados. "
        f"({linhas_negativas_ignoradas} linha(s) negativa(s) ignoradas com segurança para auditoria ENEL).\n"
    )

    # Conferência de status de medidores e seriais na ZMMCSE001
    todos_seriais = []
    for it in itens_coletados:
        todos_seriais.extend(it.get("Seriais", []))

    if todos_seriais:
        log_callback(f"\n🔬 Iniciando conferência obrigatória de {len(todos_seriais)} serial(is) na ZMMCSE001...\n")
        try:
            res_conf = conferir_seriais_zmmcse001(sessao, equipe_origem, todos_seriais, log_callback)
            itens_coletados = aplicar_filtro_conferencia_medidores(itens_coletados, res_conf, log_callback)
        except Exception as e:
            log_callback(f"❌ ERRO NA CONFERÊNCIA ZMMCSE001: {e}\n")
            raise RuntimeError(f"Falha na validação de seriais na ZMMCSE001: {e}. Transferência bloqueada para proteção contábil.")

    # Resumo visual das ordens de origem
    _exibir_resumo_ordens(itens_coletados, log_callback, fase="ZUDWM_237")

    return itens_coletados


# =========================================================================
# ETAPA 1.5: CONFERÊNCIA DE MEDIDORES E STATUS SAP (ZMMCSE001)
# =========================================================================

def _normalizar_status(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().upper()


def _normalizar_serial_chave(valor):
    """Normaliza serial removendo zeros à esquerda para comparação com SAP."""
    texto = str(valor or "").strip().upper()
    if texto.isdigit():
        return texto.lstrip("0") or "0"
    return texto


def _exibir_resumo_ordens(itens, log_callback, fase="Planejamento"):
    """Exibe agrupamento detalhado dos materiais por Ordem de Origem para MIGO Z62."""
    if not itens:
        return

    ordens = {}
    for it in itens:
        ord_orig = it.get("Ordem Origem") or "(sem ordem informada)"
        ordens.setdefault(ord_orig, []).append(it)

    log_callback(f"\n📋 [IMPUTAÇÃO CONTÁBIL POR ORDEM DE ORIGEM - {fase.upper()}]\n")
    log_callback(f"ℹ️ Total de ordens distintas de origem: {len(ordens)}\n")
    for num_ordem, materiais in ordens.items():
        log_callback(f"  🔹 Ordem Origem: {num_ordem} ({len(materiais)} material/materiais elegíveis):\n")
        for m in materiais:
            seriais = m.get("Seriais", [])
            ser_txt = f" | {len(seriais)} serial(is)" if seriais else " | Não serializado"
            log_callback(f"     • {m['Material']} - {m['Descricao']} | Qtd: {m['Quantidade']} {m['Unidade']}{ser_txt}\n")
    log_callback("\n")


def conferir_seriais_zmmcse001(sessao, equipe, seriais, log_callback):
    """
    Acessa a transação ZMMCSE001 > Análise de Equipe (BT003) e consulta
    em lote todos os seriais de medidores coletados.
    Retorna {serial_original: {"apto": bool, "status_sis": str, "status_usr": str, "mov": str, "recurso": str}}
    """
    if not seriais:
        return {}

    log_callback(f"\n🔬 [CONFERÊNCIA ZMMCSE001] Verificando status de {len(seriais)} serial(is) de medidor(es)...\n")
    
    _navegar_transacao(sessao, "ZMMCSE001")

    botao_analise = _aguardar_controle(sessao, "wnd[0]/usr/btnBT003", timeout=25)
    botao_analise.press()
    time.sleep(1.0)

    base = "wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/"
    radio_equipamento = _aguardar_controle(sessao, base + "radRD_EQUI", timeout=15)
    radio_equipamento.selected = True

    campo_serial = _aguardar_controle(sessao, base + "ctxtS_SERNR-LOW", timeout=10)

    # Abre seleção múltipla
    sessao.findById(base + "btn%_S_SERNR_%_APP_%-VALU_PUSH").press()
    _aguardar_controle(sessao, "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA", timeout=12)

    texto_clipboard = "\r\n".join(seriais)
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, texto_clipboard)
    finally:
        win32clipboard.CloseClipboard()

    # Clica no botão colar (btn[24])
    sessao.findById("wnd[1]/tbar[0]/btn[24]").press()
    time.sleep(0.5)
    # Executa F8 na janela modal
    sessao.findById("wnd[1]/tbar[0]/btn[8]").press()
    _aguardar_controle(sessao, base + "ctxtS_SERNR-LOW", timeout=10)

    # Executa F8 na tela principal
    sessao.findById("wnd[0]").sendVKey(8)
    time.sleep(1.5)

    grid_id = "wnd[0]/usr/cntlCO_ITEM/shellcont/shell"
    grid = _aguardar_controle(sessao, grid_id, timeout=30)

    fim_espera = time.time() + 25
    while time.time() < fim_espera and int(grid.RowCount) < 1:
        time.sleep(0.3)

    por_serial = {}
    total_linhas = int(grid.RowCount)
    try:
        linhas_visiveis = max(1, int(grid.VisibleRowCount))
    except Exception:
        linhas_visiveis = 10

    for inicio in range(0, total_linhas, linhas_visiveis):
        try:
            grid.FirstVisibleRow = inicio
            time.sleep(0.1)
        except Exception:
            pass

        fim = min(inicio + linhas_visiveis, total_linhas)
        for linha in range(inicio, fim):
            s_grid = ""
            for col_s in ["SERIAL", "SERNR", "EQUIPMENT"]:
                try:
                    v = str(grid.getCellValue(linha, col_s) or "").strip()
                    if v:
                        s_grid = v
                        break
                except Exception:
                    pass

            if not s_grid:
                continue

            mat_grid = ""
            for c in ["MATNR", "MATERIAL"]:
                try:
                    v = str(grid.getCellValue(linha, c) or "").strip()
                    if v:
                        mat_grid = v
                        break
                except Exception:
                    pass

            ordem_grid = ""
            for c in ["AUFNR", "ORDEM", "NUM_ORDEM"]:
                try:
                    v = str(grid.getCellValue(linha, c) or "").strip()
                    if v:
                        ordem_grid = v
                        break
                except Exception:
                    pass

            sts_sis = ""
            for c in ["STATUS", "STS_SISTEMA"]:
                try:
                    v = str(grid.getCellValue(linha, c) or "").strip()
                    if v:
                        sts_sis = v
                        break
                except Exception:
                    pass

            sts_usr = ""
            for c in ["STUSER", "STS_USUARIO"]:
                try:
                    v = str(grid.getCellValue(linha, c) or "").strip()
                    if v:
                        sts_usr = v
                        break
                except Exception:
                    pass

            st_mov = ""
            for c in ["ST_MOV", "MOVIMENTO", "ST_MOVIMENTO"]:
                try:
                    v = str(grid.getCellValue(linha, c) or "").strip()
                    if v:
                        st_mov = v
                        break
                except Exception:
                    pass

            parnr = ""
            for c in ["PARNR", "RECURSO", "EQUIPE"]:
                try:
                    v = str(grid.getCellValue(linha, c) or "").strip()
                    if v:
                        parnr = v
                        break
                except Exception:
                    pass

            por_serial[_normalizar_serial_chave(s_grid)] = {
                "Material": mat_grid,
                "Ordem Origem": ordem_grid,
                "Recurso": parnr,
                "STS Sistema": sts_sis,
                "STS Usuario": sts_usr,
                "ST Mov Mercadoria": st_mov,
            }

    # Sai da ZMMCSE001 fechando modais
    _fechar_janelas_modais(sessao)

    resultado_conferencia = {}
    for s in seriais:
        chave = _normalizar_serial_chave(s)
        dados = por_serial.get(chave)
        if not dados:
            log_callback(f"  ⚠️ Serial {s}: NÃO localizado no cockpit da equipe na ZMMCSE001.\n")
            resultado_conferencia[s] = {
                "apto": False,
                "motivo": "Não localizado no grid ZMMCSE001",
                "status_sis": "",
                "status_usr": "",
                "mov": "",
                "recurso": "",
                "ordem_origem": "",
            }
            continue

        sts_sis = dados["STS Sistema"].upper()
        sts_usr = dados["STS Usuario"].upper()
        st_mov = dados["ST Mov Mercadoria"].upper()
        rec = dados["Recurso"].upper()
        ordem_serial = dados["Ordem Origem"].upper()

        # Regra Estrita: Apenas medidores com LIDI no Sistema E DISP no Usuário (sem restrições)
        eh_lidi_disp = ("LIDI" in sts_sis) and ("DISP" in sts_usr)
        eh_inapto = (
            any(x in sts_sis for x in ["MONT", "INST", "INTS", "RETE", "DEFI"]) or
            any(x in sts_usr for x in ["RETE", "DEFI", "MONT", "INST"]) or
            ("UTILIZADO" in st_mov)
        )
        eh_mov_disp = ("DISPONIVEL" in st_mov or "DISP" in st_mov or not st_mov) and ("UTILIZADO" not in st_mov)

        # Valida se pertence à equipe de origem (se informado na ZMMCSE001)
        eh_recurso_correto = (not rec) or (_normalizar_status(rec) == _normalizar_status(equipe))

        apto = eh_lidi_disp and (not eh_inapto) and eh_mov_disp and eh_recurso_correto

        if apto:
            log_callback(f"  ✅ Serial {s}: STS Sis={sts_sis} | STS Usr={sts_usr} | Mov={st_mov} | Rec={rec} | Ord={ordem_serial} -> APTO LIDI DISP\n")
            resultado_conferencia[s] = {
                "apto": True,
                "motivo": "LIDI DISP Disponível",
                "status_sis": sts_sis,
                "status_usr": sts_usr,
                "mov": st_mov,
                "recurso": rec,
                "ordem_origem": ordem_serial,
            }
        else:
            motivos = []
            if not eh_recurso_correto:
                motivos.append(f"Recurso divergente ({rec} != {equipe})")
            if not eh_lidi_disp:
                motivos.append(f"Requer LIDI DISP (encontrado Sis='{sts_sis}', Usr='{sts_usr}')")
            if any(x in sts_sis for x in ["MONT", "INST", "INTS"]):
                motivos.append("Montado/Instalado")
            if "RETE" in sts_sis or "RETE" in sts_usr:
                motivos.append("Retido")
            if "DEFI" in sts_sis or "DEFI" in sts_usr:
                motivos.append("Defeito")
            if "UTILIZADO" in st_mov:
                motivos.append("Utilizado")
            if not motivos:
                motivos.append("Status divergente")
            motivo_str = " / ".join(motivos)

            log_callback(f"  🚫 Serial {s}: STS Sis={sts_sis} | STS Usr={sts_usr} | Mov={st_mov} | Rec={rec} | Ord={ordem_serial} -> INAPTO ({motivo_str}) - DESCARTADO\n")
            resultado_conferencia[s] = {
                "apto": False,
                "motivo": motivo_str,
                "status_sis": sts_sis,
                "status_usr": sts_usr,
                "mov": st_mov,
                "recurso": rec,
                "ordem_origem": ordem_serial,
            }

    return resultado_conferencia


def aplicar_filtro_conferencia_medidores(itens, resultado_conferencia, log_callback):
    """
    Ajusta a lista de itens com base na conferência de seriais.
    Para itens serializados (medidores):
    - Mantém apenas seriais com apto == True E pertencentes à ordem informada.
    - REAJUSTA RIGOROSAMENTE a quantidade do item para coincidir com a contagem de seriais aptos.
    - Se nenhum serial for apto, remove o item da transferência.
    """
    if not resultado_conferencia:
        return itens

    itens_filtrados = []
    for it in itens:
        seriais_orig = it.get("Seriais", [])
        if not seriais_orig:
            itens_filtrados.append(it)
            continue

        ordem_item = str(it.get("Ordem Origem") or "").strip().upper()
        seriais_validos = []

        for s in seriais_orig:
            conf = resultado_conferencia.get(s, {})
            if not conf.get("apto", False):
                continue

            ordem_conf = str(conf.get("ordem_origem") or "").strip().upper()
            # Se a conferência na ZMMCSE001 retornou uma ordem específica, deve bater com a ordem do item
            if ordem_conf and ordem_item and (_normalizar_status(ordem_conf) != _normalizar_status(ordem_item)):
                log_callback(
                    f"  🚫 Serial {s} DESCARTADO: Pertence à ordem {ordem_conf} na ZMMCSE001, "
                    f"mas o item {it.get('Material')} está na ordem {ordem_item}!\n"
                )
                continue

            seriais_validos.append(s)

        if not seriais_validos:
            log_callback(
                f"⚠️ Material {it['Material']} ({it['Descricao']}): "
                f"Todos os {len(seriais_orig)} serial(is) foram reprovados (fora de LIDI DISP ou ordem divergente). "
                f"Item 100% EXCLUÍDO da transferência para evitar inconsistência SAP.\n"
            )
            continue

        qtd_antiga = str(it.get("Quantidade", "0")).strip()
        qtd_aprovada = len(seriais_validos)
        nova_qtd = str(qtd_aprovada)

        qtd_limpa_antiga = qtd_antiga.replace(",000", "").replace(".000", "").strip()
        if qtd_limpa_antiga != nova_qtd:
            log_callback(
                f"ℹ️ Material {it['Material']} ({it['Descricao']}): "
                f"Quantidade reajustada de {qtd_antiga} para {nova_qtd} "
                f"({qtd_aprovada} medidor(es) LIDI DISP apto(s) de {len(seriais_orig)}).\n"
            )
        else:
            log_callback(
                f"✅ Material {it['Material']} ({it['Descricao']}): "
                f"Quantidade confirmada em {nova_qtd} ({qtd_aprovada} medidor(es) LIDI DISP aprovados).\n"
            )

        it_copia = dict(it)
        it_copia["Seriais"] = seriais_validos
        it_copia["Quantidade"] = nova_qtd
        itens_filtrados.append(it_copia)

    return itens_filtrados


# =========================================================================
# ETAPA 2: BUSCA DA ORDEM ATIVA NA IW39
# =========================================================================

def buscar_ordem_destino_iw39(sessao, equipe_destino, log_callback):
    """
    Acessa a IW39 com DY_PARVW = 'Z2' e DY_PARNR = equipe_destino,
    extraindo a ordem ativa para imputação no Z61.
    """
    log_callback(f"🔎 Consultando IW39 para localizar ordem ativa da equipe destino '{equipe_destino}'...\n")
    _navegar_transacao(sessao, "IW39")

    combo_parceiro = _aguardar_controle(sessao, "wnd[0]/usr/cmbDY_PARVW", timeout=12)
    combo_parceiro.key = "Z2"
    
    campo_parceiro = _aguardar_controle(sessao, "wnd[0]/usr/ctxtDY_PARNR", timeout=10)
    campo_parceiro.text = equipe_destino
    campo_parceiro.caretPosition = len(equipe_destino)
    _pressionar_enter(sessao, 0.4)

    # Status da Ordem na IW39: Apenas ordens ativas (Pendente e Em Processamento), NUNCA Encerrado ou Histórico
    status_config = [
        ("wnd[0]/usr/chkDY_OFN", True),   # Pendente
        ("wnd[0]/usr/chkDY_IAR", True),   # Em processamento
        ("wnd[0]/usr/chkDY_MAB", False),  # Encerrado (MAB = Mit Abgeschlossenen) -> NUNCA MARCAR
        ("wnd[0]/usr/chkDY_ABG", False),  # Encerrado alternativo
        ("wnd[0]/usr/chkDY_HIS", False),  # Histórico -> DESMARCADO
    ]
    for chk_id, marcar in status_config:
        ctrl = _controle_opcional(sessao, chk_id)
        if ctrl is not None:
            try:
                ctrl.selected = marcar
            except Exception:
                pass

    # Executa F8
    sessao.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(1.2)

    ordem_encontrada = ""

    # Caso 1: Abre diretamente no detalhe da Ordem (uma única ordem)
    campo_ordem_detalhe = _controle_opcional(
        sessao,
        "wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR"
    )
    if campo_ordem_detalhe is not None:
        ordem_encontrada = str(campo_ordem_detalhe.text).strip()

    # Caso 2: Abre na grade de ordens (múltiplas ordens)
    if not ordem_encontrada:
        grid_iw39 = _controle_opcional(sessao, "wnd[0]/usr/cntlGRID1/shellcont/shell")
        if grid_iw39 is not None:
            try:
                # Ordena por GLTRP decrescente para pegar a mais recente
                grid_iw39.setCurrentCell(-1, "GLTRP")
                grid_iw39.selectColumn("GLTRP")
                sessao.findById("wnd[0]/tbar[1]/btn[40]").press() # Ordenação decrescente
                time.sleep(0.5)
                grid_iw39 = sessao.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            except Exception:
                pass

            total_ordens = int(grid_iw39.RowCount)
            for r in range(total_ordens):
                try:
                    num_ord = str(grid_iw39.getCellValue(r, "AUFNR")).strip()
                    rec_ord = str(grid_iw39.getCellValue(r, "PARNR")).strip().upper()
                    if num_ord:
                        if not rec_ord or rec_ord == equipe_destino.upper():
                            ordem_encontrada = num_ord
                            break
                except Exception:
                    pass

    # Retorna para tela inicial
    try:
        sessao.findById("wnd[0]/tbar[0]/btn[15]").press()
        sessao.findById("wnd[0]/tbar[0]/btn[15]").press()
    except Exception:
        pass

    if not ordem_encontrada:
        raise RuntimeError(f"Nenhuma ordem ativa localizada na IW39 para o recurso de destino '{equipe_destino}'.")

    log_callback(f"🎯 Ordem ativa de destino localizada com sucesso: {ordem_encontrada}\n")
    return ordem_encontrada


# =========================================================================
# ETAPA 3 & 4: PREENCHIMENTO E EXECUÇÃO MIGO (Z62 e Z61)
# =========================================================================

def _ids_aba(nome):
    return [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{p}/"
        f"subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/"
        f"tabsTS_GOITEM/tabpOK_GOITEM_{nome}"
        for p in ("0003", "0002")
    ]


def _garantir_detalhe_aberto(sessao):
    for id_mat in _ids_aba("MATERIAL"):
        if _controle_opcional(sessao, id_mat) is not None:
            return
    botoes_abrir = [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL",
    ]
    try:
        _aguardar_controle(sessao, botoes_abrir, timeout=6).press()
    except Exception:
        pass
    _aguardar_controle(sessao, _ids_aba("MATERIAL"), timeout=10)


def _selecionar_aba(sessao, nome):
    _garantir_detalhe_aberto(sessao)
    for id_aba in _ids_aba(nome):
        ctrl = _controle_opcional(sessao, id_aba)
        if ctrl is not None:
            ctrl.select()
            return id_aba
    raise RuntimeError(f"Aba MIGO '{nome}' não encontrada.")


def _abrir_migo(sessao, movimento):
    _navegar_transacao(sessao, "MIGO")
    _garantir_detalhe_aberto(sessao)

    # Ação: A01 para Z62, A07 para Z61 (ou conforme padrão de tela)
    acoes = [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{p}/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION"
        for p in ("0003", "0002")
    ]
    try:
        _aguardar_controle(sessao, acoes, timeout=8).Key = "A01" if movimento == "Z62" else "A07"
        time.sleep(0.4)
    except Exception:
        pass

    # Movimento Z62 / Z61
    campos_bwart = [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{p}/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART"
        for p in ("0003", "0002")
    ]
    campo_bw = _aguardar_controle(sessao, campos_bwart, timeout=8)
    campo_bw.text = movimento
    _pressionar_enter(sessao, 0.8)


def _preencher_seriais_migo(sessao, seriais, log_callback):
    """Preenche os números de série no MIGO."""
    if not seriais:
        return

    _selecionar_aba(sessao, "SERIAL")
    time.sleep(0.3)

    # Se forem até 2 seriais, preenche diretamente nas linhas da grade se disponível
    if len(seriais) <= 2:
        try:
            for s_idx, serial in enumerate(seriais):
                id_s = f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,{s_idx}]"
                campo_serial = _controle_opcional(sessao, id_s)
                if campo_serial is not None:
                    campo_serial.text = serial
            _pressionar_enter(sessao, 0.4)
            return
        except Exception:
            pass

    # Preenchimento em lote via botão de busca e clipboard
    btn_search = _controle_opcional(
        sessao,
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_SEARCH"
    )
    if btn_search is not None:
        btn_search.press()
        time.sleep(0.5)

        btn_multiplos = _aguardar_controle(sessao, "wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH", timeout=8)
        btn_multiplos.press()
        time.sleep(0.5)

        # Coloca os seriais no clipboard
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, "\r\n".join(seriais))
        finally:
            win32clipboard.CloseClipboard()

        # Clica no botão colar (btn[24])
        _aguardar_controle(sessao, "wnd[1]/tbar[0]/btn[24]", timeout=8).press()
        time.sleep(0.3)
        # Executa F8 na janela modal
        _aguardar_controle(sessao, "wnd[1]/tbar[0]/btn[8]", timeout=8).press()
        time.sleep(0.5)
        # Executa F8 na busca de seriais
        _aguardar_controle(sessao, "wnd[0]/tbar[1]/btn[8]", timeout=8).press()
        time.sleep(0.8)

        # Seleciona todos e transfere
        grid_ser = _controle_opcional(sessao, "wnd[0]/usr/cntlGRID1/shellcont/shell")
        if grid_ser is not None:
            try:
                grid_ser.setCurrentCell(-1, "")
                grid_ser.selectAll()
                _aguardar_controle(sessao, "wnd[0]/tbar[1]/btn[47]", timeout=6).press()
                time.sleep(0.5)
            except Exception as e:
                log_callback(f"⚠️ Aviso na seleção de seriais: {e}\n")


def _preencher_item_migo(sessao, item, ordem_contabil, movimento, indice, total_itens, log_callback):
    """Preenche um item de mercadoria na MIGO."""
    log_callback(f"  📝 Preenchendo item {indice+1}/{total_itens}: Material {item['Material']} (Qtd: {item['Quantidade']}) | Ordem: {ordem_contabil}...\n")

    # 1. Aba Material
    aba_mat = _selecionar_aba(sessao, "MATERIAL")
    campo_mat = _aguardar_controle(
        sessao,
        aba_mat + "/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX",
        timeout=8
    )
    campo_mat.text = item["Material"]
    campo_mat.caretPosition = len(item["Material"])
    _pressionar_enter(sessao, 0.6)

    # 2. Aba Quantidade
    aba_qtd = _selecionar_aba(sessao, "QUANTITIES")
    campo_qtd = _aguardar_controle(
        sessao,
        aba_qtd + "/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG",
        timeout=8
    )
    campo_qtd.text = item["Quantidade"]
    _pressionar_enter(sessao, 0.6)

    # 3. Aba Destino
    aba_dest = _selecionar_aba(sessao, "DESTINAT.")
    prefixo_dest = aba_dest + "/ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/"
    try:
        _aguardar_controle(sessao, prefixo_dest + "ctxtGOITEM-NAME1", timeout=5).text = CENTRO_PADRAO
        _aguardar_controle(sessao, prefixo_dest + "ctxtGOITEM-LGOBE", timeout=5).text = DEPOSITO_PADRAO
        _aguardar_controle(sessao, prefixo_dest + "txtGOITEM-WEMPF", timeout=5).text = RECEBEDOR_PADRAO
        _pressionar_enter(sessao, 0.6)
    except Exception as e:
        log_callback(f"⚠️ Aviso no preenchimento de destino: {e}\n")

    # 4. Aba Classificação Contábil (Ordem e Operação)
    aba_conta = _selecionar_aba(sessao, "ACCOUNT")
    prefixo_conta = aba_conta + "/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/"
    campo_ordem = _aguardar_controle(sessao, prefixo_conta + "ctxtCOBL-AUFNR", timeout=8)
    campo_ordem.text = ordem_contabil
    campo_operacao = _aguardar_controle(sessao, prefixo_conta + "ctxtCOBL-VORNR_AUF", timeout=8)
    campo_operacao.text = OPERACAO_PADRAO
    _pressionar_enter(sessao, 0.6)

    # 5. Aba Serial (se houver seriais para o item)
    if item.get("Seriais"):
        _preencher_seriais_migo(sessao, item["Seriais"], log_callback)

    # Marca Item OK se disponível
    chk_ok = _controle_opcional(
        sessao,
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE"
    )
    if chk_ok is not None:
        try:
            chk_ok.selected = True
            _pressionar_enter(sessao, 0.3)
        except Exception:
            pass


def _avancar_proximo_item(sessao):
    botoes_next = [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM",
    ]
    _aguardar_controle(sessao, botoes_next, timeout=8).press()
    time.sleep(0.6)
    _garantir_detalhe_aberto(sessao)


def _verificar_e_gravar(sessao, movimento, log_callback):
    """Pressiona Verificar (btn[7]) e Gravar (btn[23]), retornando o Doc. Material."""
    log_callback(f"🔍 Executando verificação de consistência no SAP para {movimento}...\n")
    sessao.findById("wnd[0]/tbar[1]/btn[7]").press()
    time.sleep(1.5)

    # Verifica se abriu popup de erro ou log
    try:
        if sessao.Children.Count > 1:
            detalhe = _extrair_detalhe_popup(sessao)
            raise RuntimeError(f"SAP abriu mensagem durante Verificar: {detalhe}")
    except Exception as e:
        if "SAP abriu mensagem" in str(e):
            raise

    sbar_text = str(sessao.findById("wnd[0]/sbar").Text).strip()
    sbar_type = str(sessao.findById("wnd[0]/sbar").MessageType).upper()
    log_callback(f"  Status SAP (Verificar): {sbar_text}\n")

    if sbar_type in {"E", "A"}:
        raise RuntimeError(f"Erro na verificação do {movimento}: {sbar_text}")

    log_callback(f"💾 Registrando movimento {movimento} no SAP...\n")
    sessao.findById("wnd[0]/tbar[1]/btn[23]").press()

    limite = time.time() + 45
    doc_material = ""
    msg_final = ""

    while time.time() < limite:
        try:
            if sessao.Children.Count > 1:
                detalhe = _extrair_detalhe_popup(sessao)
                raise RuntimeError(f"SAP exibiu mensagem ao gravar {movimento}: {detalhe}")
        except Exception as e:
            if "SAP exibiu mensagem" in str(e):
                raise

        sbar = sessao.findById("wnd[0]/sbar")
        msg_final = str(sbar.Text).strip()
        m_tipo = str(sbar.MessageType).upper()

        match = re.search(r"Doc\.?\s*material\s+(\d+)\s+registrad", msg_final, re.IGNORECASE)
        if match:
            doc_material = match.group(1)
            break

        if m_tipo in {"E", "A"}:
            raise RuntimeError(f"SAP recusou gravação de {movimento}: {msg_final}")

        time.sleep(0.5)

    if not doc_material:
        raise RuntimeError(f"Tempo esgotado aguardando documento SAP de {movimento}. Última mensagem: {msg_final}")

    log_callback(f"🎉 DOCUMENTO SAP {movimento} REGISTRADO COM SUCESSO: {doc_material}\n")
    return doc_material


# =========================================================================
# ORQUESTRADOR PRINCIPAL
# =========================================================================

def executar_mudanca_recurso(recurso_antigo, recurso_novo, log_callback, serial_callback=None, progresso_callback=None):
    """
    Executa o fluxo completo de Mudança de Recurso:
    1. Consulta ZUDWM_237 + Auditoria ZMMCSE001 (Etapa 1/4 - 25%)
    2. MIGO Z62 Estorno na origem (Etapa 2/4 - 50%)
    3. IW39 Ordem ativa destino (Etapa 3/4 - 75%)
    4. MIGO Z61 Saída no destino (Etapa 4/4 - 100%)
    5. Gravação em histórico
    """
    recurso_antigo = str(recurso_antigo or "").strip().upper()
    recurso_novo = str(recurso_novo or "").strip().upper()

    if not recurso_antigo or not recurso_novo:
        raise ValueError("Informe o recurso de origem (antigo) e o recurso de destino (novo).")

    if recurso_antigo == recurso_novo:
        raise ValueError("O recurso de origem e o recurso de destino devem ser diferentes.")

    log_callback(f"\n=======================================================\n")
    log_callback(f"🚀 INICIANDO MUDANÇA DE RECURSO: {recurso_antigo} ➡️ {recurso_novo}\n")
    log_callback(f"Data e Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    log_callback(f"Configuração: Centro {CENTRO_PADRAO} | Depósito {DEPOSITO_PADRAO} | Recebedor {RECEBEDOR_PADRAO}\n")
    log_callback(f"=======================================================\n\n")

    pythoncom.CoInitialize()
    etapa_atual = 1
    try:
        if progresso_callback:
            progresso_callback(1, 0.05, "Etapa 1/4: Conectando ao SAP GUI...")

        sessao = _conectar_sap()

        # ETAPA 1: Consulta ZUDWM_237 e Conferência ZMMCSE001
        etapa_atual = 1
        if progresso_callback:
            progresso_callback(1, 0.25, "Etapa 1/4 (25%): Consulta ZUDWM_237 e Auditoria de Medidores ZMMCSE001...")
        log_callback("--- [1/4] CONSULTA DE SALDOS NA ORIGEM (ZUDWM_237) ---\n")
        itens = consultar_origem(sessao, recurso_antigo, log_callback)
        if not itens:
            log_callback(f"⚠️ Nenhum saldo disponível positivo encontrado para o recurso {recurso_antigo}. Processo cancelado sem movimentações.\n")
            if progresso_callback:
                progresso_callback(1, 1.0, "Cancelado: Sem saldos positivos.")
            return {"status": "sem_saldos", "mensagem": f"Nenhum saldo positivo para {recurso_antigo}"}

        # Filtro de Segurança / Auditoria ENEL pré-MIGO (garantia dupla contra qualquer quantidade negativa ou zerada)
        itens_auditados = []
        for it in itens:
            eh_p, eh_n, _ = _analisar_saldo(it.get("Quantidade"))
            if eh_p and not eh_n:
                itens_auditados.append(it)
            else:
                log_callback(f"  🚫 Auditoria de segurança descartou item com quantidade não-positiva: {it.get('Material')} (Qtd: {it.get('Quantidade')})\n")
        
        if not itens_auditados:
            log_callback("⚠️ Nenhum item elegível após a auditoria de segurança. Nenhuma movimentação contábil será realizada.\n")
            if progresso_callback:
                progresso_callback(1, 1.0, "Cancelado: Itens inelegíveis.")
            return {"status": "sem_saldos", "mensagem": "Nenhum item válido após auditoria"}
        
        itens = itens_auditados

        if serial_callback:
            todos_seriais = []
            for it in itens:
                todos_seriais.extend(it.get("Seriais", []))
            if todos_seriais:
                serial_callback(todos_seriais)

        # ETAPA 2: MIGO Z62 (Devolução/Estorno das ordens antigas)
        etapa_atual = 2
        if progresso_callback:
            progresso_callback(2, 0.50, "Etapa 2/4 (50%): Estorno de Saldos das Ordens Antigas (MIGO Z62)...")
        log_callback("\n--- [2/4] ESTORNO DE SALDOS DAS ORDENS ANTIGAS (MIGO Z62) ---\n")
        _exibir_resumo_ordens(itens, log_callback, fase="MIGO Z62")
        _abrir_migo(sessao, "Z62")
        for idx, item in enumerate(itens):
            _preencher_item_migo(
                sessao=sessao,
                item=item,
                ordem_contabil=item["Ordem Origem"],
                movimento="Z62",
                indice=idx,
                total_itens=len(itens),
                log_callback=log_callback
            )
            if idx + 1 < len(itens):
                _avancar_proximo_item(sessao)

        doc_z62 = _verificar_e_gravar(sessao, "Z62", log_callback)

        # ETAPA 3: Localizar Ordem Destino na IW39
        etapa_atual = 3
        if progresso_callback:
            progresso_callback(3, 0.75, "Etapa 3/4 (75%): Consulta da Ordem Ativa Destino (IW39)...")
        log_callback(f"\n--- [3/4] CONSULTA DA ORDEM ATIVA DO RECURSO NOVO (IW39: {recurso_novo}) ---\n")
        ordem_destino = buscar_ordem_destino_iw39(sessao, recurso_novo, log_callback)

        # ETAPA 4: MIGO Z61 (Saída para a ordem do novo recurso)
        etapa_atual = 4
        if progresso_callback:
            progresso_callback(4, 0.90, "Etapa 4/4 (90%): Saída para Ordem Destino (MIGO Z61)...")
        log_callback(f"\n--- [4/4] SAÍDA PARA A ORDEM DESTINO {ordem_destino} (MIGO Z61) ---\n")
        _abrir_migo(sessao, "Z61")
        for idx, item in enumerate(itens):
            _preencher_item_migo(
                sessao=sessao,
                item=item,
                ordem_contabil=ordem_destino,
                movimento="Z61",
                indice=idx,
                total_itens=len(itens),
                log_callback=log_callback
            )
            if idx + 1 < len(itens):
                _avancar_proximo_item(sessao)

        doc_z61 = _verificar_e_gravar(sessao, "Z61", log_callback)

        # ETAPA 5: Salvar Histórico
        caminho_hist = os.path.join(_pasta_assistente(), "Historico_Transferencias_SAP.csv")
        campos = [
            "Data Hora", "Tipo Movimento", "Recurso Origem", "Recurso Destino",
            "Ordem Destino", "Doc SAP Z62", "Doc SAP Z61", "Total Itens", "Detalhes"
        ]
        detalhes_str = " | ".join(f"{it['Material']} (Qtd: {it['Quantidade']})" for it in itens)
        registro = {
            "Data Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Tipo Movimento": "MUDANÇA_DE_RECURSO",
            "Recurso Origem": recurso_antigo,
            "Recurso Destino": recurso_novo,
            "Ordem Destino": ordem_destino,
            "Doc SAP Z62": doc_z62,
            "Doc SAP Z61": doc_z61,
            "Total Itens": len(itens),
            "Detalhes": detalhes_str,
        }

        try:
            existe = os.path.isfile(caminho_hist) and os.path.getsize(caminho_hist) > 0
            with open(caminho_hist, "a", newline="", encoding="utf-8-sig") as f_csv:
                writer = csv.DictWriter(f_csv, fieldnames=campos, delimiter=";")
                if not existe:
                    writer.writeheader()
                writer.writerow(registro)
            log_callback(f"📁 Histórico de transferência atualizado em: {caminho_hist}\n")
        except Exception as e:
            log_callback(f"⚠️ Aviso ao gravar histórico CSV: {e}\n")

        log_callback(
            f"\n=======================================================\n"
            f"🏆 MUDANÇA DE RECURSO CONCLUÍDA COM ÊXITO!\n"
            f"Origem: {recurso_antigo} ➡️ Destino: {recurso_novo}\n"
            f"Documento Z62 (Estorno Origem): {doc_z62}\n"
            f"Ordem Destino (IW39): {ordem_destino}\n"
            f"Documento Z61 (Saída Destino): {doc_z61}\n"
            f"=======================================================\n"
        )

        if progresso_callback:
            progresso_callback(4, 1.00, "Concluído (100%): Mudança de Recurso realizada com sucesso!")

        return {
            "status": "sucesso",
            "recurso_antigo": recurso_antigo,
            "recurso_novo": recurso_novo,
            "doc_z62": doc_z62,
            "ordem_destino": ordem_destino,
            "doc_z61": doc_z61,
            "itens": itens,
        }

    except Exception as e:
        msg_erro = str(e)
        nomes_etapas = {
            1: "Consulta ZUDWM_237 e Auditoria ZMMCSE001",
            2: "Estorno de Saldos das Ordens Antigas (MIGO Z62)",
            3: "Consulta da Ordem Ativa Destino (IW39)",
            4: "Transferência para a Ordem Nova (MIGO Z61)",
        }
        nome_etapa = nomes_etapas.get(etapa_atual, "Execução SAP")
        bloco_erro = (
            f"\n=======================================================\n"
            f"❌ FALHA NA EXECUÇÃO (ETAPA {etapa_atual}/4: {nome_etapa.upper()})\n"
            f"• Mensagem/Motivo: {msg_erro}\n"
            f"• Recurso Origem: {recurso_antigo} | Recurso Destino: {recurso_novo}\n"
            f"• Horário: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
            f"• Ação recomendada: Verifique os detalhes das mensagens do SAP acima no terminal.\n"
            f"=======================================================\n"
        )
        log_callback(bloco_erro)
        raise RuntimeError(f"[Etapa {etapa_atual}/4 - {nome_etapa}] {msg_erro}")

    finally:
        pythoncom.CoUninitialize()


def executar(recurso_antigo, log_callback, serial_callback=None, progresso_callback=None, **kwargs):
    """Ponto de entrada compatível com a interface gráfica."""
    recurso_novo = kwargs.get("recurso_novo") or kwargs.get("equipe_destino") or ""
    return executar_mudanca_recurso(
        recurso_antigo=recurso_antigo,
        recurso_novo=recurso_novo,
        log_callback=log_callback,
        serial_callback=serial_callback,
        progresso_callback=progresso_callback
    )
