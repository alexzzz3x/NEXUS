# -*- coding: utf-8 -*-
"""
utils_backup_diario.py — Sistema de Backup Diário e Auditoria Local de Saídas OSME.

Salva um arquivo Excel (.xlsx) por dia na pasta Documentos do usuário:
Documentos/NEXUS_Backups/Saidas_OSME_AAAA-MM-DD.xlsx

Garante redundância 100% local, imune a bloqueios de rede, concorrência e perdas de dados.
"""

import os
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def normalizar_texto(valor):
    if valor is None:
        return ""
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", str(valor))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r"[^A-Z0-9]", "", sem_acento.strip().upper())


def obter_pasta_backup_diario():
    """Retorna o caminho da pasta local Documentos/NEXUS_Backups/."""
    pasta_docs = os.path.join(os.path.expanduser("~"), "Documents")
    if not os.path.exists(pasta_docs):
        # Fallback caso 'Documents' não exista diretamente
        pasta_docs = os.path.expandvars(r"%USERPROFILE%\Documents")
    pasta_backup = os.path.join(pasta_docs, "NEXUS_Backups")
    os.makedirs(pasta_backup, exist_ok=True)
    return pasta_backup


def obter_caminho_backup_diario(data_dt=None):
    """Retorna o caminho do arquivo .xlsx diário no formato: Saidas_OSME_AAAA-MM-DD.xlsx."""
    if data_dt is None:
        data_dt = datetime.now()
    elif isinstance(data_dt, str):
        # Tenta interpretar formatos comuns
        m = re.match(r"^(\d{1,2})[\.\/-](\d{1,2})[\.\/-](\d{4})", data_dt.strip())
        if m:
            d, m_mes, a = m.groups()
            data_dt = datetime(int(a), int(m_mes), int(d))
        else:
            data_dt = datetime.now()

    str_data = data_dt.strftime("%Y-%m-%d")
    return os.path.join(obter_pasta_backup_diario(), f"Saidas_OSME_{str_data}.xlsx")


def _aplicar_estilo_cabecalho(ws):
    """Estiliza o cabeçalho da planilha de backup diário."""
    cor_fundo = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fonte_cabecalho = Font(name="Segoe UI", size=10, bold=True, color="F8FAFC")
    alinhamento = Alignment(horizontal="center", vertical="center", wrap_text=True)
    borda_fina = Border(
        bottom=Side(style="medium", color="0284C7"),
        top=Side(style="thin", color="334155"),
        left=Side(style="thin", color="334155"),
        right=Side(style="thin", color="334155")
    )

    ws.row_dimensions[1].height = 28
    for col in range(1, 10):
        cell = ws.cell(row=1, column=col)
        cell.fill = cor_fundo
        cell.font = fonte_cabecalho
        cell.alignment = alinhamento
        cell.border = borda_fina


def _ajustar_largura_colunas(ws):
    """Ajusta automaticamente a largura das colunas com margem de folga."""
    larguras_minimas = {
        1: 22, # SERIAL
        2: 12, # MATERIAL
        3: 10, # DEPOSITO
        4: 12, # EQUIPE
        5: 14, # DATA LCTO
        6: 15, # DOC SAP
        7: 20, # HORA PROCESSAMENTO
        8: 18, # STATUS GRAVACAO
        9: 30  # PLANILHA MESTRE
    }
    for col_idx, min_w in larguras_minimas.items():
        letra = get_column_letter(col_idx)
        max_len = min_w
        for row in range(1, min(ws.max_row + 1, 100)):
            val = ws.cell(row=row, column=col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)) + 2)
        ws.column_dimensions[letra].width = max_len


def gravar_saida_backup_diario(equipe, doc_sap, data_lcto, itens, status_mestre="Gravado", caminho_mestre=None):
    """
    Grava os seriais e materiais processados no arquivo diário do usuário.
    Nunca interrompe o fluxo principal caso ocorra qualquer erro de disco.
    Retorna (sucesso, caminho_arquivo).
    """
    try:
        caminho_backup = obter_caminho_backup_diario()
        existe = os.path.exists(caminho_backup)

        if existe:
            wb = load_workbook(caminho_backup)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "SAIDAS_OSME"
            # Cria cabeçalhos
            cabecalhos = [
                "SERIAL",
                "MATERIAL",
                "DEPÓSITO",
                "EQUIPE",
                "DATA LÇTO",
                "DOC SAP",
                "HORA N.E.X.U.S",
                "STATUS MESTRE",
                "PLANILHA MESTRE"
            ]
            ws.append(cabecalhos)
            _aplicar_estilo_cabecalho(ws)

        agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        nome_mestre = os.path.basename(caminho_mestre) if caminho_mestre else "CONTROLE_SERIAIS.vBeta.xlsm"

        # Borda simples para linhas de dados
        borda_linha = Border(
            bottom=Side(style="thin", color="E2E8F0"),
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0")
        )
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        fonte_dados = Font(name="Segoe UI", size=9)

        for it in itens:
            serial_val = str(it.get("serial") or "").strip()
            # Se for material não serializado
            if it.get("eh_nao_serializado") or not serial_val:
                serial_val = f"[NÃO SERIALIZADO: {it.get('quantidade', 1)} {it.get('unidade', 'UN')}]"

            linha_dados = [
                serial_val,
                str(it.get("material") or "").strip(),
                str(it.get("deposito") or "CL03").strip(),
                str(equipe or "").strip().upper(),
                str(data_lcto or "").strip(),
                str(doc_sap or "").strip(),
                agora_str,
                str(status_mestre or "Gravado"),
                nome_mestre
            ]
            ws.append(linha_dados)
            r_idx = ws.max_row
            ws.row_dimensions[r_idx].height = 20

            for c_idx in range(1, 10):
                c = ws.cell(row=r_idx, column=c_idx)
                c.font = fonte_dados
                c.alignment = alinhamento_centro
                c.border = borda_linha

                # Destaque de cor no status
                if c_idx == 8:
                    if "pendente" in str(status_mestre).lower() or "fila" in str(status_mestre).lower():
                        c.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
                        c.font = Font(name="Segoe UI", size=9, bold=True, color="B45309")
                    else:
                        c.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                        c.font = Font(name="Segoe UI", size=9, bold=True, color="15803D")

        _ajustar_largura_colunas(ws)
        wb.save(caminho_backup)
        wb.close()
        return True, caminho_backup
    except Exception as e:
        print(f"Erro ao salvar backup diário: {e}")
        return False, str(e)


def abrir_backup_diario(data_dt=None):
    """Abre a planilha diária diretamente no Excel do Windows."""
    caminho = obter_caminho_backup_diario(data_dt)
    if not os.path.exists(caminho):
        # Cria a planilha do dia vazia com cabeçalho caso ainda não haja saídas
        wb = Workbook()
        ws = wb.active
        ws.title = "SAIDAS_OSME"
        ws.append([
            "SERIAL", "MATERIAL", "DEPÓSITO", "EQUIPE", "DATA LÇTO",
            "DOC SAP", "HORA N.E.X.U.S", "STATUS MESTRE", "PLANILHA MESTRE"
        ])
        _aplicar_estilo_cabecalho(ws)
        _ajustar_largura_colunas(ws)
        wb.save(caminho)
        wb.close()

    try:
        os.startfile(caminho)
        return True, caminho
    except Exception as e:
        return False, str(e)


def comparar_backup_com_mestre(caminho_mestre, data_dt=None):
    """
    Cruza as saídas registradas no backup diário com a aba FORNECIDOS da planilha mestre.
    Retorna um resumo completo com os itens confirmados e os que faltam registrar.
    """
    caminho_backup = obter_caminho_backup_diario(data_dt)
    if not os.path.exists(caminho_backup):
        return {
            "sucesso": False,
            "erro": f"Nenhum arquivo de backup encontrado para a data informada:\n{caminho_backup}"
        }

    if not caminho_mestre or not os.path.exists(caminho_mestre):
        return {
            "sucesso": False,
            "erro": f"Planilha mestre não encontrada ou inacessível:\n{caminho_mestre}"
        }

    try:
        # 1. Carrega itens do backup diário
        wb_bak = load_workbook(caminho_backup, read_only=True, data_only=True)
        ws_bak = wb_bak.active

        itens_backup = []
        for row in ws_bak.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            s_raw = str(row[0]).strip()
            # Ignora itens informativos não serializados na conferência por serial
            if s_raw.startswith("["):
                continue
            itens_backup.append({
                "serial": s_raw,
                "serial_norm": normalizar_texto(s_raw),
                "material": str(row[1] or "").strip(),
                "deposito": str(row[2] or "").strip(),
                "equipe": str(row[3] or "").strip(),
                "data_lcto": str(row[4] or "").strip(),
                "doc_sap": str(row[5] or "").strip(),
                "hora_nexus": str(row[6] or "").strip()
            })
        wb_bak.close()

        if not itens_backup:
            return {
                "sucesso": True,
                "total_backup": 0,
                "confirmados": 0,
                "faltantes": [],
                "caminho_backup": caminho_backup,
                "caminho_mestre": caminho_mestre,
                "mensagem": "Nenhuma saída de serial registrada no backup diário de hoje."
            }

        # 2. Carrega seriais já gravados na aba FORNECIDOS da planilha mestre
        wb_mestre = load_workbook(caminho_mestre, read_only=True, data_only=True)
        if "FORNECIDOS" not in wb_mestre.sheetnames:
            wb_mestre.close()
            return {
                "sucesso": False,
                "erro": "Aba 'FORNECIDOS' não encontrada na planilha mestre."
            }

        ws_forn = wb_mestre["FORNECIDOS"]
        seriais_mestre_set = set()
        for row in ws_forn.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                seriais_mestre_set.add(normalizar_texto(row[0]))
        wb_mestre.close()

        # 3. Cruzamento
        confirmados = []
        faltantes = []
        for it in itens_backup:
            if it["serial_norm"] in seriais_mestre_set:
                confirmados.append(it)
            else:
                faltantes.append(it)

        return {
            "sucesso": True,
            "total_backup": len(itens_backup),
            "confirmados": len(confirmados),
            "faltantes": faltantes,
            "caminho_backup": caminho_backup,
            "caminho_mestre": caminho_mestre
        }

    except Exception as e:
        return {
            "sucesso": False,
            "erro": f"Erro durante a comparação: {e}"
        }
