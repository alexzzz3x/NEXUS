import subprocess
import os
import tempfile
import csv
import re
import uuid
import time
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pythoncom
import win32com.client
import win32clipboard

NOME = "Transferencia de OSMES"
DESCRICAO = "ZUDWM_237 - recursos por equipe"

# ========== ETAPA 1 ==========
# Preenche equipe, marca incluir ordens encerradas, executa F8

# ========== ETAPA 2 ==========
# Para cada ordem com material positivo:
#   abre ordem → le Fim Base → se vencida → abre seriais → copia

SCRIPT_VBS = '''
Dim fso, arquivo, arquivoMateriais, arquivoDestino
Set fso = CreateObject("Scripting.FileSystemObject")
arquivo = "{ARQUIVO_SAIDA}"
arquivoMateriais = "{ARQUIVO_MATERIAIS}"
arquivoDestino = "{ARQUIVO_DESTINO}"
If fso.FileExists(arquivo) Then fso.DeleteFile arquivo
If fso.FileExists(arquivoMateriais) Then fso.DeleteFile arquivoMateriais
If fso.FileExists(arquivoDestino) Then fso.DeleteFile arquivoDestino

Function SaldoDisponivelPositivo(valor)
    Dim textoSaldo, numeroSaldo
    SaldoDisponivelPositivo = False
    textoSaldo = Trim(CStr(valor))
    If textoSaldo = "" Then Exit Function

    textoSaldo = Replace(textoSaldo, Chr(160), "")
    textoSaldo = Replace(textoSaldo, " ", "")

    ' SAP pode exibir negativos como -10, 10- ou (10).
    If Left(textoSaldo, 1) = "-" Then Exit Function
    If Right(textoSaldo, 1) = "-" Then Exit Function
    If Left(textoSaldo, 1) = "(" And Right(textoSaldo, 1) = ")" Then Exit Function

    On Error Resume Next
    Err.Clear
    numeroSaldo = CDbl(textoSaldo)
    If Err.Number = 0 Then
        If numeroSaldo > 0 Then SaldoDisponivelPositivo = True
    End If
    Err.Clear
    On Error Goto 0
End Function

Function LinhaPossuiSaldoPositivo(objGrid, numeroLinha)
    Dim saldoConfirmado
    LinhaPossuiSaldoPositivo = False
    saldoConfirmado = ""
    On Error Resume Next
    Err.Clear
    saldoConfirmado = objGrid.getCellValue(numeroLinha, "QTD_DISP")
    If Err.Number = 0 Then
        LinhaPossuiSaldoPositivo = SaldoDisponivelPositivo(saldoConfirmado)
    End If
    Err.Clear
    On Error Goto 0
End Function

Sub RegistrarOrdemDestino(numeroOrdem, dataFim)
    Dim arquivoDestinoObj
    If fso.FileExists(arquivoDestino) Then Exit Sub
    On Error Resume Next
    Set arquivoDestinoObj = fso.OpenTextFile(arquivoDestino, 2, True)
    arquivoDestinoObj.WriteLine numeroOrdem & "|" & dataFim
    arquivoDestinoObj.Close
    On Error Goto 0
    WScript.Echo "ORDEM DESTINO ATUAL: " & numeroOrdem & " (Fim Base=" & dataFim & ")"
End Sub

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

On Error Resume Next
session.findById("wnd[0]").sendVKey 3
WScript.Sleep 200
session.findById("wnd[0]").sendVKey 3
WScript.Sleep 200
session.findById("wnd[0]").sendVKey 3
WScript.Sleep 200
On Error Goto 0

session.findById("wnd[0]").maximize
On Error Resume Next
session.findById("wnd[0]").sendVKey 12
WScript.Sleep 500
session.findById("wnd[0]").sendVKey 12
WScript.Sleep 500
session.findById("wnd[0]").sendVKey 12
WScript.Sleep 500
On Error Goto 0

session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUDWM_237"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

' Aguardar a tela inicial da ZUDWM_237 ficar pronta
Dim campoEquipe, tentativa, statusText
Set campoEquipe = Nothing
For tentativa = 1 To 60
    On Error Resume Next
    Err.Clear
    Set campoEquipe = session.findById("wnd[0]/usr/ctxtS_EQUIPE-LOW")
    On Error Goto 0
    If Not (campoEquipe Is Nothing) Then Exit For
    WScript.Sleep 500
Next

If campoEquipe Is Nothing Then
    statusText = ""
    On Error Resume Next
    statusText = session.findById("wnd[0]/sbar").text
    On Error Goto 0
    WScript.Echo "ERRO: tela inicial da ZUDWM_237 nao encontrada. Status SAP: " & statusText
    WScript.Quit 2
End If

' O checkbox pode nao existir em todas as variantes da transacao
Dim chkOrdensEncerradas
Set chkOrdensEncerradas = Nothing
On Error Resume Next
Err.Clear
Set chkOrdensEncerradas = session.findById("wnd[0]/usr/chkP_ORDENC")
On Error Goto 0
If Not (chkOrdensEncerradas Is Nothing) Then
    chkOrdensEncerradas.selected = true
Else
    WScript.Echo "AVISO: opcao Incluir ordens encerradas nao encontrada; continuando."
End If

campoEquipe.text = "{EQUIPE}"
campoEquipe.caretPosition = Len("{EQUIPE}")
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2000

Dim grid
Set grid = session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell")

grid.setCurrentCell -1, "AUFNR"
grid.selectColumn "AUFNR"
grid.pressToolbarButton "&SORT_DSC"
WScript.Sleep 500

' A ordenacao pode recriar o controle ALV e invalidar a referencia COM.
' Readquirir a grade evita "Falha catastrofica" na primeira leitura.
Set grid = Nothing
Dim tentativaGrid
For tentativaGrid = 1 To 40
    On Error Resume Next
    Err.Clear
    Set grid = session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell")
    On Error Goto 0
    If Not (grid Is Nothing) Then Exit For
    WScript.Sleep 250
Next
If grid Is Nothing Then
    WScript.Echo "ERRO: grade de ordens nao ficou disponivel apos a ordenacao."
    WScript.Quit 3
End If

Dim rowCount, i, ordem_val, qtd_disp
rowCount = grid.RowCount - 1
WScript.Echo "Linhas encontradas: " & (rowCount + 1)

Dim todayDate
todayDate = Date()

Dim ordemAtual, ordemVencida, fimBaseText, fimBaseDate, abriuOrdem
ordemAtual = ""
ordemVencida = False
fimBaseText = ""

For i = 0 to rowCount
    On Error Resume Next
    ordem_val = grid.getCellValue(i, "AUFNR")
    If Err.Number <> 0 Then ordem_val = ""
    Err.Clear
    qtd_disp = grid.getCellValue(i, "QTD_DISP")
    If Err.Number <> 0 Then qtd_disp = ""
    On Error Goto 0

    If qtd_disp <> "" And Not SaldoDisponivelPositivo(qtd_disp) Then
        WScript.Echo "  Linha " & (i + 1) & " ignorada: QTD_DISP=" & qtd_disp
    End If

    If ordem_val <> "" Then
        ' === NOVA ORDEM ===
        WScript.Echo ">>> Ordem " & ordem_val & " (" & (i + 1) & "/" & (rowCount + 1) & ")"
        ordemAtual = ordem_val
        ordemVencida = False

        If Not SaldoDisponivelPositivo(qtd_disp) Then
            WScript.Echo "  IGNORADA: linha com saldo disponivel zero ou negativo"
        Else

        ' Tentar ler Fim Base da grid
        On Error Resume Next
        fimBaseText = grid.getCellValue(i, "GLTRP")
        If Err.Number <> 0 Then fimBaseText = ""
        On Error Goto 0

        ' Se nao tem na grid, abrir ordem (so UMA vez)
        abriuOrdem = False
        If fimBaseText = "" Then
            If LinhaPossuiSaldoPositivo(grid, i) Then
                grid.currentCellColumn = "AUFNR"
                grid.currentCellRow = i
                grid.selectedRows = CStr(i)
                grid.clickCurrentCell
                WScript.Sleep 1500
                abriuOrdem = True

                On Error Resume Next
                fimBaseText = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP").text
                If Err.Number <> 0 Then fimBaseText = ""
                On Error Goto 0

                session.findById("wnd[0]/tbar[0]/btn[3]").press
                WScript.Sleep 800
            Else
                WScript.Echo "  BLOQUEADA antes de abrir ordem: saldo nao positivo"
            End If
        End If

        ' Decidir se venceu (parsing manual DD.MM.YYYY)
        If fimBaseText <> "" And InStr(fimBaseText, ".") > 0 Then
            Dim d, m, y
            On Error Resume Next
            d = CInt(Mid(fimBaseText, 1, 2))
            m = CInt(Mid(fimBaseText, 4, 2))
            y = CInt(Mid(fimBaseText, 7, 4))
            If Err.Number = 0 And y >= 1900 Then
                fimBaseDate = DateSerial(y, m, d)
                If fimBaseDate < todayDate Then
                    ordemVencida = True
                End If
                If ordemVencida Then
                    WScript.Echo "  VENCIDA (Fim Base=" & fimBaseText & ")"
                Else
                    RegistrarOrdemDestino ordemAtual, fimBaseText
                End If
            Else
                WScript.Echo "  ERRO parsing data: " & fimBaseText
            End If
            On Error Goto 0
        ElseIf fimBaseText <> "" Then
            ' Tenta CDate como fallback
            On Error Resume Next
            fimBaseDate = CDate(fimBaseText)
            If Err.Number = 0 And fimBaseDate < todayDate Then
                ordemVencida = True
            End If
            On Error Goto 0
            If ordemVencida Then
                WScript.Echo "  VENCIDA (Fim Base=" & fimBaseText & ")"
            Else
                RegistrarOrdemDestino ordemAtual, fimBaseText
            End If
        Else
            WScript.Echo "  ERRO: nao foi possivel ler Fim Base"
        End If

        ' Processar linha atual da ordem (QTD_DISP)
        If ordemVencida And LinhaPossuiSaldoPositivo(grid, i) Then
            Dim material, descricaoMaterial, unidadeMedida
            material = ""
            descricaoMaterial = ""
            unidadeMedida = ""
            On Error Resume Next
            Err.Clear
            material = grid.getCellValue(i, "MATNR")
            Err.Clear
            descricaoMaterial = grid.getCellValue(i, "MAKTX")
            Err.Clear
            unidadeMedida = grid.getCellValue(i, "MEINS")
            On Error Goto 0

            WScript.Echo "    Material QTD=" & qtd_disp
            grid.currentCellRow = i
            grid.currentCellColumn = "QTD_DISP"
            grid.selectedRows = CStr(i)
            grid.clickCurrentCell
            WScript.Sleep 1500

            Dim serialGrid, sRowCount, s, sernr, serialList, tipoMaterial, fm
            serialList = ""
            On Error Resume Next
            Err.Clear
            Set serialGrid = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
            If Err.Number = 0 Then
                sRowCount = serialGrid.RowCount - 1
                For s = 0 to sRowCount
                    Err.Clear
                    sernr = serialGrid.getCellValue(s, "SERNR")
                    If Err.Number = 0 And sernr <> "" Then
                        If serialList <> "" Then serialList = serialList & ", "
                        serialList = serialList & sernr
                        Set f = fso.OpenTextFile(arquivo, 8, True)
                        f.WriteLine ordemAtual & "|" & fimBaseText & "|" & sernr
                        f.Close
                    End If
                Next
                session.findById("wnd[1]").close
                On Error Goto 0
                WScript.Echo "    SERIAIS (" & ordemAtual & "): " & serialList
            Else
                On Error Goto 0
                WScript.Echo "    Sem seriais (popup nao abriu)"
            End If

            tipoMaterial = "NAO_SERIALIZADO"
            If serialList <> "" Then tipoMaterial = "SERIALIZADO"
            descricaoMaterial = Replace(descricaoMaterial, "|", "/")
            descricaoMaterial = Replace(descricaoMaterial, vbCr, " ")
            descricaoMaterial = Replace(descricaoMaterial, vbLf, " ")
            On Error Resume Next
            Set fm = fso.OpenTextFile(arquivoMateriais, 8, True)
            fm.WriteLine ordemAtual & "|" & fimBaseText & "|" & material & "|" & descricaoMaterial & "|" & unidadeMedida & "|" & qtd_disp & "|" & tipoMaterial & "|" & serialList
            fm.Close
            On Error Goto 0
            WScript.Sleep 500
        End If

        End If

    ElseIf ordemAtual <> "" And ordemVencida Then
        ' === SUB-LINHA da ordem atual (AUFNR em branco) ===
        ' So processa se tiver quantidade positiva
        If LinhaPossuiSaldoPositivo(grid, i) Then
            Dim material2, descricaoMaterial2, unidadeMedida2
            material2 = ""
            descricaoMaterial2 = ""
            unidadeMedida2 = ""
            On Error Resume Next
            Err.Clear
            material2 = grid.getCellValue(i, "MATNR")
            Err.Clear
            descricaoMaterial2 = grid.getCellValue(i, "MAKTX")
            Err.Clear
            unidadeMedida2 = grid.getCellValue(i, "MEINS")
            On Error Goto 0

            WScript.Echo "  + Material QTD=" & qtd_disp & " (ordem " & ordemAtual & ")"
            grid.currentCellRow = i
            grid.currentCellColumn = "QTD_DISP"
            grid.selectedRows = CStr(i)
            grid.clickCurrentCell
            WScript.Sleep 1500

            Dim serialGrid2, sRowCount2, s2, sernr2, serialList2, tipoMaterial2, fm2
            serialList2 = ""
            On Error Resume Next
            Err.Clear
            Set serialGrid2 = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
            If Err.Number = 0 Then
                sRowCount2 = serialGrid2.RowCount - 1
                For s2 = 0 to sRowCount2
                    Err.Clear
                    sernr2 = serialGrid2.getCellValue(s2, "SERNR")
                    If Err.Number = 0 And sernr2 <> "" Then
                        If serialList2 <> "" Then serialList2 = serialList2 & ", "
                        serialList2 = serialList2 & sernr2
                        Set f = fso.OpenTextFile(arquivo, 8, True)
                        f.WriteLine ordemAtual & "|" & fimBaseText & "|" & sernr2
                        f.Close
                    End If
                Next
                session.findById("wnd[1]").close
                On Error Goto 0
                WScript.Echo "    SERIAIS (" & ordemAtual & "): " & serialList2
            Else
                On Error Goto 0
                WScript.Echo "    Sem seriais (popup nao abriu)"
            End If

            tipoMaterial2 = "NAO_SERIALIZADO"
            If serialList2 <> "" Then tipoMaterial2 = "SERIALIZADO"
            descricaoMaterial2 = Replace(descricaoMaterial2, "|", "/")
            descricaoMaterial2 = Replace(descricaoMaterial2, vbCr, " ")
            descricaoMaterial2 = Replace(descricaoMaterial2, vbLf, " ")
            On Error Resume Next
            Set fm2 = fso.OpenTextFile(arquivoMateriais, 8, True)
            fm2.WriteLine ordemAtual & "|" & fimBaseText & "|" & material2 & "|" & descricaoMaterial2 & "|" & unidadeMedida2 & "|" & qtd_disp & "|" & tipoMaterial2 & "|" & serialList2
            fm2.Close
            On Error Goto 0
            WScript.Sleep 500
        End If

    End If
Next
'''


def _salvar_csv(equipe, registros):
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    os.makedirs(pasta, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(pasta, f"Seriais_OSME_Vencidas_{equipe}_{timestamp}.csv")
    with open(caminho, "w", newline="", encoding="utf-8-sig") as arquivo_csv:
        writer = csv.DictWriter(
            arquivo_csv,
            fieldnames=["Equipe", "Ordem Destino", "Ordem Origem", "Fim Base", "Serial"],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(registros)
    return caminho


def _salvar_csv_materiais(equipe, registros):
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    os.makedirs(pasta, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(pasta, f"Materiais_OSME_Vencidas_{equipe}_{timestamp}.csv")
    campos = [
        "Equipe",
        "Ordem Destino",
        "Ordem Origem",
        "Fim Base",
        "Material",
        "Descricao",
        "Quantidade",
        "Unidade",
        "Tipo",
        "Seriais",
    ]
    with open(caminho, "w", newline="", encoding="utf-8-sig") as arquivo_csv:
        writer = csv.DictWriter(arquivo_csv, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(registros)
    return caminho


def _agrupar_para_transferencia(equipe, ordem_destino, materiais):
    agrupados = {}
    ordem_chaves = []
    for item in materiais:
        quantidade = _quantidade_sap_para_decimal(item["Quantidade"])
        if quantidade is None:
            continue
        chave = (
            item["Ordem Origem"],
            item["Material"],
            item["Unidade"],
        )
        if chave not in agrupados:
            agrupados[chave] = {
                "Equipe": equipe,
                "Ordem Destino": ordem_destino,
                "Ordem Origem": item["Ordem Origem"],
                "Material": item["Material"],
                "Descricao": item["Descricao"],
                "Quantidade Decimal": Decimal("0"),
                "Unidade": item["Unidade"],
                "Tipo": item["Tipo"],
                "Seriais Lista": [],
            }
            ordem_chaves.append(chave)
        grupo = agrupados[chave]
        grupo["Quantidade Decimal"] += quantidade
        for serial in item["Seriais"].split(","):
            serial = serial.strip()
            if serial and serial not in grupo["Seriais Lista"]:
                grupo["Seriais Lista"].append(serial)

    resultado = []
    for chave in ordem_chaves:
        grupo = agrupados[chave]
        resultado.append({
            "Equipe": grupo["Equipe"],
            "Ordem Destino": grupo["Ordem Destino"],
            "Ordem Origem": grupo["Ordem Origem"],
            "Material": grupo["Material"],
            "Descricao": grupo["Descricao"],
            "Quantidade": _formatar_quantidade_sap(grupo["Quantidade Decimal"]),
            "Unidade": grupo["Unidade"],
            "Tipo": grupo["Tipo"],
            "Seriais": ", ".join(grupo["Seriais Lista"]),
        })
    return resultado


def _salvar_csv_transferencia(equipe, registros):
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    os.makedirs(pasta, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(pasta, f"Transferencia_OSME_Vencidas_{equipe}_{timestamp}.csv")
    campos = [
        "Equipe",
        "Ordem Destino",
        "Ordem Origem",
        "Material",
        "Descricao",
        "Quantidade",
        "Unidade",
        "Tipo",
        "Seriais",
    ]
    with open(caminho, "w", newline="", encoding="utf-8-sig") as arquivo_csv:
        writer = csv.DictWriter(arquivo_csv, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(registros)
    return caminho


def _extrair_seriais_unicos(materiais):
    seriais = []
    vistos = set()
    for item in materiais:
        for serial in item["Seriais"].split(","):
            serial = serial.strip()
            if serial and serial not in vistos:
                vistos.add(serial)
                seriais.append(serial)
    return seriais


def _normalizar_status(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto).strip().upper()


def _normalizar_serial_chave(valor):
    """Iguala o serial sem zeros da OSME ao serial de 18 posicoes do SAP."""
    texto = str(valor or "").strip().upper()
    if texto.isdigit():
        return texto.lstrip("0") or "0"
    return texto


def _salvar_csv_conferencia(equipe, registros):
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    os.makedirs(pasta, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho = os.path.join(pasta, f"Conferencia_Seriais_{equipe}_{timestamp}.csv")
    campos = [
        "Equipe",
        "Serial",
        "Material",
        "Ordem Origem",
        "Recurso",
        "STS Sistema",
        "STS Usuario",
        "ST Mov Mercadoria",
        "Apto",
        "Erro",
    ]
    with open(caminho, "w", newline="", encoding="utf-8-sig") as arquivo_csv:
        writer = csv.DictWriter(arquivo_csv, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(registros)
    return caminho


def _obter_sessao_sap():
    sap_gui = win32com.client.GetObject("SAPGUI")
    engine = sap_gui.GetScriptingEngine
    for indice_conexao in range(engine.Children.Count):
        conexao = engine.Children(indice_conexao)
        for indice_sessao in range(conexao.Children.Count):
            sessao = conexao.Children(indice_sessao)
            try:
                if str(sessao.Info.Transaction).upper() != "SESSION_MANAGER":
                    return sessao
            except Exception:
                return sessao
    raise RuntimeError("nenhuma sessao SAP ativa foi encontrada")


def _aguardar_controle(sessao, identificador, timeout=30):
    limite = time.time() + timeout
    while time.time() < limite:
        try:
            return sessao.findById(identificador)
        except Exception:
            time.sleep(0.25)
    raise RuntimeError(f"controle SAP nao encontrado: {identificador}")


def _consultar_seriais_zmmcse004(equipe, seriais, log_callback):
    if not seriais:
        return []

    log_callback(
        f"Abrindo ZMMCSE001 > Analise de equipe para conferir {len(seriais)} serial(is)...\n"
    )
    registros = []
    pythoncom.CoInitialize()
    try:
        sessao = _obter_sessao_sap()
        # Fluxo correto: abrir o cockpit ZMMCSE001 e clicar em Analise de equipe.
        sessao.findById("wnd[0]/tbar[0]/okcd").text = "/nZMMCSE001"
        sessao.findById("wnd[0]").sendVKey(0)
        botao_analise = _aguardar_controle(
            sessao, "wnd[0]/usr/btnBT003", timeout=30
        )
        botao_analise.press()
        # A consulta manual foi realizada no modo "Por equipamento".
        base = "wnd[0]/usr/subSUBSCREEN_9000:ZUDWM_PP350_CTRL_ANALISE_EQUIP:9000/"
        radio_equipamento = _aguardar_controle(sessao, base + "radRD_EQUI")
        radio_equipamento.selected = True
        campo_serial = _aguardar_controle(sessao, base + "ctxtS_SERNR-LOW")

        # Abre a seleção múltipla, cola todos os seriais e transfere a lista
        # de uma só vez para a tela de seleção.
        sessao.findById(base + "btn%_S_SERNR_%_APP_%-VALU_PUSH").press()
        _aguardar_controle(sessao, "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA")
        texto_clipboard = "\r\n".join(seriais)
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(
                win32clipboard.CF_UNICODETEXT, texto_clipboard
            )
        finally:
            win32clipboard.CloseClipboard()
        sessao.findById("wnd[1]/tbar[0]/btn[24]").press()
        time.sleep(0.5)
        sessao.findById("wnd[1]/tbar[0]/btn[8]").press()
        _aguardar_controle(sessao, base + "ctxtS_SERNR-LOW")

        sessao.findById("wnd[0]").sendVKey(8)
        grid_id = "wnd[0]/usr/cntlCO_ITEM/shellcont/shell"
        grid = _aguardar_controle(sessao, grid_id)

        fim_espera = time.time() + 30
        while time.time() < fim_espera and int(grid.RowCount) < 1:
            time.sleep(0.25)

        por_serial = {}
        total_linhas = int(grid.RowCount)
        try:
            linhas_visiveis = max(1, int(grid.VisibleRowCount))
        except Exception:
            linhas_visiveis = 10

        # O grid SAP carrega blocos conforme a rolagem. Mover FirstVisibleRow
        # equivale a avancar com Page Down e materializa todas as linhas.
        for inicio in range(0, total_linhas, linhas_visiveis):
            try:
                grid.FirstVisibleRow = inicio
                time.sleep(0.1)
            except Exception:
                pass

            fim = min(inicio + linhas_visiveis, total_linhas)
            for linha in range(inicio, fim):
                serial_grid = str(grid.getCellValue(linha, "SERIAL") or "").strip()
                if not serial_grid:
                    continue
                por_serial[_normalizar_serial_chave(serial_grid)] = {
                    "Material": str(grid.getCellValue(linha, "MATNR") or "").strip(),
                    "Ordem Origem": str(grid.getCellValue(linha, "AUFNR") or "").strip(),
                    "Recurso": str(grid.getCellValue(linha, "PARNR") or "").strip(),
                    "STS Sistema": str(grid.getCellValue(linha, "STATUS") or "").strip(),
                    "STS Usuario": str(grid.getCellValue(linha, "STUSER") or "").strip(),
                    "ST Mov Mercadoria": str(grid.getCellValue(linha, "ST_MOV") or "").strip(),
                }

        try:
            grid.FirstVisibleRow = 0
        except Exception:
            pass

        for serial in seriais:
            dados = por_serial.get(_normalizar_serial_chave(serial))
            registro = {
                "Equipe": equipe,
                "Serial": serial,
                "Material": dados["Material"] if dados else "",
                "Ordem Origem": dados["Ordem Origem"] if dados else "",
                "Recurso": dados["Recurso"] if dados else "",
                "STS Sistema": dados["STS Sistema"] if dados else "",
                "STS Usuario": dados["STS Usuario"] if dados else "",
                "ST Mov Mercadoria": dados["ST Mov Mercadoria"] if dados else "",
                "Apto": "NAO",
                "Erro": "" if dados else "Serial nao localizado no grid",
            }
            if dados:
                registro["Apto"] = (
                    "SIM"
                    if _normalizar_status(registro["Recurso"]) == _normalizar_status(equipe)
                    and _normalizar_status(registro["STS Sistema"]) == "LIDI"
                    and _normalizar_status(registro["STS Usuario"]) == "DISP"
                    and _normalizar_status(registro["ST Mov Mercadoria"]) == "DISPONIVEL"
                    else "NAO"
                )
            registros.append(registro)
        localizados = sum(1 for item in registros if not item["Erro"])
        log_callback(
            f"Conferencia ZMMCSE004: {localizados}/{len(seriais)} serial(is) "
            "localizado(s) no grid.\n"
        )
    finally:
        pythoncom.CoUninitialize()
    return registros


def _filtrar_materiais_conferencia(materiais, conferencia):
    """Mantem apenas seriais aprovados e recalcula sua quantidade total."""
    seriais_validos = {
        item["Serial"] for item in conferencia if item.get("Apto") == "SIM"
    }
    filtrados = []
    for item in materiais:
        if not item["Seriais"]:
            filtrados.append(item)
            continue
        seriais = [
            serial.strip()
            for serial in item["Seriais"].split(",")
            if serial.strip() in seriais_validos
        ]
        if not seriais:
            continue
        item_filtrado = dict(item)
        item_filtrado["Seriais"] = ", ".join(seriais)
        item_filtrado["Quantidade"] = _formatar_quantidade_sap(Decimal(len(seriais)))
        filtrados.append(item_filtrado)
    return filtrados


def _quantidade_sap_para_decimal(valor):
    texto = str(valor).strip().replace("\xa0", "").replace(" ", "")
    if not texto or texto.startswith("-") or texto.endswith("-"):
        return None
    if texto.startswith("(") and texto.endswith(")"):
        return None
    # Formato SAP pt-BR: ponto para milhar e virgula para decimal.
    texto = texto.replace(".", "").replace(",", ".")
    try:
        numero = Decimal(texto)
    except InvalidOperation:
        return None
    return numero if numero > 0 else None


def _formatar_quantidade_sap(valor):
    texto = f"{valor:,.3f}"
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def executar(equipe, log_callback, serial_callback=None):
    equipe = equipe.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{1,12}", equipe):
        log_callback("ERRO: codigo de equipe invalido. Use somente letras e numeros.\n")
        return

    identificador = uuid.uuid4().hex
    serials_path = os.path.join(tempfile.gettempdir(), f"SERIAIS_COLETADOS_{identificador}.txt")
    materiais_path = os.path.join(tempfile.gettempdir(), f"MATERIAIS_COLETADOS_{identificador}.txt")
    destino_path = os.path.join(tempfile.gettempdir(), f"ORDEM_DESTINO_{identificador}.txt")
    vbs_path = os.path.join(tempfile.gettempdir(), f"TRANSFERENCIA_OSMES_{identificador}.vbs")
    vbs_content = (
        SCRIPT_VBS
        .replace("{EQUIPE}", equipe)
        .replace("{ARQUIVO_SAIDA}", serials_path)
        .replace("{ARQUIVO_MATERIAIS}", materiais_path)
        .replace("{ARQUIVO_DESTINO}", destino_path)
    )
    with open(vbs_path, "w", encoding="utf-8") as f:
        f.write(vbs_content)
    log_callback("Versao: filtro por linha com dupla validacao de saldo (v3)\n")
    log_callback(f"Executando Transferencia de OSMES para equipe: {equipe}\n")
    try:
        result = subprocess.run(
            ["cscript.exe", "//Nologo", vbs_path],
            capture_output=True, text=True, timeout=600,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.stdout:
            log_callback(result.stdout)
        if result.stderr:
            log_callback(f"STDERR: {result.stderr}\n")
        if result.returncode != 0:
            detalhe = (result.stderr or result.stdout or "sem detalhes").strip()
            raise RuntimeError(
                f"consulta ZUDWM_237 interrompida pelo VBScript "
                f"(codigo {result.returncode}): {detalhe}"
            )
        log_callback("Concluido!\n")

        ordem_destino = ""
        fim_base_destino = ""
        if os.path.exists(destino_path):
            with open(destino_path, "r", encoding="utf-8") as f:
                partes_destino = [parte.strip() for parte in f.readline().split("|", 1)]
            if len(partes_destino) == 2:
                ordem_destino, fim_base_destino = partes_destino
                log_callback(
                    f"Ordem destino atual: {ordem_destino} (Fim Base={fim_base_destino})\n"
                )
        if not ordem_destino:
            log_callback("AVISO: nenhuma ordem nao vencida foi identificada.\n")

        registros = []
        vistos = set()
        if os.path.exists(serials_path):
            with open(serials_path, "r", encoding="utf-8") as f:
                for line in f:
                    partes = [parte.strip() for parte in line.strip().split("|", 2)]
                    if len(partes) != 3 or not partes[0] or not partes[2]:
                        continue
                    ordem, fim_base, serial = partes
                    chave = (ordem, fim_base, serial)
                    if chave in vistos:
                        continue
                    vistos.add(chave)
                    registros.append({
                        "Equipe": equipe,
                        "Ordem Destino": ordem_destino,
                        "Ordem Origem": ordem,
                        "Fim Base": fim_base,
                        "Serial": serial,
                    })

        if registros:
            log_callback(f"Seriais unicos encontrados: {len(registros)}\n")
            try:
                csv_path = _salvar_csv(equipe, registros)
                log_callback(f"Lista CSV salva em: {csv_path}\n")
            except Exception as e:
                log_callback(f"AVISO: nao foi possivel salvar o CSV: {e}\n")

        else:
            log_callback("Nenhum serial encontrado em OSMEs vencidas.\n")

        materiais = []
        materiais_vistos = set()
        if os.path.exists(materiais_path):
            with open(materiais_path, "r", encoding="utf-8") as f:
                for line in f:
                    partes = [parte.strip() for parte in line.rstrip("\r\n").split("|", 7)]
                    if len(partes) != 8:
                        continue
                    ordem, fim_base, material, descricao, unidade, quantidade, tipo, seriais = partes
                    chave = (ordem, material, quantidade, seriais)
                    if not material or chave in materiais_vistos:
                        continue
                    materiais_vistos.add(chave)
                    materiais.append({
                        "Equipe": equipe,
                        "Ordem Destino": ordem_destino,
                        "Ordem Origem": ordem,
                        "Fim Base": fim_base,
                        "Material": material,
                        "Descricao": descricao,
                        "Quantidade": quantidade,
                        "Unidade": unidade,
                        "Tipo": tipo,
                        "Seriais": seriais,
                    })

        if materiais:
            log_callback(f"Materiais com saldo positivo em OSMEs vencidas: {len(materiais)}\n")
            try:
                materiais_csv = _salvar_csv_materiais(equipe, materiais)
                log_callback(f"Lista de materiais salva em: {materiais_csv}\n")
            except Exception as e:
                log_callback(f"AVISO: nao foi possivel salvar a lista de materiais: {e}\n")

            seriais_popup = _extrair_seriais_unicos(materiais)
            conferencia = []
            conferencia_concluida = not seriais_popup
            if seriais_popup:
                try:
                    conferencia = _consultar_seriais_zmmcse004(
                        equipe, seriais_popup, log_callback
                    )
                    conferencia_concluida = True
                    conferencia_csv = _salvar_csv_conferencia(equipe, conferencia)
                    aptos = sum(1 for item in conferencia if item["Apto"] == "SIM")
                    removidos = len(seriais_popup) - aptos
                    log_callback(
                        f"Conferencia concluida: {aptos}/{len(conferencia)} serial(is) apto(s).\n"
                    )
                    if removidos:
                        log_callback(
                            f"{removidos} serial(is) removido(s) da transferencia por status ou recurso divergente.\n"
                        )
                    log_callback(f"Conferencia CSV salva em: {conferencia_csv}\n")
                except Exception as e:
                    log_callback(f"AVISO: nao foi possivel concluir a conferencia ZMMCSE004: {e}\n")

            if conferencia_concluida:
                materiais_para_transferencia = _filtrar_materiais_conferencia(
                    materiais, conferencia
                ) if seriais_popup else materiais
            else:
                # Sem conferencia concluida, nao liberar nenhum serial para transferencia.
                materiais_para_transferencia = [
                    item for item in materiais if not item["Seriais"]
                ]

            transferencia = _agrupar_para_transferencia(
                equipe, ordem_destino, materiais_para_transferencia
            )
            if transferencia:
                try:
                    transferencia_csv = _salvar_csv_transferencia(equipe, transferencia)
                    log_callback(
                        f"Lista para transferencia salva em: {transferencia_csv}\n"
                    )
                except Exception as e:
                    log_callback(
                        f"AVISO: nao foi possivel salvar a lista para transferencia: {e}\n"
                    )

            if serial_callback:
                totais_nao_serializados = {}

                for item in materiais_para_transferencia:
                    if not item["Seriais"]:
                        quantidade = _quantidade_sap_para_decimal(item["Quantidade"])
                        if quantidade is None:
                            continue
                        chave = (item["Material"], item["Unidade"])
                        totais_nao_serializados[chave] = (
                            totais_nao_serializados.get(chave, Decimal("0")) + quantidade
                        )

                destino_exibicao = ordem_destino or "Nao identificada"
                linhas_resultado = [
                    f"Recurso: {equipe}",
                    f"Ordem destino atual: {destino_exibicao}",
                    "",
                    "Serializados:",
                ]
                if seriais_popup:
                    linhas_resultado.extend(seriais_popup)
                else:
                    linhas_resultado.append("Nenhum")

                linhas_resultado.extend(["", "Cabos e conectores:"])
                if totais_nao_serializados:
                    for (material, unidade), quantidade in totais_nao_serializados.items():
                        linhas_resultado.append(
                            f"{material}: {_formatar_quantidade_sap(quantidade)} {unidade}"
                        )
                else:
                    linhas_resultado.append("Nenhum")

                linhas_resultado.extend(["", "Detalhamento para transferencia:"])
                if transferencia:
                    ordem_atual = None
                    for item in transferencia:
                        if item["Ordem Origem"] != ordem_atual:
                            ordem_atual = item["Ordem Origem"]
                            linhas_resultado.append("")
                            linhas_resultado.append(f"Ordem origem: {ordem_atual}")
                        linha_material = (
                            f"{item['Material']}: {item['Quantidade']} {item['Unidade']}"
                        )
                        linhas_resultado.append(linha_material)
                        if item["Seriais"]:
                            linhas_resultado.append(f"  Seriais: {item['Seriais']}")
                else:
                    linhas_resultado.append("Nenhum")

                linhas_resultado.extend(["", "Conferencia ZMMCSE004:"])
                if conferencia:
                    aptos = [item for item in conferencia if item["Apto"] == "SIM"]
                    nao_aptos = [item for item in conferencia if item["Apto"] != "SIM"]
                    linhas_resultado.append(
                        f"Aptos: {len(aptos)}/{len(conferencia)}"
                    )
                    for item in nao_aptos:
                        linhas_resultado.append(
                            f"NAO APTO - {item['Serial']}: "
                            f"Recurso={item['Recurso']} | "
                            f"{item['STS Sistema']} / {item['STS Usuario']} / "
                            f"{item['ST Mov Mercadoria']}"
                        )
                elif seriais_popup:
                    linhas_resultado.append("Conferencia nao concluida")
                else:
                    linhas_resultado.append("Nenhum serial para conferir")

                serial_callback(linhas_resultado)
        else:
            log_callback("Nenhum material com saldo positivo encontrado em OSMEs vencidas.\n")
    except subprocess.TimeoutExpired:
        log_callback("\nTempo limite excedido (600s).\n")
    except FileNotFoundError:
        log_callback("\nERRO: cscript.exe nao encontrado.\n")
    except Exception as e:
        log_callback(f"\nERRO: {e}\n")
    finally:
        for caminho_temporario in (vbs_path, serials_path, materiais_path, destino_path):
            try:
                os.remove(caminho_temporario)
            except OSError:
                pass
