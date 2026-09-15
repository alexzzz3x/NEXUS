import subprocess
import os
import tempfile
import csv
import re
import uuid
import glob
import time
import threading
from datetime import datetime
from decimal import Decimal, InvalidOperation

import pythoncom
import win32com.client
import win32clipboard

from modulos import transferencia_osmes as consulta_transferencia

NOME = "Transferencia de OSMES"
DESCRICAO = "ZUDWM_237 - recursos por equipe"
MOVIMENTA_SAP = True

_CONSULTAS_ATUAIS = {}
_CONSULTAS_LOCK = threading.Lock()

# ========== ETAPA 1 ==========
# Preenche equipe, marca incluir ordens encerradas, executa F8

# ========== ETAPA 2 ==========
# Para cada ordem com material positivo:
#   abre ordem → le Fim Base → se vencida → abre seriais → copia

SCRIPT_VBS = '''
Dim fso, arquivo
Set fso = CreateObject("Scripting.FileSystemObject")
arquivo = "{ARQUIVO_SAIDA}"
If fso.FileExists(arquivo) Then fso.DeleteFile arquivo

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
WScript.Sleep 80
session.findById("wnd[0]").sendVKey 12
WScript.Sleep 80
session.findById("wnd[0]").sendVKey 12
WScript.Sleep 80
On Error Goto 0

session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUDWM_237"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 300

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
WScript.Sleep 300

Dim grid
Set grid = session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell")

grid.setCurrentCell -1, "AUFNR"
grid.selectColumn "AUFNR"
grid.pressToolbarButton "&SORT_DSC"
WScript.Sleep 80

grid.currentCellColumn = "AUFNR"

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

        ' So processa se QTD_DISP for numerico e positivo
        If qtd_disp <> "" And Not SaldoDisponivelPositivo(qtd_disp) Then
            WScript.Echo "  Linha " & (i + 1) & " ignorada: QTD_DISP=" & qtd_disp
            qtd_disp = ""
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
                WScript.Sleep 80
                abriuOrdem = True

                On Error Resume Next
                fimBaseText = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subTERM:SAPLCOIH:7300/ctxtCAUFVD-GLTRP").text
                If Err.Number <> 0 Then fimBaseText = ""
                On Error Goto 0

                session.findById("wnd[0]/tbar[0]/btn[3]").press
                WScript.Sleep 100
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
                    WScript.Echo "  DENTRO DO PRAZO (Fim Base=" & fimBaseText & ")"
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
            ElseIf IsDate(fimBaseText) Then
                WScript.Echo "  DENTRO DO PRAZO (Fim Base=" & fimBaseText & ")"
            End If
        Else
            WScript.Echo "  ERRO: nao foi possivel ler Fim Base"
        End If

        ' Processar linha atual da ordem (QTD_DISP)
        If ordemVencida Then
            If qtd_disp <> "" Then
                If IsNumeric(qtd_disp) Then
                    If LinhaPossuiSaldoPositivo(grid, i) Then
                        WScript.Echo "    Material QTD=" & qtd_disp
                        grid.currentCellRow = i
                        grid.currentCellColumn = "QTD_DISP"
                        grid.selectedRows = CStr(i)
                        grid.clickCurrentCell
                        WScript.Sleep 2000

                        Dim serialGrid, sRowCount, s, sernr, serialList
                        On Error Resume Next
                        Err.Clear
                        Set serialGrid = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
                        If Err.Number = 0 Then
                            sRowCount = serialGrid.RowCount - 1
                            serialList = ""
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
                        WScript.Sleep 80
                    End If
                End If
            End If
        End If

        End If

     ElseIf ordemAtual <> "" And ordemVencida Then
         ' === SUB-LINHA da ordem atual (AUFNR em branco) ===
         ' So processa se tiver quantidade positiva
         If qtd_disp <> "" Then
             If IsNumeric(qtd_disp) Then
                 If LinhaPossuiSaldoPositivo(grid, i) Then
                     WScript.Echo "  + Material QTD=" & qtd_disp & " (ordem " & ordemAtual & ")"
                     grid.currentCellRow = i
                     grid.currentCellColumn = "QTD_DISP"
                     grid.selectedRows = CStr(i)
                     grid.clickCurrentCell
                     WScript.Sleep 2000

                     Dim serialGrid2, sRowCount2, s2, sernr2, serialList2
                     On Error Resume Next
                     Err.Clear
                     Set serialGrid2 = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
                     If Err.Number = 0 Then
                         sRowCount2 = serialGrid2.RowCount - 1
                         serialList2 = ""
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
                     WScript.Sleep 80
                 End If
             End If
         End If

    End If
Next
'''


def _carregar_amostra_migo(equipe):
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    padrao = os.path.join(pasta, f"Transferencia_OSME_Vencidas_{equipe}_*.csv")
    arquivos = sorted(glob.glob(padrao), key=os.path.getmtime, reverse=True)
    if not arquivos:
        return []
    with open(arquivos[0], "r", newline="", encoding="utf-8-sig") as arquivo_csv:
        registros = list(csv.DictReader(arquivo_csv, delimiter=";"))
    if not registros:
        return []

    ordem = registros[0].get("Ordem Origem", "")
    amostra = []
    tipos = set()
    for registro in registros:
        if registro.get("Ordem Origem", "") != ordem:
            continue
        tipo = registro.get("Tipo", "")
        if tipo in tipos:
            continue
        item = {
            "Material": registro.get("Material", "").strip(),
            "Quantidade": registro.get("Quantidade", "").strip(),
            "Unidade": registro.get("Unidade", "").strip(),
            "Ordem Origem": ordem,
            "Seriais": registro.get("Seriais", "").strip(),
            "Tipo": tipo,
        }
        if item["Material"]:
            amostra.append(item)
            tipos.add(tipo)
        if len(tipos) == 2:
            break
    return amostra


def _sessao_sap_sap():
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
    raise RuntimeError("nenhuma sessao SAP ativa encontrada")


def _aguardar_sap_sap(sessao, identificadores, timeout=20):
    limite = time.time() + timeout
    while time.time() < limite:
        for identificador in identificadores:
            try:
                return sessao.findById(identificador)
            except Exception:
                pass
        time.sleep(0.25)
    raise RuntimeError("controle SAP nao encontrado: " + " / ".join(identificadores))


def _celula_item_sap(sessao, campo, linha):
    prefixos = [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/",
    ]
    for prefixo in prefixos:
        for tipo in ("ctxt", "txt", "cmb", "chk"):
            for coluna in range(0, 60):
                identificador = f"{prefixo}{tipo}GOITEM-{campo}[{coluna},{linha}]"
                try:
                    return sessao.findById(identificador)
                except Exception:
                    pass
    return None


def _definir_item_sap(sessao, item, linha, log_callback):
    material_ids = [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX",
    ]
    try:
        material = _aguardar_sap_sap(sessao, material_ids, timeout=2)
    except RuntimeError:
        detalhes = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/btnBUTTON_ITEMDETAIL",
        ]
        _aguardar_sap_sap(sessao, detalhes).press()
        material = _aguardar_sap_sap(sessao, material_ids)
    # Em algumas variantes o campo MAKTX só aceita entrada depois que a
    # própria aba Material é selecionada.  Selecionar explicitamente evita
    # que o valor fique apenas no controle visual sem ser aplicado ao item.
    try:
        aba_material = sessao.findById(
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/"
            "subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_MATERIAL"
        )
        aba_material.select()
    except Exception:
        pass
    material.text = ""
    material.text = item["Material"]
    try:
        material.caretPosition = len(item["Material"])
    except Exception:
        pass
    sessao.findById("wnd[0]").sendVKey(0)
    time.sleep(1.0)

    # Só prossegue para conta/serial depois que o SAP reconheceu o material.
    # Isso também produz um diagnóstico útil quando o código não existe ou
    # quando a confirmação do material não foi efetivada.
    celula_material = _celula_item_sap(sessao, "MATNR", linha)
    if celula_material is not None:
        try:
            reconhecido = str(celula_material.text).strip().lstrip("0")
            esperado = str(item["Material"]).strip().lstrip("0")
            if reconhecido and esperado and reconhecido != esperado:
                raise RuntimeError(
                    f"material {item['Material']} nao reconhecido no item {linha + 1} "
                    f"(SAP exibiu {celula_material.text})"
                )
        except RuntimeError:
            raise
        except Exception:
            pass

    # Para materiais serializados o MIGO exige que a linha esteja marcada
    # como "Item OK" no grid antes de habilitar a aba de números de série.
    # Em linhas novas o checkbox de detalhe ainda não está disponível.
    try:
        item_grid_ok = _celula_item_sap(sessao, "TAKE_IT", linha)
        if item_grid_ok is not None:
            item_grid_ok.selected = True
            sessao.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
    except Exception:
        pass

    # A quantidade exibida no grid do MIGO é ERFMG (MENGE existe apenas em
    # algumas variantes antigas).  Sem preencher ERFMG o item permanece com
    # 0,000 e o SAP não disponibiliza a aba de seriais.
    for campo, valor in (
        ("ERFMG", item["Quantidade"]),
        ("WERKS", "EMPR"),
        ("LGORT", "CL03"),
        ("BWART", "Z62"),
    ):
        controle = _celula_item_sap(sessao, campo, linha)
        if controle is None and campo == "ERFMG":
            controle = _celula_item_sap(sessao, "MENGE", linha)
        if controle is not None:
            try:
                controle.text = valor
            except Exception:
                pass

    conta = _aguardar_sap_sap(sessao, [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT",
    ])
    conta.select()
    ordem = _aguardar_sap_sap(sessao, [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-AUFNR",
    ])
    ordem.text = item["Ordem Origem"]
    operacao = sessao.findById(
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_ACCOUNT/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/ctxtCOBL-VORNR_AUF"
    )
    operacao.text = "0010"
    # Confirma a imputação da ordem/operação para o SAP propagar a atribuição
    # à linha. Sem este Enter, a linha serializada permanece sem AUFNR e o
    # campo de quantidade fica bloqueado, ocultando a aba Serial.
    sessao.findById("wnd[0]").sendVKey(0)
    time.sleep(0.8)

    # Depois da confirmação da ordem, alguns perfis liberam ERFMG somente
    # neste momento (especialmente materiais serializados).
    if item["Seriais"]:
        quantidade = _celula_item_sap(sessao, "ERFMG", linha)
        if quantidade is not None:
            try:
                quantidade.text = item["Quantidade"]
                sessao.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
            except Exception:
                pass

    if item["Seriais"]:
        serial_tab = _aguardar_sap_sap(sessao, [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL",
        ])
        serial_tab.select()
        seriais = [s.strip() for s in item["Seriais"].split(",") if s.strip()]
        campo_serial = sessao.findById(
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]"
        )
        campo_serial.text = seriais[0]

    item_ok = sessao.findById(
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE"
    )
    item_ok.selected = True
    try:
        sessao.findById("wnd[0]").sendVKey(0)
        time.sleep(0.4)
    except Exception:
        pass
    log_callback(
        f"MIGO Z62: item {linha + 1} material {item['Material']} preenchido.\n"
    )


def executar_migo_z62_sap(equipe, log_callback):
    itens = _carregar_amostra_migo(equipe)
    if not itens:
        log_callback(
            "AVISO: nenhuma lista Transferencia_OSME encontrada para montar a amostra MIGO.\n"
        )
        return
    pythoncom.CoInitialize()
    try:
        sessao = _sessao_sap_sap()
        sessao.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
        sessao.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        for indice, item in enumerate(itens):
            # O botão de número da linha é o que ativa os campos editáveis do
            # item recém-criado. Sem essa etapa ERFMG fica somente leitura e
            # permanece em 0,000 (a aba de seriais também não aparece).
            if indice > 0:
                linha_btn = _aguardar_sap_sap(sessao, [
                    f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/ssubSUB_ITEMLIST:SAPLMIGO:0200/"
                    f"tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0,{indice}]",
                    f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/"
                    f"tblSAPLMIGOTV_GOITEM/btnGOITEM-ZEILE[0,{indice}]",
                ])
                linha_btn.press()
                time.sleep(0.4)
            _definir_item_sap(sessao, item, indice, log_callback)
            if indice + 1 < len(itens):
                sessao.findById(
                    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM"
                ).press()
                time.sleep(0.5)
        sessao.findById("wnd[0]/tbar[1]/btn[7]").press()
        time.sleep(1)
        log_callback(
            f"MIGO Z62 concluido ate Verificar. SAP: {sessao.findById('wnd[0]/sbar').text}\n"
        )
    finally:
        pythoncom.CoUninitialize()


def _pasta_assistente():
    pasta = os.path.join(os.path.expanduser("~"), "Documents", "Assistente Almox")
    os.makedirs(pasta, exist_ok=True)
    return pasta


def _registrar_consulta_atual(equipe, caminho, caminho_conferencia):
    with _CONSULTAS_LOCK:
        _CONSULTAS_ATUAIS[equipe] = {
            "caminho": caminho,
            "conferencia": caminho_conferencia,
            "registrada_em": datetime.now(),
        }


def _limpar_consulta_atual(equipe):
    with _CONSULTAS_LOCK:
        _CONSULTAS_ATUAIS.pop(equipe, None)


def _obter_consulta_atual(equipe):
    with _CONSULTAS_LOCK:
        consulta = _CONSULTAS_ATUAIS.get(equipe)
    if not consulta:
        raise RuntimeError(
            "nenhuma consulta atual vinculada a esta equipe; clique Executar antes do fluxo MIGO"
        )
    caminho = consulta["caminho"]
    caminho_conferencia = consulta.get("conferencia")
    if not os.path.isfile(caminho):
        raise RuntimeError("o arquivo da consulta atual nao existe mais")
    if caminho_conferencia and not os.path.isfile(caminho_conferencia):
        raise RuntimeError("o arquivo de conferencia da consulta atual nao existe mais")
    return caminho, caminho_conferencia


def _decimal_quantidade(valor):
    texto = str(valor or "").strip().replace(" ", "")
    if not texto:
        raise ValueError("quantidade vazia")
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        quantidade = Decimal(texto)
    except InvalidOperation as erro:
        raise ValueError(f"quantidade invalida: {valor}") from erro
    if quantidade <= 0:
        raise ValueError(f"quantidade nao positiva: {valor}")
    return quantidade


def _quantidade_para_migo(valor):
    quantidade = _decimal_quantidade(valor)
    if quantidade == quantidade.to_integral_value():
        return str(int(quantidade))
    return format(quantidade, "f").replace(".", ",")


def _carregar_itens_consulta_atual(equipe):
    caminho, caminho_conferencia = _obter_consulta_atual(equipe)
    with open(caminho, "r", newline="", encoding="utf-8-sig") as arquivo_csv:
        registros = list(csv.DictReader(arquivo_csv, delimiter=";"))
    if not registros:
        raise RuntimeError("a consulta atual nao possui itens para transferencia")

    colunas = {
        "Equipe", "Ordem Destino", "Ordem Origem", "Material",
        "Quantidade", "Unidade", "Tipo", "Seriais",
    }
    if not colunas.issubset(registros[0]):
        faltantes = sorted(colunas.difference(registros[0]))
        raise RuntimeError("CSV da consulta incompleto: " + ", ".join(faltantes))

    conferencia = []
    if caminho_conferencia:
        with open(
            caminho_conferencia, "r", newline="", encoding="utf-8-sig"
        ) as arquivo_conferencia:
            conferencia = list(csv.DictReader(arquivo_conferencia, delimiter=";"))
    colunas_conferencia = {
        "Equipe", "Serial", "STS Sistema", "STS Usuario", "Apto",
    }
    if conferencia and not colunas_conferencia.issubset(conferencia[0]):
        faltantes = sorted(colunas_conferencia.difference(conferencia[0]))
        raise RuntimeError(
            "CSV de conferencia incompleto: " + ", ".join(faltantes)
        )
    seriais_aptos = set()
    for numero_linha, registro in enumerate(conferencia, start=2):
        equipe_conferencia = registro.get("Equipe", "").strip().upper()
        if equipe_conferencia != equipe:
            raise RuntimeError(
                f"conferencia linha {numero_linha}: recurso "
                f"{equipe_conferencia or '(vazio)'} diverge de {equipe}"
            )
        serial = registro.get("Serial", "").strip()
        sistema = registro.get("STS Sistema", "").strip().upper()
        usuario = registro.get("STS Usuario", "").strip().upper()
        apto = registro.get("Apto", "").strip().upper()
        if apto == "SIM" and sistema == "LIDI" and usuario == "DISP" and serial:
            seriais_aptos.add(serial)

    destinos = set()
    itens = []
    for numero_linha, registro in enumerate(registros, start=2):
        equipe_registro = registro.get("Equipe", "").strip().upper()
        if equipe_registro != equipe:
            raise RuntimeError(
                f"linha {numero_linha}: recurso {equipe_registro or '(vazio)'} "
                f"diverge da consulta {equipe}"
            )
        destino = registro.get("Ordem Destino", "").strip()
        origem = registro.get("Ordem Origem", "").strip()
        material = registro.get("Material", "").strip()
        tipo = registro.get("Tipo", "").strip().upper()
        seriais = [
            serial.strip()
            for serial in registro.get("Seriais", "").split(",")
            if serial.strip()
        ]
        quantidade = _decimal_quantidade(registro.get("Quantidade", ""))
        if not destino or not origem or not material:
            raise RuntimeError(f"linha {numero_linha}: ordem ou material vazio")
        if tipo == "SERIALIZADO":
            if not seriais:
                raise RuntimeError(f"linha {numero_linha}: material serializado sem serial")
            if quantidade != Decimal(len(seriais)):
                raise RuntimeError(
                    f"linha {numero_linha}: quantidade {quantidade} difere de "
                    f"{len(seriais)} serial(is) aprovados"
                )
            nao_aprovados = [serial for serial in seriais if serial not in seriais_aptos]
            if nao_aprovados:
                raise RuntimeError(
                    f"linha {numero_linha}: serial(is) sem aprovacao LIDI/DISP "
                    f"na conferencia atual: {', '.join(nao_aprovados)}"
                )
        elif tipo != "NAO_SERIALIZADO":
            raise RuntimeError(f"linha {numero_linha}: tipo de material invalido: {tipo}")

        destinos.add(destino)
        itens.append({
            "Equipe": equipe,
            "Ordem Destino": destino,
            "Ordem Origem": origem,
            "Material": material,
            "Quantidade": _quantidade_para_migo(registro.get("Quantidade", "")),
            "Unidade": registro.get("Unidade", "").strip(),
            "Tipo": tipo,
            "Seriais": seriais,
        })

    if len(destinos) != 1:
        raise RuntimeError(
            "a consulta deve possuir exatamente uma ordem valida de destino; encontradas: "
            + ", ".join(sorted(destinos))
        )
    return caminho, itens, next(iter(destinos))


def _pressionar_enter(sessao, espera=0.8):
    sessao.findById("wnd[0]").sendVKey(0)
    time.sleep(espera)


def _controle_opcional(sessao, identificador):
    try:
        return sessao.findById(identificador)
    except Exception:
        return None


def _abrir_migo_movimento(sessao, movimento):
    sessao.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
    _pressionar_enter(sessao, 1.5)
    _garantir_detalhe_aberto(sessao)
    acoes = [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{portador}/"
        "subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION"
        for portador in ("0003", "0002")
    ]
    referencias = [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{portador}/"
        "subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-REFDOC"
        for portador in ("0003", "0002")
    ]
    # Padrao validado no script-base: devolucao Z62 usa entrada A01;
    # saida Z61 para a ordem valida usa A07. Ambas usam referencia R10.
    _aguardar_sap_sap(sessao, acoes).Key = "A01" if movimento == "Z62" else "A07"
    time.sleep(0.5)
    _aguardar_sap_sap(sessao, referencias).Key = "R10"
    time.sleep(0.5)
    movimento_controle = _aguardar_sap_sap(sessao, [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{portador}/"
        "subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART"
        for portador in ("0003", "0002")
    ])
    movimento_controle.text = movimento
    _pressionar_enter(sessao, 1.0)
    data_contabilizacao = _controle_opcional(
        sessao,
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/"
        "subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/"
        "tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/"
        "ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/ctxtGOHEAD-BUDAT",
    )
    if data_contabilizacao is not None:
        data_contabilizacao.text = datetime.now().strftime("%d.%m.%Y")


def _ids_aba(nome):
    return [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{portador}/"
        "subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/"
        f"tabsTS_GOITEM/tabpOK_GOITEM_{nome}"
        for portador in ("0003", "0002")
    ]


def _garantir_detalhe_aberto(sessao):
    for identificador in _ids_aba("MATERIAL"):
        if _controle_opcional(sessao, identificador) is not None:
            return
    botoes_abrir = [
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/"
        "subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL",
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/"
        "subSUB_ITEMDETAIL:SAPLMIGO:0302/btnBUTTON_DETAIL",
    ]
    _aguardar_sap_sap(sessao, botoes_abrir, timeout=10).press()
    _aguardar_sap_sap(sessao, _ids_aba("MATERIAL"), timeout=15)


def _selecionar_aba(sessao, nome):
    _garantir_detalhe_aberto(sessao)
    identificadores = _ids_aba(nome)
    _aguardar_sap_sap(sessao, identificadores)
    for identificador in identificadores:
        controle = _controle_opcional(sessao, identificador)
        if controle is not None:
            controle.select()
            return identificador
    raise RuntimeError(f"aba MIGO nao encontrada: {nome}")


def _preparar_linha_item(sessao, item, linha, log_callback):
    _garantir_detalhe_aberto(sessao)


def _avancar_proximo_item(sessao, proxima_linha, proximo_material, log_callback):
    botoes_proximo = [
        f"wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:{portador}/"
        "subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/"
        "btnOK_NEXT_ITEM"
        for portador in ("0003", "0002")
    ]
    _aguardar_sap_sap(sessao, botoes_proximo).press()
    time.sleep(0.7)
    _garantir_detalhe_aberto(sessao)
    log_callback(
        f"Avancou para a linha {proxima_linha + 1}; "
        f"proximo material: {proximo_material}.\n"
    )


def _preencher_seriais(sessao, seriais):
    aba_serial = _selecionar_aba(sessao, "SERIAL")
    botao_pesquisa = _aguardar_sap_sap(sessao, [
        aba_serial + "/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_SEARCH"
    ])
    botao_pesquisa.press()
    botao_multiplos = _aguardar_sap_sap(sessao, [
        "wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH"
    ])
    botao_multiplos.press()

    texto_anterior = None
    win32clipboard.OpenClipboard()
    try:
        try:
            texto_anterior = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        except Exception:
            texto_anterior = None
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText("\r\n".join(seriais), win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()
    try:
        _aguardar_sap_sap(sessao, ["wnd[1]/tbar[0]/btn[24]"]).press()
        _aguardar_sap_sap(sessao, ["wnd[1]/tbar[0]/btn[8]"]).press()
    finally:
        if texto_anterior is not None:
            try:
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(
                    texto_anterior, win32clipboard.CF_UNICODETEXT
                )
            finally:
                win32clipboard.CloseClipboard()

    _aguardar_sap_sap(sessao, ["wnd[0]/tbar[1]/btn[8]"]).press()
    time.sleep(1.0)
    grid_resultado = _controle_opcional(sessao, "wnd[0]/usr/cntlGRID1/shellcont/shell")
    if grid_resultado is not None:
        grid_resultado.setCurrentCell(-1, "")
        grid_resultado.selectAll()
        _aguardar_sap_sap(sessao, ["wnd[0]/tbar[1]/btn[47]"]).press()
        time.sleep(0.8)


def _preencher_item_movimento(sessao, item, linha, movimento, ordem, log_callback):
    _preparar_linha_item(sessao, item, linha, log_callback)
    aba_material = _selecionar_aba(sessao, "MATERIAL")
    campo_material = _aguardar_sap_sap(sessao, [
        aba_material
        + "/ssubSUB_TS_GOITEM_MATERIAL:SAPLMIGO:0310/ctxtGOITEM-MAKTX"
    ])
    campo_material.text = item["Material"]
    _pressionar_enter(sessao, 0.8)

    aba_quantidades = _selecionar_aba(sessao, "QUANTITIES")
    campo_quantidade = _aguardar_sap_sap(sessao, [
        aba_quantidades
        + "/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG"
    ])
    campo_quantidade.text = item["Quantidade"]
    _pressionar_enter(sessao, 0.8)

    aba_destino = _selecionar_aba(sessao, "DESTINAT.")
    prefixo_destino = (
        aba_destino + "/ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/"
    )
    movimento_item = _controle_opcional(sessao, prefixo_destino + "ctxtGOITEM-BWART")
    if movimento_item is not None:
        movimento_item.text = movimento
        _pressionar_enter(sessao, 0.5)
    _aguardar_sap_sap(sessao, [prefixo_destino + "ctxtGOITEM-NAME1"]).text = "EMPR"
    _aguardar_sap_sap(sessao, [prefixo_destino + "ctxtGOITEM-LGOBE"]).text = "CL03"
    _aguardar_sap_sap(sessao, [prefixo_destino + "txtGOITEM-WEMPF"]).text = "EMPR 1CL03"
    _pressionar_enter(sessao, 0.8)

    aba_conta = _selecionar_aba(sessao, "ACCOUNT")
    prefixo_conta = (
        aba_conta
        + "/ssubSUB_TS_GOITEM_ACCOUNT:SAPLMIGO:0345/"
        "ssubSUB_ACCOUNTINGBLOCK:SAPLKACB:1003/"
    )
    _aguardar_sap_sap(sessao, [prefixo_conta + "ctxtCOBL-AUFNR"]).text = ordem
    _aguardar_sap_sap(sessao, [prefixo_conta + "ctxtCOBL-VORNR_AUF"]).text = "0010"
    _pressionar_enter(sessao, 0.8)

    if item["Seriais"]:
        _preencher_seriais(sessao, item["Seriais"])

    item_ok = _controle_opcional(
        sessao,
        "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/"
        "subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/"
        "subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE",
    )
    if item_ok is not None:
        item_ok.selected = True
        _pressionar_enter(sessao, 0.4)
    log_callback(
        f"{movimento}: item {linha + 1}/{item['Material']} preparado "
        f"na ordem {ordem}.\n"
    )


def _mensagem_popup(sessao):
    try:
        if sessao.Children.Count > 1:
            return str(sessao.findById("wnd[1]").Text).strip() or "janela SAP aberta"
    except Exception:
        pass
    return ""


def _verificar_documento(sessao, movimento):
    sessao.findById("wnd[0]/tbar[1]/btn[7]").press()
    time.sleep(1.5)
    popup = _mensagem_popup(sessao)
    status = str(sessao.findById("wnd[0]/sbar").Text).strip()
    tipo = str(sessao.findById("wnd[0]/sbar").MessageType).upper()
    if popup:
        raise RuntimeError(f"{movimento}: SAP abriu mensagem durante Verificar: {popup}")
    if tipo in {"E", "A"}:
        raise RuntimeError(f"{movimento}: erro na verificacao SAP: {status}")
    if "DOCUMENTO O.K." not in status.upper():
        raise RuntimeError(
            f"{movimento}: verificacao sem confirmacao 'Documento O.K.': {status or '(vazio)'}"
        )
    return status


def _registrar_documento_sap(sessao, movimento, timeout=45):
    sessao.findById("wnd[0]/tbar[1]/btn[23]").press()
    limite = time.time() + timeout
    ultima_mensagem = ""
    while time.time() < limite:
        popup = _mensagem_popup(sessao)
        if popup:
            raise RuntimeError(f"{movimento}: SAP abriu mensagem ao registrar: {popup}")
        barra = sessao.findById("wnd[0]/sbar")
        ultima_mensagem = str(barra.Text).strip()
        encontrado = re.search(
            r"Doc\.?\s*material\s+(\d+)\s+registrad",
            ultima_mensagem,
            flags=re.IGNORECASE,
        )
        if encontrado:
            return encontrado.group(1), ultima_mensagem
        if str(barra.MessageType).upper() in {"E", "A"}:
            raise RuntimeError(f"{movimento}: SAP recusou o registro: {ultima_mensagem}")
        time.sleep(0.5)
    raise RuntimeError(
        f"{movimento}: documento nao identificado apos registrar. SAP: "
        f"{ultima_mensagem or '(sem mensagem)'}"
    )


def _salvar_historico_documento(
    identificador, equipe, movimento, documento, itens, ordem_destino,
    mensagem, arquivo_consulta,
):
    caminho = os.path.join(_pasta_assistente(), "Historico_Transferencias_SAP.csv")
    ordens_origem = ", ".join(sorted({item["Ordem Origem"] for item in itens}))
    detalhes = " | ".join(
        f"{item['Material']} x {item['Quantidade']}"
        + (" [" + ", ".join(item["Seriais"]) + "]" if item["Seriais"] else "")
        for item in itens
    )
    campos = [
        "Id Execucao", "Data Hora", "Equipe", "Movimento",
        "Documento SAP", "Ordens Origem", "Ordem Destino",
        "Itens", "Mensagem SAP", "Arquivo Consulta",
    ]
    registro = {
        "Id Execucao": identificador,
        "Data Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Equipe": equipe,
        "Movimento": movimento,
        "Documento SAP": documento,
        "Ordens Origem": ordens_origem,
        "Ordem Destino": ordem_destino,
        "Itens": detalhes,
        "Mensagem SAP": mensagem,
        "Arquivo Consulta": arquivo_consulta,
    }

    def gravar(destino):
        existe = os.path.isfile(destino) and os.path.getsize(destino) > 0
        with open(destino, "a", newline="", encoding="utf-8-sig") as arquivo_csv:
            writer = csv.DictWriter(arquivo_csv, fieldnames=campos, delimiter=";")
            if not existe:
                writer.writeheader()
            writer.writerow(registro)

    try:
        gravar(caminho)
        return caminho
    except PermissionError:
        # O Excel bloqueia CSVs abertos. Preserva imediatamente o documento em
        # um arquivo da mesma execucao, sem interromper a Z62/Z61.
        caminho_pendente = os.path.join(
            _pasta_assistente(),
            f"Historico_Transferencias_SAP_Pendente_{identificador}.csv",
        )
        gravar(caminho_pendente)
        return caminho_pendente


def _executar_movimento_migo(sessao, movimento, itens, ordem_destino, log_callback):
    _abrir_migo_movimento(sessao, movimento)
    for indice, item in enumerate(itens):
        ordem = item["Ordem Origem"] if movimento == "Z62" else ordem_destino
        _preencher_item_movimento(
            sessao, item, indice, movimento, ordem, log_callback
        )
        if indice + 1 < len(itens):
            _avancar_proximo_item(
                sessao, indice + 1, itens[indice + 1]["Material"], log_callback
            )
    mensagem_verificacao = _verificar_documento(sessao, movimento)
    log_callback(f"{movimento}: {mensagem_verificacao}\n")
    return _registrar_documento_sap(sessao, movimento)


def executar_fluxo_transferencia_sap(equipe, log_callback):
    equipe = equipe.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{1,12}", equipe):
        raise RuntimeError("codigo de equipe invalido")
    arquivo_consulta, itens, ordem_destino = _carregar_itens_consulta_atual(equipe)
    identificador = uuid.uuid4().hex
    log_callback(
        f"Fluxo completo da consulta atual: {len(itens)} item(ns), "
        f"destino {ordem_destino}.\n"
    )
    pythoncom.CoInitialize()
    try:
        sessao = _sessao_sap_sap()
        documento_z62, mensagem_z62 = _executar_movimento_migo(
            sessao, "Z62", itens, ordem_destino, log_callback
        )
        historico = _salvar_historico_documento(
            identificador, equipe, "Z62", documento_z62, itens,
            ordem_destino, mensagem_z62, arquivo_consulta,
        )
        log_callback(f"Z62 registrada. Documento SAP: {documento_z62}\n")
        log_callback(f"Historico salvo em: {historico}\n")
        if "_Pendente_" in os.path.basename(historico):
            log_callback(
                "AVISO: historico principal bloqueado; documento preservado "
                "no arquivo pendente.\n"
            )

        documento_z61, mensagem_z61 = _executar_movimento_migo(
            sessao, "Z61", itens, ordem_destino, log_callback
        )
        historico_z61 = _salvar_historico_documento(
            identificador, equipe, "Z61", documento_z61, itens,
            ordem_destino, mensagem_z61, arquivo_consulta,
        )
        log_callback(f"Z61 registrada. Documento SAP: {documento_z61}\n")
        log_callback(f"Historico salvo em: {historico_z61}\n")
        if "_Pendente_" in os.path.basename(historico_z61):
            log_callback(
                "AVISO: historico principal bloqueado; documento preservado "
                "no arquivo pendente.\n"
            )
        log_callback(
            "FLUXO CONCLUIDO\n"
            f"Documento Z62: {documento_z62}\n"
            f"Documento Z61: {documento_z61}\n"
        )
        return {"Z62": documento_z62, "Z61": documento_z61}
    finally:
        pythoncom.CoUninitialize()


def executar_processo_completo(equipe, log_callback, serial_callback=None):
    equipe = equipe.strip().upper()
    log_callback("ETAPA 1/3 - Consulta e conferencia na ZUDWM_237\n")
    _executar_consulta(equipe, log_callback, serial_callback=serial_callback)
    # A consulta somente e registrada em memoria depois que a
    # lista e a conferencia da mesma execucao passam por todas as validacoes.
    _obter_consulta_atual(equipe)
    log_callback("ETAPA 2/3 - Estorno das ordens vencidas com Z62\n")
    resultado = executar_fluxo_transferencia_sap(equipe, log_callback)
    log_callback("ETAPA 3/3 - Saida para a ordem valida com Z61 concluida\n")
    return resultado


def _executar_consulta(equipe, log_callback, serial_callback=None):
    equipe = equipe.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{1,12}", equipe):
        log_callback("ERRO: codigo de equipe invalido. Use somente letras e numeros.\n")
        return
    pasta = _pasta_assistente()
    padrao = os.path.join(pasta, f"Transferencia_OSME_Vencidas_{equipe}_*.csv")
    padrao_conferencia = os.path.join(
        pasta, f"Conferencia_Seriais_{equipe}_*.csv"
    )
    arquivos_anteriores = set(glob.glob(padrao))
    conferencias_anteriores = set(glob.glob(padrao_conferencia))
    _limpar_consulta_atual(equipe)
    inicio = time.time()
    try:
        consulta_transferencia.executar(
            equipe, log_callback, serial_callback=serial_callback
        )
        arquivos_novos = [
            caminho for caminho in glob.glob(padrao)
            if caminho not in arquivos_anteriores
            and os.path.getmtime(caminho) >= inicio - 2
        ]
        conferencias_novas = [
            caminho for caminho in glob.glob(padrao_conferencia)
            if caminho not in conferencias_anteriores
            and os.path.getmtime(caminho) >= inicio - 2
        ]
        if not arquivos_novos:
            log_callback(
                "AVISO: a consulta nao gerou uma lista atual; "
                "o fluxo MIGO permanece bloqueado.\n"
            )
            return
        caminho_atual = max(arquivos_novos, key=os.path.getmtime)
        conferencia_atual = (
            max(conferencias_novas, key=os.path.getmtime)
            if conferencias_novas else None
        )
        # Carrega e valida antes de liberar o arquivo para o botao MIGO.
        _registrar_consulta_atual(equipe, caminho_atual, conferencia_atual)
        try:
            _, itens, destino = _carregar_itens_consulta_atual(equipe)
        except Exception:
            _limpar_consulta_atual(equipe)
            raise
        log_callback(
            f"Consulta atual vinculada ao recurso {equipe}: {len(itens)} item(ns), "
            f"ordem valida {destino}.\n"
        )
        log_callback("Fluxo completo Z62/Z61 liberado para esta consulta.\n")
    except Exception as e:
        _limpar_consulta_atual(equipe)
        log_callback(f"\nERRO na consulta: {e}\n")


def executar(equipe, log_callback, serial_callback=None):
    """Executa consulta atual, Z62, Z61 e persistencia dos documentos SAP."""
    return executar_processo_completo(
        equipe, log_callback, serial_callback=serial_callback
    )
