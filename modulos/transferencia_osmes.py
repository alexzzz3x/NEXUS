import subprocess
import os
import io
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

SCRIPT_IW39_VBS = '''
Dim fso, arquivoSaida
Set fso = CreateObject("Scripting.FileSystemObject")
arquivoSaida = "{ARQUIVO_IW39_SAIDA}"
If fso.FileExists(arquivoSaida) Then fso.DeleteFile arquivoSaida

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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 500

session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = "{EQUIPE}"
session.findById("wnd[0]/usr/chkDY_OFN").selected = True
session.findById("wnd[0]/usr/chkDY_IAR").selected = True
session.findById("wnd[0]/usr/chkDY_MAB").selected = False
session.findById("wnd[0]/usr/chkDY_HIS").selected = False

session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 800

Dim maxPop
maxPop = 0
Do While session.Children.Count > 1 And maxPop < 8
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
    maxPop = maxPop + 1
Loop

Function FindValueControlByPattern(parentObj, patternStr)
    Set FindValueControlByPattern = Nothing
    If parentObj Is Nothing Then Exit Function
    On Error Resume Next
    Dim idx, c, res
    For idx = 0 To parentObj.Children.Count - 1
        Set c = parentObj.Children(CInt(idx))
        If c.Type <> "GuiLabel" And c.Type <> "GuiTitlebar" And InStr(UCase(c.Id), "/LBL") = 0 Then
            If InStr(UCase(c.Id), UCase(patternStr)) > 0 Then
                Set FindValueControlByPattern = c
                Exit Function
            End If
        End If
        If c.Children.Count > 0 Then
            Set res = FindValueControlByPattern(c, patternStr)
            If Not res Is Nothing Then
                Set FindValueControlByPattern = res
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function

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
    On Error GoTo 0
End Function

Dim outF
Set outF = fso.OpenTextFile(arquivoSaida, 2, True)

Dim ordemIndiv, vencIndiv, stIndiv, stUsrIndiv
ordemIndiv = ""
vencIndiv = ""
stIndiv = ""
stUsrIndiv = ""

On Error Resume Next
Dim wTitle, reOrd, mOrd
wTitle = session.ActiveWindow.Text
Set reOrd = New RegExp
reOrd.Pattern = "\\b(1000\\d{8}|\\d{10,12})\\b"
Set mOrd = reOrd.Execute(wTitle)
If mOrd.Count > 0 Then
    ordemIndiv = Trim(mOrd(0).Value)
End If
On Error GoTo 0

If ordemIndiv = "" Then
    On Error Resume Next
    Dim txtOrd
    Set txtOrd = FindValueControlByPattern(session.findById("wnd[0]/usr"), "CAUFVD-AUFNR")
    If Not txtOrd Is Nothing Then
        If IsNumeric(Trim(txtOrd.text)) And Len(Trim(txtOrd.text)) >= 8 Then
            ordemIndiv = Trim(txtOrd.text)
        End If
    End If
    On Error GoTo 0
End If

If ordemIndiv <> "" Then
    On Error Resume Next
    Dim txtVenc
    Set txtVenc = FindValueControlByPattern(session.findById("wnd[0]/usr"), "CAUFVD-GLTRP")
    If Not txtVenc Is Nothing Then
        vencIndiv = Trim(txtVenc.text)
    End If
    Dim txtSt
    Set txtSt = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:7110/txtCAUFVD-STTXT")
    If Not txtSt Is Nothing Then
        stIndiv = Trim(txtSt.text)
    End If
    Dim txtUsr
    Set txtUsr = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subHEADER:SAPLCOIH:7110/txtCAUFVD-ASTTX")
    If Not txtUsr Is Nothing Then
        stUsrIndiv = Trim(txtUsr.text)
    End If
    On Error GoTo 0
    outF.WriteLine ordemIndiv & "|" & vencIndiv & "|" & stIndiv & "|" & stUsrIndiv
    outF.Close
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Quit 0
End If

Dim grid
Set grid = FindFirstGridView(session.findById("wnd[0]"))
If Not grid Is Nothing Then
    Dim totalLinhas, r, cOrd, cVenc, cSt, cStUsr
    totalLinhas = 0
    On Error Resume Next
    totalLinhas = grid.RowCount
    On Error GoTo 0
    For r = 0 To totalLinhas - 1
        cOrd = ""
        cVenc = ""
        cSt = ""
        cStUsr = ""
        On Error Resume Next
        cOrd = Trim(grid.GetCellValue(r, "AUFNR"))
        If cOrd = "" Then cOrd = Trim(grid.GetCellValue(r, "ORDERID"))
        cVenc = Trim(grid.GetCellValue(r, "GLTRP"))
        If cVenc = "" Then cVenc = Trim(grid.GetCellValue(r, "GSTRP"))
        cSt = Trim(grid.GetCellValue(r, "STTXT"))
        cStUsr = Trim(grid.GetCellValue(r, "ASTTX"))
        On Error GoTo 0
        If cOrd <> "" Then
            outF.WriteLine cOrd & "|" & cVenc & "|" & cSt & "|" & cStUsr
        End If
    Next
    outF.Close
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey 0
    WScript.Quit 0
End If

outF.Close
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0
'''

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

    ' SAP pode exibir negativos como -10, 10- ou (10). Rejeita sumariamente qualquer negativo!
    If InStr(textoSaldo, "-") > 0 Then Exit Function
    If InStr(textoSaldo, "(") > 0 Or InStr(textoSaldo, ")") > 0 Then Exit Function

    ' Rejeita saldos zerados
    If textoSaldo = "0" Or textoSaldo = "0,0" Or textoSaldo = "0,00" Or textoSaldo = "0,000" Then Exit Function
    If textoSaldo = "0.0" Or textoSaldo = "0.00" Or textoSaldo = "0.000" Then Exit Function

    ' Trata separador de milhar e decimal pt-BR
    If InStr(textoSaldo, ",") > 0 And InStr(textoSaldo, ".") > 0 Then
        textoSaldo = Replace(textoSaldo, ".", "")
    End If

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

Function LinhaEstaEncerradaTecnicamente(objGrid, numeroLinha)
    LinhaEstaEncerradaTecnicamente = False
    On Error Resume Next
    Dim cols, c, colName, valStr
    Set cols = objGrid.ColumnOrder
    If Not (cols Is Nothing) Then
        For c = 0 To cols.Count - 1
            colName = cols(c)
            valStr = LCase(Trim(CStr(objGrid.getCellValue(numeroLinha, colName))))
            If InStr(valStr, "encerrad") > 0 Or InStr(valStr, "teco") > 0 Or InStr(valStr, "enct") > 0 Or InStr(valStr, "fechad") > 0 Then
                LinhaEstaEncerradaTecnicamente = True
                Exit Function
            End If
        Next
    End If

    Dim colunasPossiveis, col
    colunasPossiveis = Array("LAYOUT", "STATUS", "STTXT", "TXT30", "TXT04", "ASTTX", "ZZSTATUS", "STAT", "PHASEN", "TXT04_1", "SYSSTATUS", "USERSTATUS")
    For Each col In colunasPossiveis
        valStr = LCase(Trim(CStr(objGrid.getCellValue(numeroLinha, col))))
        If InStr(valStr, "encerrad") > 0 Or InStr(valStr, "teco") > 0 Or InStr(valStr, "enct") > 0 Or InStr(valStr, "fechad") > 0 Then
            LinhaEstaEncerradaTecnicamente = True
            Exit Function
        End If
    Next
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

' Pré-registra a Ordem Destino Vigente (descoberta na IW39)
Dim ordemDestinoPrevia, fimBaseDestinoPrevia
ordemDestinoPrevia = "{ORDEM_DESTINO_PREVIA}"
fimBaseDestinoPrevia = "{FIM_BASE_DESTINO_PREVIA}"
If ordemDestinoPrevia <> "" Then
    RegistrarOrdemDestino ordemDestinoPrevia, fimBaseDestinoPrevia
End If

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

Dim ordemOrigemAlvo
ordemOrigemAlvo = "{ORDEM_ORIGEM_ALVO}"
If ordemOrigemAlvo <> "" Then
    Dim camposPossiveis, idCampo, campoOrdemObj, preencheuOrdem
    preencheuOrdem = False
    camposPossiveis = Array( _
        "wnd[0]/usr/ctxtS_AUFNR-LOW", _
        "wnd[0]/usr/ctxtP_AUFNR", _
        "wnd[0]/usr/txtP_AUFNR", _
        "wnd[0]/usr/ctxtS_ORDEM-LOW", _
        "wnd[0]/usr/ctxtP_ORDEM", _
        "wnd[0]/usr/txtP_ORDEM", _
        "wnd[0]/usr/ctxtS_ORDR-LOW", _
        "wnd[0]/usr/ctxtS_AUFNR", _
        "wnd[0]/usr/txtS_AUFNR-LOW", _
        "wnd[0]/usr/ctxtAUFNR" _
    )
    For Each idCampo In camposPossiveis
        On Error Resume Next
        Err.Clear
        Set campoOrdemObj = session.findById(idCampo)
        If Err.Number = 0 And Not (campoOrdemObj Is Nothing) Then
            campoOrdemObj.text = ordemOrigemAlvo
            campoOrdemObj.caretPosition = Len(ordemOrigemAlvo)
            preencheuOrdem = True
            WScript.Echo "Campo de ordem de origem preenchido via " & idCampo & ": " & ordemOrigemAlvo
            On Error GoTo 0
            Exit For
        End If
        Err.Clear
        On Error GoTo 0
    Next

    If Not preencheuOrdem Then
        Dim usrObj, idxCtrl, cCtrl
        On Error Resume Next
        Set usrObj = session.findById("wnd[0]/usr")
        If Not (usrObj Is Nothing) Then
            For idxCtrl = 0 To usrObj.Children.Count - 1
                Set cCtrl = usrObj.Children(CInt(idxCtrl))
                If (cCtrl.Type = "GuiCTextField" Or cCtrl.Type = "GuiTextField") And InStr(UCase(cCtrl.Id), "EQUIPE") = 0 Then
                    If InStr(UCase(cCtrl.Id), "AUFNR") > 0 Or InStr(UCase(cCtrl.Id), "ORD") > 0 Then
                        cCtrl.text = ordemOrigemAlvo
                        cCtrl.caretPosition = Len(ordemOrigemAlvo)
                        preencheuOrdem = True
                        WScript.Echo "Campo de ordem de origem encontrado dinamicamente (" & cCtrl.Id & "): " & ordemOrigemAlvo
                        Exit For
                    End If
                End If
            Next
        End If
        On Error GoTo 0
    End If

    On Error Resume Next
    Dim campoHigh
    Set campoHigh = session.findById("wnd[0]/usr/ctxtS_AUFNR-HIGH")
    If Not (campoHigh Is Nothing) Then campoHigh.text = ""
    Set campoHigh = session.findById("wnd[0]/usr/ctxtS_ORDEM-HIGH")
    If Not (campoHigh Is Nothing) Then campoHigh.text = ""
    On Error GoTo 0

    session.findById("wnd[0]").sendVKey 0
End If

session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 2000

Dim grid
Set grid = session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell")

grid.setCurrentCell -1, "AUFNR"
grid.selectColumn "AUFNR"
grid.pressToolbarButton "&SORT_DSC"
WScript.Sleep 500

' Readquirir a grade evita erro na primeira leitura
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

Dim ordemAtual, ordemVencida, ordemEhDestino, fimBaseText
ordemAtual = ordemOrigemAlvo
If ordemAtual <> "" And ordemAtual <> ordemDestinoPrevia Then
    ordemVencida = True
    ordemEhDestino = False
    WScript.Echo "Ordem inicial definida como origem: " & ordemAtual
Else
    ordemVencida = False
    ordemEhDestino = False
End If
fimBaseText = ""

For i = 0 to rowCount
    On Error Resume Next
    ordem_val = Trim(CStr(grid.getCellValue(i, "AUFNR")))
    If Err.Number <> 0 Then ordem_val = ""
    If ordem_val = "" Then
        ordem_val = Trim(CStr(grid.getCellValue(i, "ORDEM")))
        If Err.Number <> 0 Then ordem_val = ""
    End If
    Err.Clear
    qtd_disp = Trim(CStr(grid.getCellValue(i, "QTD_DISP")))
    If Err.Number <> 0 Then qtd_disp = ""
    On Error Goto 0

    If ordem_val <> "" And ordem_val <> ordemAtual Then
        ' === NOVA ORDEM ===
        ordemAtual = ordem_val
        fimBaseText = ""
        WScript.Echo ">>> Ordem " & ordemAtual & " (linha " & (i + 1) & "/" & (rowCount + 1) & ")"

        ' Verifica se esta ordem é a nova ordem vigente de destino identificada na IW39
        If ordemDestinoPrevia <> "" And ordemAtual = ordemDestinoPrevia Then
            ordemEhDestino = True
            ordemVencida = False
            WScript.Echo "  ORDEM DESTINO VIGENTE ATIVA: materiais serao mantidos intactos."
        Else
            ordemEhDestino = False
            ordemVencida = True
            WScript.Echo "  ORDEM DE ORIGEM: processando materiais validos com saldo positivo."
        End If

        ' Tenta ler data de fim-base da grade se disponível
        On Error Resume Next
        fimBaseText = Trim(CStr(grid.getCellValue(i, "GLTRP")))
        If Err.Number <> 0 Then fimBaseText = ""
        On Error Goto 0
    End If

    ' Se pertencer à ordem de destino vigente: nunca altera nem coleta
    If ordemEhDestino Then
        ' Linha da nova ordem de destino: mantém intacta
    ElseIf ordemVencida Then
        ' Linha da ordem de origem anterior
        ' PULANDO LINHAS NEGATIVAS OU ZERADAS:
        ' Se não tiver saldo estritamente positivo, pula sumariamente sem clicar!
        If Not SaldoDisponivelPositivo(qtd_disp) Then
            WScript.Echo "    Linha " & (i + 1) & " IGNORADA: quantidade negativa ou zero (" & qtd_disp & ") mantida na ordem anterior."
        Else
            Dim material, descricaoMaterial, unidadeMedida
            material = ""
            descricaoMaterial = ""
            unidadeMedida = ""
            On Error Resume Next
            Err.Clear
            material = Trim(CStr(grid.getCellValue(i, "MATNR")))
            Err.Clear
            descricaoMaterial = Trim(CStr(grid.getCellValue(i, "MAKTX")))
            Err.Clear
            unidadeMedida = Trim(CStr(grid.getCellValue(i, "MEINS")))
            On Error Goto 0

            WScript.Echo "    Coletando linha " & (i + 1) & " - Material=" & material & " QTD=" & qtd_disp & " (" & ordemAtual & ")"
            grid.currentCellRow = i
            grid.currentCellColumn = "QTD_DISP"
            grid.selectedRows = CStr(i)
            grid.clickCurrentCell
            WScript.Sleep 1500

            Dim serialGrid, sRowCount, s, sernr, serialList, tipoMaterial, fm, f
            serialList = ""
            On Error Resume Next
            Err.Clear
            Set serialGrid = session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell")
            If Err.Number = 0 And Not (serialGrid Is Nothing) Then
                sRowCount = serialGrid.RowCount - 1
                For s = 0 to sRowCount
                    Err.Clear
                    sernr = Trim(CStr(serialGrid.getCellValue(s, "SERNR")))
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
                WScript.Echo "      SERIAIS: " & serialList
            Else
                On Error Goto 0
                WScript.Echo "      Sem seriais (item nao serializado ou popup nao abriu)"
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
            WScript.Sleep 400
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
                sts_sis = _normalizar_status(registro["STS Sistema"])
                sts_usr = _normalizar_status(registro["STS Usuario"])
                st_mov = _normalizar_status(registro["ST Mov Mercadoria"])
                rec = _normalizar_status(registro["Recurso"])
                eq_norm = _normalizar_status(equipe)

                registro["Apto"] = (
                    "SIM"
                    if (rec == eq_norm or not rec)
                    and "LIDI" in sts_sis
                    and "DISP" in sts_usr
                    and ("DISP" in st_mov or not st_mov or "DISPONIVEL" in st_mov)
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


def _ler_linhas_seguras(caminho):
    """Lê linhas de arquivo gerado pelo VBScript com fallback de encoding (utf-8, cp1252, latin-1)."""
    if not caminho or not os.path.exists(caminho):
        return []
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            with open(caminho, "r", encoding=enc) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    with open(caminho, "r", encoding="utf-8", errors="replace") as f:
        return f.readlines()


def _consultar_ordens_iw39(equipe, log_callback=None):
    """
    Consulta a transação IW39 no SAP para o parceiro Z2={equipe}.
    Identifica:
      - Ordem VIGENTE no período (data fim >= hoje, com maior data de vencimento e não encerrada).
      - Ordens VENCIDAS (data fim < hoje ou encerradas).
    Retorna: (ordem_vigente, fim_base_vigente, ordens_vencidas_dict)
    """
    identificador = uuid.uuid4().hex
    saida_iw39 = os.path.join(tempfile.gettempdir(), f"IW39_ORDENS_{identificador}.txt")
    vbs_iw39_path = os.path.join(tempfile.gettempdir(), f"IW39_CONSULTA_{identificador}.vbs")

    vbs_content = (
        SCRIPT_IW39_VBS
        .replace("{EQUIPE}", equipe)
        .replace("{ARQUIVO_IW39_SAIDA}", saida_iw39)
    )
    with open(vbs_iw39_path, "w", encoding="latin-1", errors="replace") as f:
        f.write(vbs_content)

    try:
        result = subprocess.run(
            ["cscript.exe", "//Nologo", vbs_iw39_path],
            capture_output=True, text=True, timeout=120,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.stdout and log_callback:
            for l in result.stdout.splitlines():
                if "ERRO" in l.upper() or "AVISO" in l.upper():
                    log_callback(f"{l}\n")
        if result.returncode != 0 and log_callback:
            log_callback(f"Aviso IW39 (cscript code {result.returncode})\n")

        linhas_saida = _ler_linhas_seguras(saida_iw39)
        hoje = datetime.now().date()
        candidatas_vigentes = []
        ordens_vencidas = {}

        for linha in linhas_saida:
            linha = linha.strip()
            if not linha:
                continue
            partes = [p.strip() for p in linha.split("|")]
            if len(partes) < 2:
                continue
            ordem = partes[0]
            venc_str = partes[1]
            st_sis = partes[2].upper() if len(partes) > 2 else ""
            st_usr = partes[3].upper() if len(partes) > 3 else ""

            if not ordem:
                continue

            dt_venc = None
            if venc_str and "." in venc_str:
                try:
                    dt_venc = datetime.strptime(venc_str, "%d.%m.%Y").date()
                except Exception:
                    dt_venc = None

            encerrada = any(x in st_sis for x in ["TECO", "ENCT", "FECH", "ENCERR"])

            if dt_venc and dt_venc >= hoje and not encerrada:
                candidatas_vigentes.append({
                    "ordem": ordem,
                    "venc": venc_str,
                    "dt_venc": dt_venc,
                    "st_sis": st_sis,
                    "st_usr": st_usr
                })
            else:
                ordens_vencidas[ordem] = {
                    "venc": venc_str,
                    "dt_venc": dt_venc,
                    "encerrada": encerrada
                }

        ordem_vigente = ""
        fim_vigente = ""
        if candidatas_vigentes:
            candidatas_vigentes.sort(key=lambda x: x["dt_venc"], reverse=True)
            ordem_vigente = candidatas_vigentes[0]["ordem"]
            fim_vigente = candidatas_vigentes[0]["venc"]

        return ordem_vigente, fim_vigente, ordens_vencidas
    except Exception as e:
        if log_callback:
            log_callback(f"Erro ao consultar IW39: {e}\n")
        return "", "", {}
    finally:
        for p in (saida_iw39, vbs_iw39_path):
            try:
                os.remove(p)
            except OSError:
                pass


def _buscar_ordem_destino_fallback(equipe):
    """Busca a OSME ativa na aba CADASTRO da planilha oficial CONTROLE_SERIAIS se não encontrada na ZUDWM_237."""
    equipe_norm = equipe.strip().upper()
    caminhos = [
        r"X:\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm",
        r"\\terra\conecta\arquivos\obras\Elpa - Almox CR Mauá\CONTROLES\CONTROLE_SERIAIS.vBeta.xlsm",
        os.path.join(os.path.expanduser("~"), "Desktop", "CONTROLE_SERIAIS.vBeta.xlsm"),
        os.path.join(os.path.expanduser("~"), "Documents", "CONTROLE_SERIAIS.vBeta.xlsm"),
    ]
    for c in caminhos:
        if os.path.exists(c):
            try:
                import openpyxl
                with open(c, "rb") as f_raw:
                    buf = io.BytesIO(f_raw.read())
                try:
                    wb = openpyxl.load_workbook(buf, read_only=True, data_only=True)
                    if "CADASTRO" in wb.sheetnames:
                        ws = wb["CADASTRO"]
                        # Coluna E (índice 4): Equipes, Coluna H (índice 7): Ordem, Coluna I (índice 8): Vencimento
                        for row in ws.iter_rows(values_only=True):
                            if len(row) > 7:
                                eq = str(row[4] or "").strip().upper()
                                if eq == equipe_norm:
                                    ordem = str(row[7] or "").strip()
                                    venc = str(row[8] or "").strip() if len(row) > 8 else ""
                                    if ordem and ordem.upper() != "NONE" and ordem != "—":
                                        return ordem, venc
                finally:
                    try:
                        wb.close()
                    except Exception:
                        pass
                    try:
                        buf.close()
                    except Exception:
                        pass
            except Exception:
                pass
    return "", ""


def executar(equipe, log_callback, serial_callback=None):
    equipe = equipe.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]{1,12}", equipe):
        log_callback("ERRO: codigo de equipe invalido. Use somente letras e numeros.\n")
        return

    # ETAPA 0: Consulta prévia e obrigatória na IW39 para identificar a nova OSME vigente
    log_callback(f"\n🔎 [IW39] Consultando OSME vigente no período para equipe {equipe}...\n")
    ordem_iw39, venc_iw39, vencidas_iw39 = _consultar_ordens_iw39(equipe, log_callback)

    ordem_destino = ""
    fim_base_destino = ""
    ordem_origem_alvo = ""

    if ordem_iw39:
        ordem_destino = ordem_iw39
        fim_base_destino = venc_iw39
        log_callback(
            f"🎯 [IW39] Nova OSME VIGENTE identificada no período (DESTINO): {ordem_destino} "
            f"(Vencimento: {fim_base_destino})\n"
        )
        if vencidas_iw39:
            venc_txt = ", ".join([f"{o} (venc: {d.get('venc', '-')})" for o, d in vencidas_iw39.items()])
            log_callback(f"⏳ [IW39] Ordens anteriores/vencidas detectadas: {venc_txt}\n")
            ordens_venc_ordenadas = sorted(
                vencidas_iw39.items(),
                key=lambda item: item[1].get("dt_venc") or datetime.min.date(),
                reverse=True
            )
            ordem_origem_alvo = ordens_venc_ordenadas[0][0]
            venc_origem = ordens_venc_ordenadas[0][1].get("venc", "-")
            log_callback(
                f"🎯 [IW39] OSME anterior de ORIGEM selecionada para abrir na ZUDWM_237: {ordem_origem_alvo} "
                f"(Vencimento: {venc_origem})\n"
            )
    else:
        log_callback(
            f"⚠️ [IW39] Nenhuma ordem com vencimento futuro localizada na IW39 para {equipe}.\n"
        )
        ordem_fb, venc_fb = _buscar_ordem_destino_fallback(equipe)
        if ordem_fb and ordem_fb not in vencidas_iw39:
            ordem_destino = ordem_fb
            fim_base_destino = venc_fb
            log_callback(
                f"Ordem destino ativa identificada na planilha CADASTRO: {ordem_destino} "
                f"(Vencimento={fim_base_destino})\n"
            )
        else:
            msg_erro = (
                f"❌ ERRO CRÍTICO: Não foi possível identificar uma nova OSME vigente para {equipe}.\n"
                f"A ordem encontrada seria idêntica à ordem anterior ou não há ordem vigente futura no SAP IW39.\n"
                f"Transferência cancelada por segurança para evitar movimentação na mesma OSME.\n"
            )
            log_callback(msg_erro)
            raise RuntimeError(msg_erro)

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
        .replace("{ORDEM_DESTINO_PREVIA}", ordem_destino)
        .replace("{FIM_BASE_DESTINO_PREVIA}", fim_base_destino)
        .replace("{ORDEM_ORIGEM_ALVO}", ordem_origem_alvo)
    )
    with open(vbs_path, "w", encoding="latin-1", errors="replace") as f:
        f.write(vbs_content)
    log_callback("Versao: filtro por linha com OSME vigente IW39 e rejeicao estrita de saldo negativo (v4)\n")
    log_callback(f"Executando varredura ZUDWM_237 para equipe: {equipe}\n")
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
        log_callback("Varredura ZUDWM_237 concluida!\n")

        # Confere ordem destino
        if os.path.exists(destino_path):
            linhas_destino = _ler_linhas_seguras(destino_path)
            if linhas_destino:
                partes_destino = [parte.strip() for parte in linhas_destino[0].split("|", 1)]
                if len(partes_destino) == 2 and partes_destino[0]:
                    ordem_destino, fim_base_destino = partes_destino

        registros = []
        vistos = set()
        if os.path.exists(serials_path):
            for line in _ler_linhas_seguras(serials_path):
                partes = [parte.strip() for parte in line.strip().split("|", 2)]
                if len(partes) != 3 or not partes[0] or not partes[2]:
                    continue
                ordem, fim_base, serial = partes
                # Não coleta seriais da própria ordem de destino
                if ordem == ordem_destino:
                    continue
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
            for line in _ler_linhas_seguras(materiais_path):
                partes = [parte.strip() for parte in line.rstrip("\r\n").split("|", 7)]
                if len(partes) != 8:
                    continue
                ordem, fim_base, material, descricao, unidade, quantidade, tipo, seriais = partes
                chave = (ordem, material, quantidade, seriais)
                if not material or chave in materiais_vistos:
                    continue

                # DUPLA VALIDAÇÃO DE SALDO: Pula sumariamente linhas negativas ou zero
                qtd_dec = _quantidade_sap_para_decimal(quantidade)
                if qtd_dec is None or qtd_dec <= 0:
                    continue

                # NUNCA transfere materiais da própria ordem de destino
                if ordem == ordem_destino:
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
                            f"⚠️ {removidos} serial(is) NÃO transferido(s) (mantidos na OSME antiga por status/recurso):\n"
                        )
                        for item in conferencia:
                            if item.get("Apto") != "SIM":
                                mot = []
                                rec = item.get("Recurso", "")
                                if rec and _normalizar_status(rec) != _normalizar_status(equipe):
                                    mot.append(f"recurso='{rec}'")
                                sis = item.get("STS Sistema", "")
                                if "LIDI" not in _normalizar_status(sis):
                                    mot.append(f"sistema='{sis}'")
                                usr = item.get("STS Usuario", "")
                                if "DISP" not in _normalizar_status(usr):
                                    mot.append(f"usuário='{usr}'")
                                stm = item.get("ST Mov Mercadoria", "")
                                if stm and "DISP" not in _normalizar_status(stm) and "DISPONIVEL" not in _normalizar_status(stm):
                                    mot.append(f"movimento='{stm}'")
                                motivo_txt = ", ".join(mot) if mot else (item.get("Erro") or "não disponível")
                                log_callback(f"   • Serial {item.get('Serial')}: {motivo_txt}\n")
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
            else:
                msg_aviso = (
                    f"⚠️ {equipe}: Nenhum material elegível para transferência restou após a conferência.\n"
                    "Todos os medidores encontrados estão com status diferente de LIDI DISP e foram mantidos na ordem anterior.\n"
                )
                log_callback(msg_aviso)
                raise RuntimeError(msg_aviso)

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
            msg_aviso = (
                f"ℹ️ {equipe}: Nenhum material com saldo positivo encontrado em OSMEs vencidas.\n"
                "Todas as linhas possuem saldo negativo ou zerado, ou já pertencem à ordem de destino.\n"
                "Nenhum material pendente de transferência.\n"
            )
            log_callback(msg_aviso)
            raise RuntimeError(msg_aviso)
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
