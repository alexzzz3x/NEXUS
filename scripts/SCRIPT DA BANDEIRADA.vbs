' Reset automatico inserido pelo Conecta Hub
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

session.findById("wnd[0]").maximize

' Fechar eventuais janelas/popups residuais
On Error Resume Next
Do While session.Children.Count > 1
    session.findById("wnd[1]/tbar[0]/btn[12]").press
    If Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    WScript.Sleep 200
Loop
On Error GoTo 0

Dim recurso
recurso = "csc354"

' 1. Acessa transacao IW39 (ou /nZUDWM_237) para pesquisar a ordem da equipe
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 300
If Err.Number <> 0 Or session.findById("wnd[0]/usr/cmbDY_PARVW") Is Nothing Then
   Err.Clear
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUDWM_237"
   session.findById("wnd[0]").sendVKey 0
   WScript.Sleep 300
   If Err.Number <> 0 Or session.findById("wnd[0]/usr/cmbDY_PARVW") Is Nothing Then
      Err.Clear
      session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00017"
      session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00017"
      WScript.Sleep 300
   End If
End If
On Error GoTo 0

' Preenche os parametros da busca
session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = recurso
session.findById("wnd[0]/usr/ctxtDY_PARNR").setFocus
session.findById("wnd[0]/usr/ctxtDY_PARNR").caretPosition = Len(recurso)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 200

' Executa a busca (btn[8])
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 500

' =========================================================================
' CAPTURA DA ORDEM: Suporta lista de OSME (Grade) E Ordem Unica (Detalhe)
' =========================================================================
Dim ordemBandeirada
ordemBandeirada = ""

' CASO 1: Grade ALV presente (lista de OSME com multiplas ordens)
Dim gridBandeirada
Set gridBandeirada = Nothing
On Error Resume Next
Set gridBandeirada = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
On Error GoTo 0

If Not gridBandeirada Is Nothing Then
    On Error Resume Next
    ' Ordenar pela data mais recente descendente
    gridBandeirada.setCurrentCell -1, "GSTRP"
    gridBandeirada.selectColumn "GSTRP"
    If Err.Number <> 0 Then
        Err.Clear
        gridBandeirada.setCurrentCell -1, "GLTRP"
        gridBandeirada.selectColumn "GLTRP"
    End If
    session.findById("wnd[0]/tbar[1]/btn[40]").press
    WScript.Sleep 300

    ' Reobtem a referencia do grid para evitar ponteiro invalido apos ordenacao
    Set gridBandeirada = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    ordemBandeirada = Trim(gridBandeirada.GetCellValue(0, "AUFNR"))
    If Err.Number <> 0 Or ordemBandeirada = "" Then
        Err.Clear
        ordemBandeirada = Trim(gridBandeirada.GetCellValue(0, "ORDERID"))
    End If
    On Error GoTo 0

    If ordemBandeirada <> "" Then
        WScript.Echo "Ordem encontrada na lista (grade): " & ordemBandeirada
    End If

    ' Sai da grade
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    WScript.Sleep 200
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    WScript.Sleep 200
    On Error GoTo 0
End If

' CASO 2: Se nao pegou pela grade, tenta ler da tela de detalhe (ordem unica aberta direto)
If Trim(ordemBandeirada) = "" Then
    Dim txtOrdemDetalhe
    Set txtOrdemDetalhe = Nothing
    On Error Resume Next
    ' ID exato fornecido pelo usuario:
    Set txtOrdemDetalhe = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR")
    If txtOrdemDetalhe Is Nothing Then
        Set txtOrdemDetalhe = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1101/txtCAUFVD-AUFNR")
    End If
    If txtOrdemDetalhe Is Nothing Then
        Set txtOrdemDetalhe = session.findById("wnd[0]/usr/txtCAUFVD-AUFNR")
    End If
    On Error GoTo 0

    ' Se ainda nao encontrou pelo ID fixo, varre controles de wnd[0] procurando CAUFVD-AUFNR
    If txtOrdemDetalhe Is Nothing Then
        On Error Resume Next
        Dim childObj
        For Each childObj In session.findById("wnd[0]/usr").Children
            If InStr(childObj.Id, "CAUFVD-AUFNR") > 0 Then
                Set txtOrdemDetalhe = childObj
                Exit For
            End If
        Next
        On Error GoTo 0
    End If

    If Not txtOrdemDetalhe Is Nothing Then
        On Error Resume Next
        ordemBandeirada = Trim(txtOrdemDetalhe.text)
        On Error GoTo 0

        If ordemBandeirada <> "" Then
            WScript.Echo "Ordem encontrada na tela inicial de detalhe: " & ordemBandeirada
        End If

        ' Sai da tela de detalhe da ordem
        On Error Resume Next
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        WScript.Sleep 200
        session.findById("wnd[0]/tbar[0]/btn[15]").press
        WScript.Sleep 200
        On Error GoTo 0
    End If
End If

' CASO 3: Nenhuma ordem encontrada
If Trim(ordemBandeirada) = "" Then
    Dim msgStatus
    msgStatus = ""
    On Error Resume Next
    msgStatus = session.findById("wnd[0]/sbar").text
    On Error GoTo 0
    If Trim(msgStatus) <> "" Then
        WScript.Echo "Status SAP: " & msgStatus
    End If
    WScript.Echo "Nenhuma ordem encontrada para o recurso: " & recurso
    WScript.Quit 2
End If

WScript.Echo "Ordem mais recente encontrada: " & ordemBandeirada & ". Bandeirando..."

' =========================================================================
' 2. Acessa transacao ZUDWM_213 para bandeirar a ordem
' =========================================================================
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUDWM_213"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 400

If Err.Number <> 0 Or session.findById("wnd[0]/usr/txtP_AUFNR") Is Nothing Then
   Err.Clear
   session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00024"
   session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00024"
   WScript.Sleep 400
End If
If Err.Number <> 0 Or session.findById("wnd[0]/usr/txtP_AUFNR") Is Nothing Then
   Err.Clear
   session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00022"
   session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00022"
   WScript.Sleep 400
End If
On Error GoTo 0

Dim txtAufnrZudwm
Set txtAufnrZudwm = Nothing
On Error Resume Next
Set txtAufnrZudwm = session.findById("wnd[0]/usr/txtP_AUFNR")
On Error GoTo 0

If txtAufnrZudwm Is Nothing Then
    WScript.Echo "ERRO: Nao foi possivel acessar a transacao ZUDWM_213."
    WScript.Quit 3
End If

txtAufnrZudwm.text = ordemBandeirada
txtAufnrZudwm.caretPosition = Len(ordemBandeirada)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 200

session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 400

session.findById("wnd[0]/tbar[1]/btn[16]").press
WScript.Sleep 400

' Confirmar popup de bandeirada se existir
On Error Resume Next
If session.Children.Count > 1 Then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 300
Else
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 300
End If
On Error GoTo 0

' Fecha ou volta tela
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[15]").press
On Error GoTo 0

' Verificar status final do SAP
Dim msgFinal
msgFinal = ""
On Error Resume Next
msgFinal = session.findById("wnd[0]/sbar").text
On Error GoTo 0
If Trim(msgFinal) <> "" Then
    WScript.Echo "Status SAP: " & msgFinal
End If

WScript.Echo "Bandeirada concluida com sucesso para ordem: " & ordemBandeirada
