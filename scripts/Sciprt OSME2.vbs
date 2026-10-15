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

' Fecha modais residuais pendentes
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

' Navega para ZUDWM_237 via comando direto (/nZUDWM_237) com fallback para arvore Easy Access (F00023)
On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZUDWM_237"
session.findById("wnd[0]").sendVKey 0
If Err.Number <> 0 Or session.findById("wnd[0]/usr/ctxtS_EQUIPE-LOW") Is Nothing Then
   Err.Clear
   session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00023"
End If
On Error GoTo 0

session.findById("wnd[0]/usr/ctxtS_EQUIPE-LOW").text = "csc300"
session.findById("wnd[0]/usr/ctxtS_EQUIPE-LOW").caretPosition = 6
session.findById("wnd[0]").sendVKey 0

' Garante desmarcar 'Ordens encerradas' se houver
On Error Resume Next
session.findById("wnd[0]/usr/chkP_ORDENC").selected = False
On Error GoTo 0

' Executa busca (F8)
session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 800

' Trata erros ou ausência de dados na ZUDWM_237
Dim statusMsg, grid
On Error Resume Next
statusMsg = session.findById("wnd[0]/sbar").text
If InStr(statusMsg, "Nenhum") > 0 Or InStr(statusMsg, "não encontr") > 0 Then
    WScript.Echo "AVISO: " & statusMsg
    WScript.Quit 0
End If

Set grid = session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell")
If Err.Number <> 0 Or grid Is Nothing Then
    WScript.Echo "AVISO: Grade ALV não encontrada para esta equipe."
    WScript.Quit 0
End If
On Error GoTo 0

' Ordena decrescente por Ordem (AUFNR) para a OSME mais recente ficar no topo
On Error Resume Next
grid.setCurrentCell -1, "AUFNR"
grid.selectColumn "AUFNR"
grid.pressToolbarButton "&SORT_DSC"
WScript.Sleep 300
On Error GoTo 0

' Seleciona a primeira linha (mais recente) e dispara exportação &XXL
grid.setCurrentCell 0, "AUFNR"
grid.selectedRows = "0"
grid.contextMenu
grid.selectContextMenuItem "&XXL"
WScript.Sleep 500

' Confirma primeira janela de exportação (Formato da Planilha)
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 600
On Error GoTo 0

' Confirma segunda janela (Gravar / Substituir arquivo se já existir)
On Error Resume Next
session.findById("wnd[1]/tbar[0]/btn[11]").press
If Err.Number <> 0 Then
    Err.Clear
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End If
WScript.Sleep 500
On Error GoTo 0

