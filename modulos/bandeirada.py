import subprocess
import os
import tempfile

NOME = "Bandeirada"
DESCRICAO = "IW39 → ZUDWM213"

SCRIPT_VBS = '''
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
recurso = "{EQUIPE}"

' 1. Acessa transacao IW39
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

session.findById("wnd[0]/usr/cmbDY_PARVW").key = "Z2"
session.findById("wnd[0]/usr/ctxtDY_PARNR").text = recurso
session.findById("wnd[0]/usr/ctxtDY_PARNR").setFocus
session.findById("wnd[0]/usr/ctxtDY_PARNR").caretPosition = Len(recurso)
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 200

session.findById("wnd[0]/tbar[1]/btn[8]").press
WScript.Sleep 500

' =========================================================================
' CAPTURA DA ORDEM: Suporta lista de OSME (Grade) E Ordem Unica (Detalhe)
' =========================================================================
Dim ordemBandeirada
ordemBandeirada = ""

' CASO 1: Grade ALV presente
Dim gridBandeirada
Set gridBandeirada = Nothing
On Error Resume Next
Set gridBandeirada = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
On Error GoTo 0

If Not gridBandeirada Is Nothing Then
    On Error Resume Next
    gridBandeirada.setCurrentCell -1, "GSTRP"
    gridBandeirada.selectColumn "GSTRP"
    If Err.Number <> 0 Then
        Err.Clear
        gridBandeirada.setCurrentCell -1, "GLTRP"
        gridBandeirada.selectColumn "GLTRP"
    End If
    session.findById("wnd[0]/tbar[1]/btn[40]").press
    WScript.Sleep 300

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

    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    WScript.Sleep 200
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    WScript.Sleep 200
    On Error GoTo 0
End If

' CASO 2: Ordem unica aberta diretamente em tela de detalhe
If Trim(ordemBandeirada) = "" Then
    Dim txtOrdemDetalhe
    Set txtOrdemDetalhe = Nothing
    On Error Resume Next
    Set txtOrdemDetalhe = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/txtCAUFVD-AUFNR")
    If txtOrdemDetalhe Is Nothing Then
        Set txtOrdemDetalhe = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1101/txtCAUFVD-AUFNR")
    End If
    If txtOrdemDetalhe Is Nothing Then
        Set txtOrdemDetalhe = session.findById("wnd[0]/usr/txtCAUFVD-AUFNR")
    End If
    On Error GoTo 0

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

On Error Resume Next
If session.Children.Count > 1 Then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 300
Else
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 300
End If
On Error GoTo 0

On Error Resume Next
session.findById("wnd[0]/tbar[0]/btn[15]").press
On Error GoTo 0

Dim msgFinal
msgFinal = ""
On Error Resume Next
msgFinal = session.findById("wnd[0]/sbar").text
On Error GoTo 0
If Trim(msgFinal) <> "" Then
    WScript.Echo "Status SAP: " & msgFinal
End If

WScript.Echo "Bandeirada concluida com sucesso para ordem: " & ordemBandeirada
'''


def executar(equipe, log_callback, **kwargs):
    vbs_content = SCRIPT_VBS.replace("{EQUIPE}", equipe)
    vbs_path = os.path.join(tempfile.gettempdir(), "BANDEIRADA_Exec.vbs")
    with open(vbs_path, "w", encoding="utf-8") as f:
        f.write(vbs_content)
    log_callback(f"Executando Bandeirada para equipe: {equipe}\n")
    try:
        result = subprocess.run(
            ["cscript.exe", "//Nologo", vbs_path],
            capture_output=True, text=True, timeout=120,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if result.stdout:
            log_callback(result.stdout)
        if result.returncode != 0:
            log_callback(f"AVISO: codigo de retorno {result.returncode}\n")
        if result.stderr:
            log_callback(f"STDERR: {result.stderr}\n")
        log_callback("Concluido!\n")
    except subprocess.TimeoutExpired:
        log_callback("\nTempo limite excedido (120s).\n")
    except FileNotFoundError:
        log_callback("\nERRO: cscript.exe nao encontrado.\n")
    except Exception as e:
        log_callback(f"\nERRO: {e}\n")
    finally:
        try:
            os.remove(vbs_path)
        except:
            pass
