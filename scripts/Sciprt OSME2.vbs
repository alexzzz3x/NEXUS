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
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").setCurrentCell 7, "MAKTX"
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").selectedRows = "7"
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlCONTAINER_ORDEM/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
